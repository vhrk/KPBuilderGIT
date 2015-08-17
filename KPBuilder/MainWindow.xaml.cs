using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Serialization;
using Ex = Microsoft.Office.Interop.Excel;
using Wd = Microsoft.Office.Interop.Word;
using System.Windows.Xps.Packaging;
using System.Threading;

namespace KPBuilder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string[] Templates;
        private Ex.Application ExcelApp;
        private Ex.Workbook ExcelWorkbook;
        private Wd.Application WordApp;
        private Wd.Document WordDoc;

       

        ExtraItem Strah = new ExtraItem() { Text="Профессиональная ответственность [ourcompany] застрахована в размере 1 000 000 долл. США на каждый страховой случай. Мы несем материальную ответственность за сохранность и исправность имущества клиента при проведении работ. В случае повреждения или утраты товарно-материальных ценностей по вине своих сотрудников [ourcompany] возмещает их стоимость в полном объеме." };

        public MainWindow()
        {
            InitializeComponent();
            Templates = Directory.GetFiles(Environment.CurrentDirectory+"\\KPTeml\\").Where(m => m.Contains("Шаблон")).ToArray();
            comboBox.ItemsSource = Templates;
            var ad= new AllData() { Our = LoadOurData() };
            DataContext = ad;
            dataGridExtra1.ItemsSource = ad.e1;
            dataGridExtraSod.ItemsSource = ad.e2;
            dataGridContacts.ItemsSource = ad.Contacts;
          
        }

        private void Wb1_Navigated(object sender, NavigationEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsm";
            dlg.Filter = "Excel Document (*.xls)|*.xls|Excel Document (*.xlsx)|*.xlsx|Excel Document (*.xlsm)|*.xlsm";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                textBoxExcelPath.Text = filename;
            }
        }
        static string SaveW(Wd.Application app, Wd.Document WData)
        {

            //string pth = Directory.GetParent(Directory.GetParent(TemplateFilePath).FullName) + "\\" + TempDir + "TMP" + DateTime.Now.Ticks + ".doc";
            string pth = Environment.CurrentDirectory + "\\Результат.doc";

            WData.SaveAs(pth);

            return pth;
        }


        OurData LoadOurData()
        {
            
            using (var ff = new FileStream("data.xml", FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if(ff.Length==0)
                {
                    return new OurData();
                }
                XmlSerializer xs = new XmlSerializer(typeof(OurData));
                var LsLoaded = (OurData)xs.Deserialize(ff);
                return LsLoaded;
            }
        }

        void SaveOurData(OurData od)
        {
            XmlSerializer s = new XmlSerializer(typeof(OurData));
            using (var ss = new FileStream("data.xml", FileMode.Create))
            {
                s.Serialize(ss, od);
            }
        }
        private XpsDocument ConvertWordToXps(string wordFilename, string xpsFilename)
        {
            // Create a WordApplication and host word document 
            Wd.Application wordApp = new Microsoft.Office.Interop.Word.Application() { Visible = false };
            try
            {
                wordApp.Documents.Open(wordFilename);

                // To Invisible the word document 
                wordApp.Application.Visible = false;


                // Minimize the opened word document 
                wordApp.WindowState = Wd.WdWindowState.wdWindowStateMinimize;


                Wd.Document doc = wordApp.ActiveDocument;


                doc.SaveAs(xpsFilename, Wd.WdSaveFormat.wdFormatXPS);


                XpsDocument xpsDocument = new XpsDocument(xpsFilename, FileAccess.Read);
                return xpsDocument;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurs, The error message is  " + ex.ToString());
                return null;
            }
            finally
            {
                wordApp.Documents.Close();
                ((Wd._Application)wordApp).Quit(Wd.WdSaveOptions.wdDoNotSaveChanges);
            }
        }

        static string GetCostText(double val, bool caps = false)
        {
            ExternalHlp.Валюта curr = ExternalHlp.Валюта.Рубли;

            string res = ExternalHlp.Сумма.Пропись(val, curr).ToString();

            if (res.Contains("США"))
            {
                res = res.Replace("США", "") + " США";
            }

            if (caps)
            {
                res = res.Substring(0, 1).ToUpper() + res.Substring(1, res.Length - 1);
            }

            return res;
        }

        string o2s(object o)
        {
            if (o == null)
                return "";

            return o.ToString();
        }
        double o2d(object o)
        {
            var ss = o2s(o);

            double d;
            double.TryParse(ss, out d);
            return Math.Round(d,2);
        }

        public Wd.Table FindTable(Wd.Document doc, Wd.Bookmark bm)
        {
            foreach (Wd.Table t in doc.Tables)
            {
                if (bm.Range.Start >= t.Range.Start && bm.Range.End <= t.Range.End)
                {

                    return t;
                }
            }
            return null;
        }

        AllData LoadDataFromExcel(string fn)
        {
            var ad = new AllData();
            try
            {
                ExcelApp = new Ex.Application() { Visible = false };
                ExcelWorkbook = ExcelApp.Workbooks.Open(fn);
                var ws = ExcelWorkbook.Worksheets["Свод"];

                ad.Address = ws.Cells[4, 3].Value;
                ad.ForWho = ws.Cells[5, 3].Value;
                ad.ServiceType = ws.Cells[7, 3].Value;
                ad.DateStart = o2s(ws.Cells[8, 3].Value);
                ad.ContractPeriod = Convert.ToString(ws.Cells[9, 3].Value);
                ad.CostMonth = Convert.ToString(Math.Round(ws.Cells[10, 3].Value,2));
                ad.CostMonthString = GetCostText(o2d(ws.Cells[10, 3].Value));
                ad.ContractPeriodYears =( o2d(ad.ContractPeriod) / 12).ToString();
                ad.DaysToStart = "1";
                ad.Date = DateTime.Now.ToString("dd.MM.yyyy");

                //////расх материалы
                ad.RashMaterials = new List<ExtraItem>();
                var wsData= ExcelWorkbook.Worksheets["Данные"];
                if (wsData != null)
                {
                    for (int i =200; i < 400; ++i)
                    {
                        if (o2s(wsData.Cells[i, 3].Value).Contains("Расходные материалы для санзон"))
                        {
                            while (o2s(wsData.Cells[i + 2, 3].Value).Contains("норма расхода"))
                            {
                                var val = o2s(wsData.Cells[i + 1, 7].Value);
                                double dd;
                                if (double.TryParse(val, out dd))
                                {
                                    double dd1;
                                    var val1 = o2s(wsData.Cells[i + 2, 7].Value);
                                    double.TryParse(val1, out dd1);
                                    if (dd != 0)
                                    {
                                        ad.RashMaterials.Add(new ExtraItem() { Name = o2s(wsData.Cells[i + 1, 3].Value), Quant = dd1, CostMonth = dd , Price=dd/dd1});
                                    }
                                }
                                i += 2;
                            }

                            break;
                        }

                    }
                }
                dataGridRash.ItemsSource = ad.RashMaterials;

                //////осн средства

                ad.OsnSredstva = new List<ExtraItem>();
                var wsOS = ExcelWorkbook.Worksheets["ОС"];
                if (wsOS != null)
                {
                    int K = 1;
                    for (int i = 1; i < 120; ++i)
                    {
                        var Cost = o2d(wsOS.Cells[i, 7].Value);
                        if(Cost!=0)
                        {
                            double quant= o2d(wsOS.Cells[i, 6].Value);
                            string name = o2s(wsOS.Cells[i, 3].Value);
                            if (!name.ToLower().Contains("итого")&&!string.IsNullOrWhiteSpace(name))
                            {
                                ad.OsnSredstva.Add(new ExtraItem() { Name = name, Quant = quant, CostMonth = Cost,Num=K});
                                ++K;
                            }
                        }

                    }
                }
                dataGridOsnSredst.ItemsSource = ad.OsnSredstva;

                //////штат расстановка
                ad.ShtatRasstanovka = new List<ExtraItem>();
                var wsFOT = ExcelWorkbook.Worksheets["ФОТ"];
                if (wsFOT != null)
                {
                    int K = 1;
                    for (int i = 1; i < 150; ++i)
                    {
                        var days = o2s(wsFOT.Cells[i, 4].Value);
                        var start = o2s(wsFOT.Cells[i, 5].Value);
                        var end= o2s(wsFOT.Cells[i, 6].Value);
                        var q= o2d(wsFOT.Cells[i, 8].Value);
                        if (!string.IsNullOrWhiteSpace(days)&&!string.IsNullOrWhiteSpace(start) &&!string.IsNullOrWhiteSpace(end)&&q!=0)
                        {
                            string name = o2s(wsFOT.Cells[i, 3].Value);
                            ad.ShtatRasstanovka.Add(new ExtraItem() { Name = name, Quant = q, Smena=days+" c "+start+" до "+end,Num=K});
                            ++K;
                        }

                    }
                }
                dataGridShtat.ItemsSource = ad.ShtatRasstanovka;

                //////себестоимость
                ad.Sebestimost = new List<ExtraItem>();
                var wsSS = ExcelWorkbook.Worksheets["С-сть"];
                if (wsSS != null)
                {
                    int K = 1;
                    for (int i = 1; i < 150; ++i)
                    {
                        string code = o2s(wsSS.Cells[i, 2].Value);
                        if(code.Contains("_10"))
                        {
                            double cost= o2d(wsSS.Cells[i, 4].Value);
                            {
                                if(cost!=0)
                                {
                                    string name= o2s(wsSS.Cells[i, 3].Value);
                                    ad.Sebestimost.Add(new ExtraItem() { Name = name, CostMonth=cost});
                                }
                            }
                        }
                    }
                }
                dataGridRashStruct.ItemsSource = ad.Sebestimost;


                //////Общее
                
                var wsPAR = ExcelWorkbook.Worksheets["Параметры"];
                if (wsPAR != null)
                {
                    int K = 1;
                    for (int i = 1; i < 50; ++i)
                    {
                        string code = o2s(wsPAR.Cells[i, 2].Value);
                        if (code.Contains("Общая площадь предприятия"))
                        {
                            double quant = o2d(wsPAR.Cells[i, 3].Value);
                            ad.Area = quant;
                            break;
                        }
                    }
                }
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (ExcelWorkbook != null)
                {
                    ExcelWorkbook.Close(false);
                    ExcelWorkbook = null;
                }
                if (ExcelApp != null) 
                    ExcelApp.Quit();
            }
            return ad;
        }

        void ReplaceAll(Wd.Document doc,string name,string val)
        {
            while(doc.Content.Find.Execute(name, ReplaceWith: val, Wrap: 1))
            {

            }
        }

        private void buttonPDF_Click(object sender, RoutedEventArgs e)
        {
            var wrd_name = Convert.ToString(comboBox.SelectedValue);
            var xls_name = textBoxExcelPath.Text;



            if (string.IsNullOrEmpty(wrd_name))
            {
                MessageBox.Show("Выберите шаблон КП");
                return;
            }

            ProgressWnd pw = new ProgressWnd("Создаем КП по шаблону...");
            pw.Show();

            string fn = "";
            try
            {


                WordApp = new Wd.Application() { Visible = false };
                WordDoc = WordApp.Documents.Open(wrd_name);
                fn=SaveW(WordApp, WordDoc);



                var ad = DataContext as AllData;


                ReplaceAll(WordDoc,"[OurCompany]", ad.Our.OurCompany);
                ReplaceAll(WordDoc, "[ForWho]", ad.ForWho);
                ReplaceAll(WordDoc, "[ContractPeriodYears]", ad.HowLongYears);
                ReplaceAll(WordDoc, "[ServiceType]", ad.ServiceType);
                ReplaceAll(WordDoc, "[DaysToStart]", ad.DaysToStart);
                ReplaceAll(WordDoc, "[DateStart]", ad.DateStart);
                ReplaceAll(WordDoc, "[CostMonthString]", ad.CostMonthString);
                ReplaceAll(WordDoc, "[CostMonth]", ad.CostMonth);
                ReplaceAll(WordDoc, "[CostYear]", (o2d(ad.CostMonth) * 12).ToString());
                ReplaceAll(WordDoc, "[Address]", ad.Address);
                ReplaceAll(WordDoc, "[Date]", ad.Date);
                ReplaceAll(WordDoc, "[Area]", ad.Area.ToString());
                ReplaceAll(WordDoc, "[NameObject]", ad.ForWho);

                int kk = 1;
                foreach(var ee in ad.e1)
                {
                    var trg="[Extra" + kk + "]";
                    if (ee.IsSelected==true)
                    {
                        ReplaceAll(WordDoc, trg, ee.Text);
                    }
                    else
                    {
                        WordDoc.Bookmarks["bExtra" + kk].Range.Delete();
                        ReplaceAll(WordDoc, trg, "");
                    }
                    ++kk;
                }

                kk = 1;
                foreach (var ee in ad.e2)
                {
                    var trg = "[Sod" + kk + "]";
                    if (ee.IsSelected == true)
                    {
                        ReplaceAll(WordDoc, trg, ee.Text);
                    }
                    else
                    {
                        WordDoc.Bookmarks["bSod" + kk].Range.Delete();
                        ReplaceAll(WordDoc, trg, "");
                    }
                    ++kk;
                }

                if(radioButton.IsChecked==true)
                {
                    ReplaceAll(WordDoc, "[NaOsnovanii]", ad.NaOsn1);
                }
                else
                {
                    ReplaceAll(WordDoc, "[NaOsnovanii]", ad.NaOsn2);
                }

                if(checkBoxServicePlan.IsChecked==false)
                {
                    Wd.Range delRange = WordDoc.Range();
                    delRange.Start=WordDoc.Bookmarks["b0"].Range.Start;
                    delRange.End = WordDoc.Bookmarks["b1"].Range.End;
                    delRange.Delete();

                }
                ReplaceAll(WordDoc, "[ServicePlanStart]", "");
                ReplaceAll(WordDoc, "[ServicePlanEnd]", "");
                //ReplaceAll(WordDoc, "[NameCost]", ad.ForWho);


                //Sebestoimost
                Wd.Table tt=FindTable(WordDoc, WordDoc.Bookmarks["tab6"]);

                double total = 0;
                foreach (var s in ad.Sebestimost)
                {
                    tt.Rows.Add();

                    tt.Cell(tt.Rows.Count, 1).Range.Text = s.Name;
                    tt.Cell(tt.Rows.Count, 2).Range.Text = s.CostMonth.ToString();
                    total += s.CostMonth;
                }
                tt.Rows.Add();
                tt.Cell(tt.Rows.Count, 1).Range.Text = "Итого";
                tt.Cell(tt.Rows.Count, 2).Range.Text = total.ToString();
                ///

                //FOT
                Wd.Table ttFOT = FindTable(WordDoc, WordDoc.Bookmarks["tab7"]);

                total = 0;
                foreach (var s in ad.ShtatRasstanovka)
                {
                    ttFOT.Rows.Add();

                    ttFOT.Cell(tt.Rows.Count, 1).Range.Text = s.Num.ToString();
                    ttFOT.Cell(tt.Rows.Count, 2).Range.Text = s.Name;
                    ttFOT.Cell(tt.Rows.Count, 3).Range.Text = s.Quant.ToString();
                    ttFOT.Cell(tt.Rows.Count, 4).Range.Text = s.Smena.ToString();
                }
                ////

                //os sredstv
                Wd.Table ttOS = FindTable(WordDoc, WordDoc.Bookmarks["tab8"]);

                total = 0;
                foreach (var s in ad.OsnSredstva)
                {
                    ttOS.Rows.Add();

                    ttOS.Cell(tt.Rows.Count, 1).Range.Text = s.Num.ToString()+".";
                    ttOS.Cell(tt.Rows.Count, 2).Range.Text = s.Name;
                    ttOS.Cell(tt.Rows.Count, 3).Range.Text = s.Quant.ToString();
                    ttOS.Cell(tt.Rows.Count, 4).Range.Text = s.CostMonth.ToString();

                    total += s.CostMonth;
                }
                ttOS.Rows.Add();
                ttOS.Cell(tt.Rows.Count, 1).Range.Text = "Итого";
                ttOS.Cell(tt.Rows.Count, 4).Range.Text = total.ToString();


                //rash san
                Wd.Table ttSAN = FindTable(WordDoc, WordDoc.Bookmarks["tab9"]);

                total = 0;
                foreach (var s in ad.RashMaterials)
                {
                    ttSAN.Rows.Add();

                    ttSAN.Cell(tt.Rows.Count, 1).Range.Text = s.Name;
                    ttSAN.Cell(tt.Rows.Count, 2).Range.Text = s.Quant.ToString();
                    ttSAN.Cell(tt.Rows.Count, 3).Range.Text = s.Price.ToString();
                    ttSAN.Cell(tt.Rows.Count, 4).Range.Text = s.CostMonth.ToString();

                    total += s.CostMonth;
                }
                ttSAN.Rows.Add();
                ttSAN.Cell(tt.Rows.Count, 1).Range.Text = "Итого в месяц";
                ttSAN.Cell(tt.Rows.Count, 4).Range.Text = total.ToString();


                StringBuilder sb = new StringBuilder();
                foreach(var c in ad.Contacts)
                {
                    if(c.IsSelected)
                    {
                        sb.AppendLine(c.Dolj);
                        sb.AppendLine(c.Name);
                        sb.AppendLine(c.Tel);
                        sb.AppendLine(c.MobTel);
                        sb.AppendLine(c.Email);
                        sb.AppendLine();
                    }
                }
                ReplaceAll(WordDoc, "[Contacts]", sb.ToString());


                WordDoc.TablesOfContents[1].Update();


                foreach (Wd.Range docRange in WordDoc.Words)
                {
                    docRange.HighlightColorIndex = Wd.WdColorIndex.wdNoHighlight;
                }

                


            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
            finally
            {
                WordDoc.Save();
                WordApp.Documents.Close();
                WordApp.Quit(Wd.WdSaveOptions.wdDoNotSaveChanges);

                string convertedXpsDoc = string.Concat(System.IO.Path.GetTempPath(), "\\", Guid.NewGuid().ToString(), ".xps");
                XpsDocument xpsDocument = ConvertWordToXps(fn, convertedXpsDoc);
                if (xpsDocument != null)
                {
                    documentViewer.Document = xpsDocument.GetFixedDocumentSequence();
                }

                
                pw.Stop();


            }


            // doc.Content.Find.Execute("[ND]", ReplaceWith: Doc.No);

            //   SaveW(app, doc, TemplateFilePath);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var ad = DataContext as AllData;
            SaveOurData(ad.Our);
        }

        private void buttonLoadEx_Click(object sender, RoutedEventArgs e)
        {
            button1_Click(null, null);

            
            ProgressWnd pw = new ProgressWnd("Загрузка данных из Excel...");
            pw.Show();
            var ad = LoadDataFromExcel(textBoxExcelPath.Text);
            pw.Stop();
            


            ad.Our = (DataContext as AllData).Our;
            DataContext = ad;

            dataGridExtra1.ItemsSource = ad.e1;
            dataGridExtraSod.ItemsSource = ad.e2;
            dataGridContacts.ItemsSource = ad.Contacts;
        }

        private void buttonPrint_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.Print();
        }

        private void buttonSaveAs_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".doc";
            dlg.Filter = "Word Document (*.doc)|*.doc";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                if (File.Exists("Результат.doc"))
                {
                    string filename = dlg.FileName;
                    File.Copy("Результат.doc", filename);
                    // Open document 
                    System.Diagnostics.Process.Start(filename);
                }
             
                
            }
        }
    }
}
