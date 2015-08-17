using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPBuilder
{
    public class AllData
    {
        public string ServiceType { get; set; }
        public string ForWho { get; set; }
        public OurData Our { get; set; }
        public string HowLongYears { get; set; }
        public string TotalCost { get; set; }
        public string TotalCostString { get; set; }
        public string Address { get;  set; }
        public string DateStart { get;  set; }
        public string ContractPeriod { get;  set; }
        public string ContractPeriodYears { get; set; }
        public string CostMonth { get;  set; }
        public string CostMonthString { get; set; }
        public string DaysToStart { get; set; }
        public string Date { get; set; }


        public List<ExtraItem> RashMaterials { get; set; }
        public List<ExtraItem> OsnSredstva { get; set; }
        public List<ExtraItem> ShtatRasstanovka { get; set; }
        public List<ExtraItem> Sebestimost { get; set; }

        public List<ExtraItem> Contacts { get; set; }

        public double Area { get; internal set; }

        public string NaOsn1 = "технического задания";
        public string NaOsn2 = "нормативов расходов ОМС";

        public ExtraItem[] e1 = new ExtraItem[] {
            new ExtraItem() {Text = "Проводить комплексную уборку внутренних помещений объектов;" },
            new ExtraItem() { Text = "Проводить поддерживающую уборку внутренних помещений объектов;" },
            new ExtraItem() { Text = "Проводить уборку территорий объекта;" }
        };

        public ExtraItem[] e2 = new ExtraItem[] {
            new ExtraItem() {Text = "Предоставить Исполнителю служебные помещения для размещения персонала и склада товарно-материальных ценностей (ТМЦ)." },
            new ExtraItem() { Text = "Предоставить Исполнителю мебель, необходимую для персонала и склада ТМЦ." },
            new ExtraItem() { Text = "Содействовать переводу в штат Исполнителя сотрудников." }
        };

        public AllData()
        {
            Contacts = File.ReadAllLines("people.txt").Select(m =>
              {
                  var mm = m.Split('*');
                  return new ExtraItem()
                  {
                      Dolj = mm[0],
                      Name = mm[1],
                      Tel = mm[2],
                      MobTel = mm[3],
                      Email = mm[4]
                  };
              }).ToList();
        }
    }
}
