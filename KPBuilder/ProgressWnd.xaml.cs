using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace KPBuilder
{
    /// <summary>
    /// Interaction logic for ProgressWnd.xaml
    /// </summary>
    public partial class ProgressWnd : Window
    {
        public ProgressWnd(string WhatToDO)
        {
            InitializeComponent();
            Title = "Идет процесс...";
            textBlock.Text = WhatToDO;
        }
     
        public void Stop()
        {
            Close();
        }
    }

  
}
