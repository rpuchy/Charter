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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using LiveCharts;
using LiveCharts.Defaults;
using LiveCharts.Wpf;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using System.IO;

namespace Charter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private Dictionary<string, List<double>> data = new Dictionary<string, List<double>>();


        public MainWindow()
        {
            InitializeComponent();


            //foreach (var series in SeriesCollection)
            //{
            //    foreach(DateTimePoint datapoint in series.Values)
            //    {
            //        var x = datapoint.DateTime;
            //        var y = datapoint.Value;
            //    }

            //}

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            // openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == true)
            {
                 ChartingCont.CsvData = File.ReadAllText(openFileDialog1.FileName);

            }

            // ExportToExcel();
        }




        

        

        



        
    }
}
