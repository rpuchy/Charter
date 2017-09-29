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

namespace Charter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private double[] Item1 =new double[] { .228, .285, .366, .478, .629, .808, 1.031, 1.110};
        private double[] Item2 = new double[] { 1.228, 1.285, 1.366, 1.478, 1.629, 1.808, 2.031, 2.110 };



        public MainWindow()
        {
            InitializeComponent();
            AllItems = new ObservableCollection<string>() { "Item1", "Item2" };
            ChartItems = new ObservableCollection<string>();
            ChartItems.CollectionChanged += ChartItems_CollectionChanged;
            SeriesCollection = new SeriesCollection
            {
            };
            //    new StackedAreaSeries
            //    {
            //        Title = "Africa",
            //        Fill = System.Windows.Media.Brushes.Transparent,
            //       Stroke = System.Windows.Media.Brushes.Black,
            //        Values = new ChartValues<DateTimePoint>
            //        {
            //            new DateTimePoint(new DateTime(1950, 1, 1), .228),
            //            new DateTimePoint(new DateTime(1960, 1, 1), .285),
            //            new DateTimePoint(new DateTime(1970, 1, 1), .366),
            //            new DateTimePoint(new DateTime(1980, 1, 1), .478),
            //            new DateTimePoint(new DateTime(1990, 1, 1), .629),
            //            new DateTimePoint(new DateTime(2000, 1, 1), .808),
            //            new DateTimePoint(new DateTime(2010, 1, 1), 1.031),
            //            new DateTimePoint(new DateTime(2013, 1, 1), 1.110)
            //        },
            //        LineSmoothness = 0
            //    },
            //    new StackedAreaSeries
            //    {
            //        Title = "N & S America",
            //        Fill = System.Windows.Media.Brushes.Black,
            //        Values = new ChartValues<DateTimePoint>
            //        {
            //            new DateTimePoint(new DateTime(1950, 1, 1), .339),
            //            new DateTimePoint(new DateTime(1960, 1, 1), .424),
            //            new DateTimePoint(new DateTime(1970, 1, 1), .519),
            //            new DateTimePoint(new DateTime(1980, 1, 1), .618),
            //            new DateTimePoint(new DateTime(1990, 1, 1), .727),
            //            new DateTimePoint(new DateTime(2000, 1, 1), .841),
            //            new DateTimePoint(new DateTime(2010, 1, 1), .942),
            //            new DateTimePoint(new DateTime(2013, 1, 1), .972)
            //        },
            //        LineSmoothness = 0
            //    },
            //    new StackedAreaSeries
            //    {
            //        Title = "Asia",
            //        Values = new ChartValues<DateTimePoint>
            //        {
            //            new DateTimePoint(new DateTime(1950, 1, 1), 1.395),
            //            new DateTimePoint(new DateTime(1960, 1, 1), 1.694),
            //            new DateTimePoint(new DateTime(1970, 1, 1), 2.128),
            //            new DateTimePoint(new DateTime(1980, 1, 1), 2.634),
            //            new DateTimePoint(new DateTime(1990, 1, 1), 3.213),
            //            new DateTimePoint(new DateTime(2000, 1, 1), 3.717),
            //            new DateTimePoint(new DateTime(2010, 1, 1), 4.165),
            //            new DateTimePoint(new DateTime(2013, 1, 1), 4.298)
            //        },
            //        LineSmoothness = 0
            //    },
            //    new StackedAreaSeries
            //    {
            //        Title = "Europe",
            //        Values = new ChartValues<DateTimePoint>
            //        {
            //            new DateTimePoint(new DateTime(1950, 1, 1), .549),
            //            new DateTimePoint(new DateTime(1960, 1, 1), .605),
            //            new DateTimePoint(new DateTime(1970, 1, 1), .657),
            //            new DateTimePoint(new DateTime(1980, 1, 1), .694),
            //            new DateTimePoint(new DateTime(1990, 1, 1), .723),
            //            new DateTimePoint(new DateTime(2000, 1, 1), .729),
            //            new DateTimePoint(new DateTime(2010, 1, 1), .740),
            //            new DateTimePoint(new DateTime(2013, 1, 1), .742)
            //        },
            //        LineSmoothness = 0
            //    }
            //};

            XFormatter = val => new DateTime((long)val).ToString("yyyy");
            YFormatter = val => val.ToString("N") + " M";

            DataContext = this;

            foreach (var series in SeriesCollection)
            {
                foreach(DateTimePoint datapoint in series.Values)
                {
                    var x = datapoint.DateTime;
                    var y = datapoint.Value;
                }

            }
           // ExportToExcel();
        }

        private void ChartItems_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (string item in e.OldItems)
                {
                    SeriesCollection.Remove(SeriesCollection.First(x => x.Title == item));
                }
            }
            if (e.NewItems != null)
            {
                foreach (string item in e.NewItems)
                {
                    double[] data;
                    if (item == "Item1")
                    {
                        data = Item1;
                    }
                    else
                    {
                        data = Item2;
                    }
                    var series = new LineSeries()
                    {
                        Title = item,
                        LineSmoothness = 1,
                        Fill = Brushes.Transparent,
                        //Stroke = Brushes.Black,
                        StrokeThickness = 2,
                        //StackMode = StackMode.Values,
                        
                        Values = new ChartValues<DateTimePoint>()
                    };
                    
                    series.Values.AddRange(data.Select((x, i) => new DateTimePoint(new DateTime(2000 + i, 1, 1), x)));
                    SeriesCollection.Add(series);
                }
            }

        }

        public void ExportToExcel()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //add data 
            for (int i=0;i<SeriesCollection.Count;i++)
            {
                xlWorkSheet.Cells[1, i + 2] = SeriesCollection[i].Title;
                for (int j=0;j<SeriesCollection[i].Values.Count;j++)
                {
                    if (i==0)
                    {
                        xlWorkSheet.Cells[j + 2, i + 1] = (SeriesCollection[i].Values[j] as DateTimePoint).DateTime;
                    }
                    xlWorkSheet.Cells[j + 2, i + 2] = (SeriesCollection[i].Values[j] as DateTimePoint).Value;
                }
            }

            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range("A1", "e9");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Excel.XlChartType.xlAreaStacked;

            xlWorkBook.SaveAs("c:\\temp\\csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file c:\\csharp.net-informations.xls");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public SeriesCollection SeriesCollection { get; set; }
        public ObservableCollection<string> ChartItems { get; set; }
        public ObservableCollection<string> AllItems { get; set; }
        public Func<double, string> XFormatter { get; set; }
        public Func<double, string> YFormatter { get; set; }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (AllInputs.SelectedIndex == -1) return;
            ChartItems.Add(AllInputs.SelectedItem.ToString());
            AllItems.Remove(AllInputs.SelectedItem.ToString());
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            if (ChartInputs.SelectedIndex == -1) return;
            AllItems.Add(ChartInputs.SelectedItem.ToString());
            ChartItems.Remove(ChartInputs.SelectedItem.ToString());
        }
    }
}
