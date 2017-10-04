using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using Microsoft.VisualBasic.FileIO;
using System.IO;
using LiveCharts;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using LiveCharts.Defaults;
using LiveCharts.Wpf;

namespace Charter
{
    /// <summary>
    /// Interaction logic for ChartingControl.xaml
    /// </summary>
    public partial class ChartingControl : UserControl
    {
        private string csvData;
        private Dictionary<string, List<double>> data = new Dictionary<string, List<double>>();
        public SeriesCollection SeriesCollection { get; set; }
        public ObservableCollection<string> ChartItems { get; set; }
        public ObservableCollection<string> AllItems { get; set; }
        public Func<double, string> XFormatter { get; set; }
        public Func<double, string> YFormatter { get; set; }

        public string CsvData
        {
            get { return csvData; }
            set { csvData = value; loadData(); }
        }

        private static Stream GenerateStreamFromString(string s)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }


        private void loadData()
        {
            data.Clear();
            SeriesCollection.Clear();
            using (TextFieldParser csvParser = new TextFieldParser(GenerateStreamFromString(csvData)))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                //Since we are only getting from timestep 1 we need a timestep 0.
                bool foundStart = false;
                int t0index = 0;
                int endindex = 0;
                double temp;
                while (!csvParser.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    var fields = csvParser.ReadFields();

                    if (foundStart)
                    {
                        data.Add(fields[0], fields.Skip(t0index).Take(endindex - t0index).Select(x => Double.TryParse(x, out temp) ? temp : 0).ToList());
                        AllItems.Add(fields[0]);
                    }

                    if (!foundStart && fields[0] == "Operator Key")
                    {
                        foundStart = true;
                        t0index = Array.IndexOf(fields.ToArray(), "0");
                        endindex = Array.IndexOf(fields.ToArray(), "all");
                    }
                }
            }

            var toRemove = new List<string>();

            foreach (string item in ChartItems)
            {
                if (!AllItems.Contains(item))
                {
                    toRemove.Add(item);
                    continue;
                }

                double[] datainfo = data[item].ToArray();
                AllItems.Remove(item);
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

                series.Values.AddRange(datainfo.Select((x, i) => new DateTimePoint(new DateTime(2000 + i, 1, 1), x)));
                SeriesCollection.Add(series);
            }

            foreach(var item in toRemove)
            {
                ChartItems.Remove(item);
            }


        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (AllInputs.SelectedIndex == -1) return;
            var tochange = new List<string>();
            foreach (string item in AllInputs.SelectedItems)
            {
                tochange.Add(item);
            }

            foreach (string item in tochange)
            {
                ChartItems.Add(item);
                AllItems.Remove(item);
            }


        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            if (ChartInputs.SelectedIndex == -1) return;
            var tochange = new List<string>();
            foreach (string item in ChartInputs.SelectedItems)
            {
                tochange.Add(item);
            }
            foreach (string item in tochange)
            {
                AllItems.Add(item);
                ChartItems.Remove(item);
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
            for (int i = 0; i < SeriesCollection.Count; i++)
            {
                xlWorkSheet.Cells[1, i + 2] = SeriesCollection[i].Title;
                for (int j = 0; j < SeriesCollection[i].Values.Count; j++)
                {
                    if (i == 0)
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
                    double[] datainfo = data[item].ToArray();
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

                    series.Values.AddRange(datainfo.Select((x, i) => new DateTimePoint(new DateTime(2000 + i, 1, 1), x)));
                    SeriesCollection.Add(series);
                }
            }

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

        public ChartingControl()
        {
            AllItems = new ObservableCollection<string>();
            ChartItems = new ObservableCollection<string>();
            ChartItems.CollectionChanged += ChartItems_CollectionChanged;
            SeriesCollection = new SeriesCollection
            {
            };


            XFormatter = val => new DateTime((long)val).ToString("yyyy");
            YFormatter = val => val.ToString("N") + " M";
            DataContext = this;
            InitializeComponent();
        }
    }
}
