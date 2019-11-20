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
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace GroupProjecto
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TopicBtn(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Topic";
            xlWorkSheet.Cells[1, 2] = "Days";




            xlWorkBook.SaveAs("d:\\csharp-Excel.csv", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.csv");
        }

        private void SelectFileBtn1_Click(object sender, RoutedEventArgs e)
        {
            //selecting a file these lines of code was given to me 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            var result = dlg.ShowDialog();
            // puts file name in the text box
            SelectFileTB.Text = dlg.FileName;
        }

        private void ReadFileBtn_Click(object sender, RoutedEventArgs e)
        {
            List<string> TopicList = new List<string>();
            List<string> DaysList = new List<string>();

            if (File.Exists(SelectFileTB.Text) == true)
            {// if file exists read all the lines
                var lines = File.ReadAllLines(SelectFileTB.Text);
                for (int i = 1; i < lines.Length-1; i++)
                {
                    var line = lines[i];
                    var column = line.Split(',');
                    string topic = column[1];
                    string days = column[2];
                    TopicList.Add(topic);
                    DaysList.Add(days);
                }
            }
        }
    }
}

