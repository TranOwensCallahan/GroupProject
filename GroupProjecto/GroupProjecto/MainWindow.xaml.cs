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
        List<Topic> TopicList = new List<Topic>();
        string docFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

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
            xlWorkSheet.Cells[1, 2] = "# Topic Days";
            xlWorkSheet.Cells[1, 3] = "Notes";




            
            xlWorkBook.SaveAs($"{docFolderPath}\\csharp", Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show($"Excel file created , you can find the file {docFolderPath}\\csharp");
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
         

            if (File.Exists(SelectFileTB.Text) == true)
            {// if file exists read all the lines
                var lines = File.ReadAllLines(SelectFileTB.Text);
                for (int i = 1; i < lines.Length; i++)
                {
                    var line = lines[i];
                    var column = line.Split(',');
                    Topic topic = new Topic(column[0], Convert.ToInt32(column[1]), column[2]);
                    TopicList.Add(topic); 
                }
                readFileStatus.Items.Add("File Read Successfully.");
            }
        }

        private void CreateBtn_Click(object sender, RoutedEventArgs e)
        {
            
            Excel.Application userExcel = new Microsoft.Office.Interop.Excel.Application();
            var firstMonday = FirstMonday.SelectedDate.Value.Date;

            if (userExcel == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = userExcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Week";
            xlWorkSheet.Cells[1, 2] = "Day";
            xlWorkSheet.Cells[1, 3] = "Date";
            xlWorkSheet.Cells[1, 4] = "Topic";
            xlWorkSheet.Cells[1, 5] = "Notes";

            int weeks = 1;
            int calendarMWProcessed = 0;
            foreach (var topic in TopicList)
            {
                for (int i = 0; i < topic.Days; i++)
                {
                    if (calendarMWProcessed % 2 == 0) // Mondays
                    {
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 1] = (calendarMWProcessed + 2) / 2;
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 2] = "Monday";
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 3] = firstMonday.AddDays(7 * (((calendarMWProcessed + 2 )/ 2)-1));
                    } else // Wednesdays
                    {
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 1] = (calendarMWProcessed + 1) / 2;
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 2] = "Wednesday";
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 3] = firstMonday.AddDays(7 * (((calendarMWProcessed + 1) / 2) - 1)+2); ;
                    }

                    xlWorkSheet.Cells[calendarMWProcessed + 2, 4] = topic.Name;
                    xlWorkSheet.Cells[calendarMWProcessed + 2, 5] = topic.Notes;

                    calendarMWProcessed++;
                }
            }
            

            xlWorkBook.SaveAs($"{docFolderPath}\\csharpgeneratedExcel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            userExcel.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(userExcel);

            MessageBox.Show($"Excel file created , you can find the file {docFolderPath}\\generatedExcel.csv");
            //Jenna sucks
        }

    }
}


