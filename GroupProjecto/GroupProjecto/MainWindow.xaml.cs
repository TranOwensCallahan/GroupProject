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
        DateTime classDate = new DateTime();
        List<DateTime> schoolDaysList = new List<DateTime>();
        List<DateTime> holidayDatesList = new List<DateTime>();
        List<string> TopicList = new List<string>();
        List<string> DaysList = new List<string>();
        List<string> NotesList = new List<string>();
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
            xlWorkSheet.Cells[1, 4] = "Enter the holidays";
            xlWorkSheet.Cells[1, 2] = "Notes";
            xlWorkSheet.Cells[1, 3] = "Enter the first Date";




            
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
                    string topic = column[0];
                    DateTime holidayDate = Convert.ToDateTime(column[3]);
                    string notes = column[1];
                    DateTime classDate = Convert.ToDateTime(column[2]);
                    TopicList.Add(topic);
                    holidayDatesList.Add(holidayDate);
                    NotesList.Add(notes);
                    
                    
                    
                }
            }
        }

        private void CreateBtn_Click(object sender, RoutedEventArgs e)
        {
            
            Excel.Application userExcel = new Microsoft.Office.Interop.Excel.Application();
            for (int i = 0; i < TopicList.Count; i++)
            {
                schoolDaysList.Add(classDate);
                schoolDaysList.Add(classDate.AddDays(2));
                classDate = classDate.AddDays(7);
              
            }

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
            for (int i = 0; i < TopicList.Count; i++)
            {
                int a = 1;
                xlWorkSheet.Cells[i + 2, 1 + a] = weeks;
                xlWorkSheet.Cells[i+2, 4] = TopicList[i];
                xlWorkSheet.Cells[i + 2, 2] = schoolDaysList[i];
                xlWorkSheet.Cells[i + 2, 5] = NotesList[i];
                weeks++;
                a++;
            }
            

            xlWorkBook.SaveAs($"{docFolderPath}\\csharpgeneratedExcel.csv", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
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


