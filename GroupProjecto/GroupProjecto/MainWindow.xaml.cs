using System;
using System.Data;
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
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace GroupProjecto
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<DateTime> listOfHolidays = new List<DateTime>();
        DateTime holidayDate = new DateTime();
        List<Topic> TopicList = new List<Topic>();
        List<Holiday> HolidayList = new List<Holiday>();
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
            Excel.Worksheet xlWorkSheet, holidayWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            holidayWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            holidayWorkSheet.Name = "Holidays";

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet.Name = "Topics";

            xlWorkSheet.Cells[1, 1] = "Topic";
            xlWorkSheet.Columns[1].ColumnWidth = 18;
            xlWorkSheet.Cells[1, 2] = "# Days";
            xlWorkSheet.Columns[2].ColumnWidth = 6;
            xlWorkSheet.Cells[1, 3] = "Notes";
            xlWorkSheet.Columns[3].ColumnWidth = 48;

            holidayWorkSheet.Cells[1, 1] = "Holiday Date";
            holidayWorkSheet.Columns[1].ColumnWidth = 12;
            holidayWorkSheet.Cells[1, 2] = "Holiday Description";
            holidayWorkSheet.Columns[2].ColumnWidth = 48;



            xlWorkBook.SaveAs($"{docFolderPath}\\Topics", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show($"Excel file created, you can find the file at {docFolderPath}\\Topics.xls.\n\nPlease fill out both the Topics and Holidays worksheets in that excel file before continuing.");
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
                FileStream fs = File.Open(SelectFileTB.Text, FileMode.Open, FileAccess.Read);
                IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fs);
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                var topicTable = result.Tables[0];
                var holidayTable = result.Tables[1];

                for (var i = 0; i < topicTable.Rows.Count; i++)
                {
                    Topic topic = new Topic(topicTable.Rows[i][0].ToString(), Convert.ToInt32(topicTable.Rows[i][1]), topicTable.Rows[i][2].ToString());
                    TopicList.Add(topic);
                }

                for (int i = 0; i < holidayTable.Rows.Count; i++)
                {
                    Holiday holiday = new Holiday(Convert.ToDateTime(holidayTable.Rows[i][0]), Convert.ToString(holidayTable.Rows[i][1]));
                    HolidayList.Add(holiday);
                }





                reader.Close();
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
            xlWorkSheet.Columns[1].ColumnWidth = 6;
            xlWorkSheet.Cells[1, 2] = "Day";
            xlWorkSheet.Columns[2].ColumnWidth = 10;
            xlWorkSheet.Cells[1, 3] = "Date";
            xlWorkSheet.Columns[3].ColumnWidth = 16;
            xlWorkSheet.Cells[1, 4] = "Topic";
            xlWorkSheet.Columns[4].ColumnWidth = 18;
            xlWorkSheet.Cells[1, 5] = "Notes";
            xlWorkSheet.Columns[5].ColumnWidth = 48;

            int calendarMWProcessed = 0;
            foreach (var topic in TopicList)
            {
                for (int i = 0; i < topic.Days; i++)
                {
                    Holiday holiday = null;
                    foreach (var item in HolidayList)
                    {
                        if(getCurrentDate(firstMonday, calendarMWProcessed) == item.HolidayDate)
                        {
                            holiday = item;
                        }
                    }
                    while (holiday != null)
                    {
                        if (calendarMWProcessed % 2 == 0) // Mondays
                        {
                            xlWorkSheet.Cells[calendarMWProcessed + 2, 1] = (calendarMWProcessed + 2) / 2;
                            xlWorkSheet.Cells[calendarMWProcessed + 2, 2] = "Monday";
                            xlWorkSheet.Cells[calendarMWProcessed + 2, 3] = getCurrentDate(firstMonday, calendarMWProcessed);
                        }
                        else // Wednesdays
                        {
                            xlWorkSheet.Cells[calendarMWProcessed + 2, 1] = (calendarMWProcessed + 1) / 2;
                            xlWorkSheet.Cells[calendarMWProcessed + 2, 2] = "Wednesday";
                            xlWorkSheet.Cells[calendarMWProcessed + 2, 3] = getCurrentDate(firstMonday, calendarMWProcessed);
                        }

                        xlWorkSheet.Cells[calendarMWProcessed + 2, 4] = holiday.HolidayDescription;

                        calendarMWProcessed++;
                        holiday = null;
                        foreach (var item in HolidayList)
                        {
                            if (getCurrentDate(firstMonday, calendarMWProcessed) == item.HolidayDate)
                            {
                                holiday = item;
                            }
                        }

                    }
                    if (calendarMWProcessed % 2 == 0) // Mondays
                    {
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 1] = (calendarMWProcessed + 2) / 2;
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 2] = "Monday";
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 3] = getCurrentDate(firstMonday, calendarMWProcessed);
                    } else // Wednesdays
                    {
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 1] = (calendarMWProcessed + 1) / 2;
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 2] = "Wednesday";
                        xlWorkSheet.Cells[calendarMWProcessed + 2, 3] = getCurrentDate(firstMonday, calendarMWProcessed);
                    }
                   
                    xlWorkSheet.Cells[calendarMWProcessed + 2, 4] = topic.Name;
                    xlWorkSheet.Cells[calendarMWProcessed + 2, 5] = topic.Notes;

                    calendarMWProcessed++;
                }
            }
            

            xlWorkBook.SaveAs($"{docFolderPath}\\TopicsGeneratedSchedule.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            userExcel.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(userExcel);

            MessageBox.Show($"Excel file created , you can find the file {docFolderPath}\\TopicsGeneratedSchedule.csv");
            //Jenna sucks
        }

        public DateTime getCurrentDate(DateTime firstMonday, int mwprocessed)
        {
            if (mwprocessed % 2 == 0) // Mondays
            {
                return firstMonday.AddDays(7 * (((mwprocessed + 2) / 2) - 1));
            }
            else // Wednesdays
            {
                return firstMonday.AddDays(7 * (((mwprocessed + 1) / 2) - 1) + 2);
            }
        }
    }
}


