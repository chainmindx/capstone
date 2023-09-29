using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Globalization;

namespace filegen
{
    internal class filegen
    {
        public static void Main(string[] args)
        {
            DateTime startdate = DateTime.Parse("8/28/2023");
            DateTime enddate = DateTime.Parse("12/6/2023");
            DateTime lastregisterdate = DateTime.Parse("9/5/2023");
            DateTime lastdropdate = DateTime.Parse("11/10/2023");
            DateTime readingdate = DateTime.Parse("12/7/2023");
            DateTime finalexamdate;
            DateTime gradesduedate = DateTime.Parse("12/18/2023");

            string crn = "14028";
            string semester = "Fall 2023";
            string duration = startdate.ToString("MM/dd-") + enddate.ToString("MM/dd");
            string course = "OPSY 4314.001 Operations Management";
            string classtimes = "MWF 3:30-4:45PM";
            string classroom = "OCNR-130";

            string wkschedule = "MWF";

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range xlRange;

            xlApp = new Excel.Application();
            xlApp.Visible = true;

            xlWorkbook = xlApp.Workbooks.Add("");
            xlWorksheet = xlWorkbook.ActiveSheet;

            xlWorksheet.get_Range("A1");

            xlWorksheet.Cells[1, 1] = "CRN";
            xlWorksheet.Cells[2, 1] = "Semester";
            xlWorksheet.Cells[3, 1] = "Duration";
            xlWorksheet.Cells[4, 1] = "Course";
            xlWorksheet.Cells[5, 1] = "Class Times";
            xlWorksheet.Cells[6, 1] = "Classroom";
            xlWorksheet.Cells[7, 1] = "Class Schedule Template";

            //for loop?
            xlWorksheet.Cells[1, 1].Font.Bold = true;
            xlWorksheet.Cells[2, 1].Font.Bold = true;
            xlWorksheet.Cells[3, 1].Font.Bold = true;
            xlWorksheet.Cells[4, 1].Font.Bold = true;
            xlWorksheet.Cells[5, 1].Font.Bold = true;
            xlWorksheet.Cells[6, 1].Font.Bold = true;
            xlWorksheet.Cells[7, 1].Font.Bold = true;

            xlWorksheet.Cells[1, 2] = crn;
            xlWorksheet.Cells[2, 2] = semester;
            xlWorksheet.Cells[3, 2] = duration;
            xlWorksheet.Cells[4, 2] = course;
            xlWorksheet.Cells[5, 2] = classtimes;
            xlWorksheet.Cells[6, 2] = classroom;

            xlWorksheet.Cells[9, 1] = "Wk";
            xlWorksheet.Cells[9, 2] = "Day";
            xlWorksheet.Cells[9, 3] = "Date";
            xlWorksheet.Cells[9, 4] = "Description";

            DateTime iterativeDate = startdate;
            int weekstotal = 1;
            int cellspacing = 0;

            for (int i = 0; i < (gradesduedate - startdate).TotalDays; i++)
            {
                string dayoftheweek = iterativeDate.ToString("ddd");
                string daymonth = iterativeDate.ToString("dd-MMM");
                if (iterativeDate == lastregisterdate)
                {
                    xlWorksheet.Cells[cellspacing + 10, 2] = dayoftheweek;
                    xlWorksheet.Cells[cellspacing + 10, 3] = daymonth;
                    xlWorksheet.Cells[cellspacing + 10, 4] = "Last date to register for classes.";
                    cellspacing += 1;
                }
                if (iterativeDate == lastdropdate)
                {
                    xlWorksheet.Cells[cellspacing + 10, 2] = dayoftheweek;
                    xlWorksheet.Cells[cellspacing + 10, 3] = daymonth;
                    xlWorksheet.Cells[cellspacing + 10, 4] = "Last date to drop classes.";
                    cellspacing += 1;
                }
                
                if (wkschedule == "MWF" && iterativeDate != lastregisterdate && iterativeDate != lastdropdate)
                {
                    if ((dayoftheweek == "Mon" || dayoftheweek == "Wed" || dayoftheweek == "Fri") && iterativeDate < enddate)
                    {
                        if(dayoftheweek == "Mon")
                        {
                            xlWorksheet.Cells[cellspacing + 10, 1] = weekstotal.ToString();
                            weekstotal += 1;
                        }
                        
                        xlWorksheet.Cells[cellspacing + 10, 2] = dayoftheweek;
                        xlWorksheet.Cells[cellspacing + 10, 3] = daymonth;
                        cellspacing += 1;
                    }
                }
                //other formats
                iterativeDate = iterativeDate.AddDays(1);
            }

            //xlWorksheet.Columns.AutoFit();

            xlApp.Visible = false;
            xlApp.UserControl = false;

            if (File.Exists("C:\\Users\\jedde\\Desktop\\CAPSTONE\\sample2.xlsx")){
                File.Delete("C:\\Users\\jedde\\Desktop\\CAPSTONE\\sample2.xlsx");
            }

            xlWorkbook.SaveAs("C:\\Users\\jedde\\Desktop\\CAPSTONE\\sample2.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlWorkbook.Close();
            xlApp.Quit();
        }
    }
}
