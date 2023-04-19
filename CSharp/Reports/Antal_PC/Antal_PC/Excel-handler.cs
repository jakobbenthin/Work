using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Antal_PC
{
    internal class Excel_handler
    {
        public string ivantiFile { get; set; }
        public string ivantiSheet { get; set; }
        public string antalPCFile { get; set; }
        public string antalPCSheet { get; set; }



        public List<string> CheckDepartmentsInPCFile()
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(this.antalPCFile);
            Excel.Worksheet sheet = workbook.Sheets[this.antalPCSheet];

            // Copy the sheet
            sheet.Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);

            // Get a reference to the new sheet
            Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];

            // Get the current year and last month number
            DateTime currentDate = DateTime.Now;
            string sheetName = currentDate.Year.ToString() + "-" + currentDate.AddMonths(-1).ToString("MM");

            // Change the name of the new sheet to the year and last month number
            newSheet.Name = sheetName;

            // Save the workbook
            workbook.Save();

            Excel.Range usedRange = sheet.UsedRange;




            int rowCount = usedRange.Rows.Count;

            int columnIndexA = 4;
            int columnIndexB = 7;

            List<string> columnData = new List<string>();

            for(int i = 2; i <= rowCount; i++)
            {
                object valueA = usedRange.Cells[i, columnIndexA].Value;
                object valueB = usedRange.Cells[i, columnIndexB].Value;

                if (valueB == null)
                {
                    break;
                }
                if (valueB.ToString() == "Capio dator")
                {
                    if (valueA.ToString() == "Kostnadsställe")
                    {
                        string s = "";
                    }
                    string currentValue = valueA.ToString();

                    if (!columnData.Contains(currentValue))
                    {
                        columnData.Add(valueA.ToString());
                    }
                }
            }

            // Clean up
            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(sheet);
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            excel.Quit();
            Marshal.ReleaseComObject(excel);

            return columnData;
        }
        public Dictionary<string, int> CheckDepartmentsInIvantiFile()
        {

            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(this.ivantiFile);
            Excel.Worksheet sheet = workbook.Sheets[this.ivantiSheet];

            //Excel.Range columnRangeDep = sheet.Range["G:G"]; // Read data from column G
            //Excel.Range columnRangeHost = sheet.Range["O:O"]; // Read data from column O

            Excel.Range usedRange = sheet.UsedRange;

            int rowCount = usedRange.Rows.Count;

            int columnIndexA = 7;
            int columnIndexB = 15;

            List<string> columnDataList = new List<string>();
            Dictionary<string, int> columnData = new Dictionary<string, int>();

            /*
            System.Array valuesDepCol = (System.Array)columnRangeDep.Cells.Value;
            System.Array valuesHostCol = (System.Array)columnRangeHost.Cells.Value;

            string[] columnData2 = valuesDepCol.OfType<object>().Select(o => o.ToString()).ToArray();
            string[] columnData = new string[valuesDepCol.Length];
            */

            for (int i = 2; i <= rowCount; i++)
            {
                //Excel.Range cellA = (Excel.Range)columnRangeDep.Cells[i];
                //Excel.Range cellB = (Excel.Range)columnRangeHost.Cells[i];

                object valueA = usedRange.Cells[i, columnIndexA].Value;
                object valueB = usedRange.Cells[i, columnIndexB].Value;

                if (valueB == null)
                {
                    break;
                }
                if (valueB != null && valueB.ToString() != "")
                {
                    string currentValue = valueA.ToString();

                    
                    if (!columnDataList.Contains(currentValue))
                    {
                        columnDataList.Add(currentValue);
                    }

                    if (columnData.ContainsKey(currentValue))
                    {
                        columnData[currentValue]++;
                    }
                    else
                    {
                        columnData.Add(currentValue, 1);
                    }
                }

                //Marshal.ReleaseComObject(cellA);
                //Marshal.ReleaseComObject(cellB);
            }

            // Get unique values
            //HashSet<string> uniqueValues = new HashSet<string>(columnData);
            List<string> uniqueValuesList = columnDataList.Distinct().ToList();
            // Use uniqueValues HashSet as needed

            // Clean up
            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(sheet);
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            excel.Quit();
            Marshal.ReleaseComObject(excel);

            return columnData;

        }

    }

  
}
