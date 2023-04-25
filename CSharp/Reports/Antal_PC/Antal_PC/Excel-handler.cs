using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Antal_PC
{
    internal class Excel_handler
    {
        public string ivantiFile { get; set; }
        public string ivantiSheet { get; set; }
        public string antalPCFile { get; set; }
        public string antalPCSheet { get; set; }
        public string totalSheet { get; set; }
        

        public Dictionary<string, int> CheckDepartmentsInPCFile()
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(this.antalPCFile);
            Excel.Worksheet sheet = workbook.Sheets[this.antalPCSheet];


            try
            {
                // Get a reference to the table
                Excel.ListObject table = sheet.ListObjects["Innevarandemånadtabell"];

                // Get a reference to the column that you want to remove the color from
                Excel.Range columnToClear = table.ListColumns[6].Range;

                // Remove the color from the column
                columnToClear.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;

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
            }
            catch { }

            Excel.Range usedRange = sheet.UsedRange;

            int rowCount = usedRange.Rows.Count;

            int columnIndexA = 4;
            int columnIndexB = 7;
            int columnIndexC = 6;

            Dictionary<string, int> columnData = new Dictionary<string, int>();

            for(int i = 2; i <= rowCount; i++)
            {
                object valueA = usedRange.Cells[i, columnIndexA].Value;
                object valueB = usedRange.Cells[i, columnIndexB].Value;
                double valueC = usedRange.Cells[i, columnIndexC].Value;
                int count = Convert.ToInt32(valueC);

                if (valueB == null)
                {
                    break;
                }
                if (valueB.ToString() == "Capio dator")
                {
                    string currentValue = valueA.ToString();
                 
                    if (!columnData.ContainsKey(currentValue))
                    {
                        columnData.Add(valueA.ToString(), count);
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

            Excel.Range usedRange = sheet.UsedRange;

            int rowCount = usedRange.Rows.Count;

            int columnIndexA = 7;
            int columnIndexB = 15;

            List<string> columnDataList = new List<string>();
            Dictionary<string, int> columnData = new Dictionary<string, int>();
 
            for (int i = 2; i <= rowCount; i++)
            {          
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
            }

            // Get unique values
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
        public void WriteDataToExcel(Dictionary<string, int> dictionary, Dictionary<string, int> diff, Dictionary<string, int> antalPCdict)
        {
            int totalCapioComp = 0;
            int extCompTotal = 0;

            // Create an Excel application object
            Excel.Application excelApp = new Excel.Application();

            // Open the Excel workbook containing the table you want to update
            Excel.Workbook workbook = excelApp.Workbooks.Open(this.antalPCFile);

            // Get the worksheet containing the table
            Excel.Worksheet worksheet = workbook.Worksheets[this.antalPCSheet];

            // Get a reference to the table
            Excel.ListObject table = worksheet.ListObjects["Innevarandemånadtabell"];

            // Loop through the table rows
            foreach (Excel.ListRow row in table.ListRows)
            {
                // Get the value from the first column of the current row
                string dictkey = row.Range[4].Value.ToString();
                string computerType = row.Range[7].Value.ToString();

                // Check if the value exists in the dictionary
                if (computerType == "Extern dator")
                {
                    int tmpCount = Convert.ToInt32(row.Range[6].Value);
                    try
                    {
                        extCompTotal += tmpCount;
                    }
                    catch { }
                }
                if (dictionary.ContainsKey(dictkey) && computerType != "Extern dator")
                {
                    // Get the value from the dictionary
                    int dictValue = dictionary[dictkey];
                    int antalPCValue = antalPCdict[dictkey];

                    int pcDiff = dictValue - antalPCValue;

                    totalCapioComp += dictValue;

                    // Update the value in the column with amount of computers of the current row
                    Excel.Range cellToUpdate = row.Range[6];
                    cellToUpdate.Value = dictValue;
                    cellToUpdate.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                    cellToUpdate = row.Range[10];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2";

                    cellToUpdate = row.Range[11];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2*12";

                    cellToUpdate = row.Range[8];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2";

                    cellToUpdate = row.Range[12];
                    cellToUpdate.Value = pcDiff;
                    if(pcDiff > 10)
                    {
                        cellToUpdate.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    if(pcDiff < -10)
                    {
                        cellToUpdate.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                }
                // Checks if value not exists in dictionary, == not in the ivantiFile
                if (!dictionary.ContainsKey(dictkey) && computerType != "Extern dator")
                {
                    Excel.Range cellToUpdate = row.Range[6];
                    cellToUpdate.Value = 0; //?
                    cellToUpdate.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                    cellToUpdate = row.Range[10];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2";

                    cellToUpdate = row.Range[11];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2*12";

                    cellToUpdate = row.Range[8];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2";
                }
                if(computerType ==  "Extern dator")
                {
                    Excel.Range cellToUpdate = row.Range[6];
                    cellToUpdate = row.Range[10];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$3";

                    cellToUpdate = row.Range[11];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$3*12";

                    cellToUpdate = row.Range[8];
                    cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$3";
                }
            }
            // Add missing key to table
            foreach (var kvp in diff)
            {
                Excel.ListRow newRow = table.ListRows.Add();
                newRow.Range[4].Value = kvp.Key;
                newRow.Range[6].Value = kvp.Value;

                totalCapioComp += kvp.Value;

                newRow.Range[7].Value = "Capio dator";
                Excel.Range cellToUpdate = newRow.Range[6];
                cellToUpdate.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                cellToUpdate = newRow.Range[10];
                cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2";

                cellToUpdate = newRow.Range[11];
                cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2*12";

                cellToUpdate = newRow.Range[8];
                cellToUpdate.Value = "=[@[Antal datorer]]*Total!$C$2";

                cellToUpdate = newRow.Range[12];
                cellToUpdate.Value = 0;
            }
            //Update total sheet
            UpdateTotalSheet(workbook, totalCapioComp, extCompTotal);

            // Save the workbook
            workbook.Save();

            excelApp.Visible = true;
        }
        public string CheckIfDictionaryKeyExists(Dictionary<string, int> dictionary, string value)
        {
            string kst = null;
            bool check = false;
            foreach (KeyValuePair<string, int> entry in dictionary)
            {
                kst = entry.Key;
                if(value == entry.Key)
                {
                    check = true;
                    break;
                }
            }
            if(check)
            {
                return null;
            }
            return kst;
        }
        public void UpdateTotalSheet(Excel.Workbook workbook, int totalCapioComp, int extCompTotal)
        {
            // Get the worksheet containing the table
            Excel.Worksheet worksheet = workbook.Worksheets[this.totalSheet];

            // Get a reference to the table
            Excel.ListObject table = worksheet.ListObjects["Totaltabell"];

            // Loop through the table rows
            foreach (Excel.ListRow row in table.ListRows)
            {
                Excel.Range cellCheck = row.Range[1];

                if (cellCheck.Value.ToString() == "Capio dator")
                {
                    Excel.Range cellToUpdate = row.Range[4];
                    cellToUpdate.Value = totalCapioComp;
                }
                if (cellCheck.Value.ToString() == "Extern dator")
                {
                    Excel.Range cellToUpdate = row.Range[4];
                    cellToUpdate.Value = extCompTotal;
                }
            }
        }
    } 
}
