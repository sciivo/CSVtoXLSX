using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace CSVtoXLSX
{
    class Program
    {
        static void Main(string[] args)
        {
            //Set default variables
            string dCSV = "";
            string dXLSX = "";
            string dLog = "";
            string dMerge = "";

            try
            {
                //Process arguments
                if (args.Count() == 4)
                {
                    dCSV = args[0];
                    dXLSX = args[1];
                    dLog = args[2];
                    dMerge = args[3].ToLower();

                    //Determine destination file existence
                    bool fileExists = File.Exists(dXLSX);

                    //Error checking
                    if (dMerge != "true" || fileExists == false)
                    {
                        dMerge = "false";
                    }

                    //Variables
                    Application excelApp = new Application();
                    bool merge = Boolean.Parse(dMerge);

                    //Open CSV
                    Workbook csv = excelApp.Workbooks.Open(dCSV);
                    Worksheet ws = csv.ActiveSheet;

                    //Get last row and column
                    int lastRow = ws.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
                    int lastCol = ws.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;

                    //Delete uneeded rows
                    ws.Rows[lastRow].Delete();
                    ws.Rows[lastRow - 1].Delete();
                    ws.Rows[2].Delete();

                    //Formatting
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, lastCol]].Font.Bold = true;
                    ws.Range[ws.Cells[1, 1], ws.Cells[lastRow, lastCol]].Columns.AutoFit();

                    //Debugging
                    Console.WriteLine("Last Row:      " + lastRow);
                    Console.WriteLine("Last Column:   " + lastCol);
                    Console.WriteLine("Target Exists: " + fileExists);
                    Console.WriteLine("Merge Files:   " + merge);

                    //Merge or not, then save as XLSX
                    if (fileExists == true && merge == true)
                    {
                        Workbook mergeXLSX = excelApp.Workbooks.Open(dXLSX);
                        ws.Move(Type.Missing, mergeXLSX.Worksheets[mergeXLSX.Worksheets.Count]);

                        mergeXLSX.Save();
                    }
                    else
                    {
                        excelApp.DisplayAlerts = false;
                        csv.SaveAs(dXLSX, XlFileFormat.xlOpenXMLWorkbook);
                        excelApp.DisplayAlerts = true;
                    }

                    excelApp.Quit();
                }
            }
            catch (Exception e)
            {
                //Create timestamp
                DateTime dt = DateTime.Now;
                string timestamp = dt.ToString("[yyyy/MM/dd HH:mm:ss]");

                //Create error string
                StringBuilder sb = new StringBuilder();
                sb.Append(timestamp + ": ");
                sb.Append(e.ToString());

                //Append to error log
                StreamWriter sw = new StreamWriter(dLog, true);
                sw.WriteLine("============================================================");
                sw.WriteLine("============= Args: CSV, XLSX, Log File, Merge =============");
                sw.WriteLine("============================================================");
                sw.WriteLine(sb);
                sw.Close();
            }
        }
    }
}