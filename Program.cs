using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;

namespace XLRefresh
{
    class Program
    {
        static void useQueryTables(Excel.Workbook theWorkbook)
        {
            Console.WriteLine("wb"); 
            Excel.Sheets oSheets = (Excel.Sheets)theWorkbook.Worksheets;
         
            foreach (Excel.Worksheet oWorkSheet in oSheets)
            {
               Console.WriteLine(oWorkSheet.Name);
               foreach (Excel.QueryTable qt in oWorkSheet.QueryTables)
                {                   
                    Console.WriteLine(" qt");
                    qt.EnableRefresh = true;
                    qt.FieldNames = false;
                    qt.RowNumbers = false;
                    qt.SavePassword = false;
                    qt.SaveData = true;
                    qt.PreserveColumnInfo = true;
                    qt.Refresh(false);
                }
            }
                      
            return;            
        }



        static void Main(string[] args)
        {
            string txtLocation = "C:/Users/aapollon/Documents/TPT/test xlrefresh/Sherpa_test - Copy.xlsx";
            object _missingValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook theWorkbook = excel.Workbooks.Open(txtLocation,
                                                            _missingValue,
                                                            false,
                                                            _missingValue,
                                                            _missingValue,
                                                            _missingValue,
                                                            true,
                                                            _missingValue,
                                                            _missingValue,
                                                            true,
                                                            _missingValue,
                                                            _missingValue,
                                                            _missingValue);
            excel.Visible=true;
            //doesn't work//useQueryTables(theWorkbook);
            
            //useConnection(theWorkbook);
            refreshPivots(theWorkbook);
            theWorkbook.Save();
            theWorkbook.Close(); 
            excel.Quit();
            Console.WriteLine("Press any key to close...");
            Console.ReadLine();
            return;
                
        }

        private static void refreshPivots(Excel.Workbook theWorkbook)
        {
            
            Console.WriteLine("wb"); 
            Excel.Sheets oSheets = (Excel.Sheets)theWorkbook.Worksheets;

            foreach (Excel.Worksheet oWorkSheet in oSheets)
            {
                Console.WriteLine(oWorkSheet.Name);
                Excel.PivotTables pivotTables1 =
                    (Excel.PivotTables)oWorkSheet.PivotTables();

                if (pivotTables1.Count > 0)
                {
                    for (int i = 1; i <= pivotTables1.Count; i++)
                    {
                        Console.WriteLine("PT update");
                        pivotTables1.Item(i).RefreshTable();
                    }
                }
                else
                {
                    Console.WriteLine("This workbook contains no pivot tables.");
                }

            }

        }

        private static void useConnection(Excel.Workbook theWorkbook)
        {
            foreach (Microsoft.Office.Interop.Excel.WorkbookConnection i in theWorkbook.Connections)
            {
                System.Console.WriteLine(i.Name);
                i.OLEDBConnection.BackgroundQuery = false;
                i.Refresh();
            }
        }

    }
}



