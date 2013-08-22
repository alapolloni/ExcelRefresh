using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using CommandLine;
using CommandLine.Text;
using System.IO;

namespace XLRefresh
{
    class Program
    {
        private static void refreshQueryTables(Excel.Workbook theWorkbook)
        {
            Console.WriteLine("WorkSheets:"); 
            Excel.Sheets oSheets = (Excel.Sheets)theWorkbook.Worksheets;
         
            foreach (Excel.Worksheet oWorkSheet in oSheets)
            {
               Console.WriteLine(" {0}",oWorkSheet.Name);
               foreach (Excel.QueryTable qt in oWorkSheet.QueryTables)
                {                   
                    Console.WriteLine("--qt:{0}",qt.Name);
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
        
        private static void refreshPivots(Excel.Workbook theWorkbook)
        {

            Console.WriteLine("WorkSheets:"); 
            Excel.Sheets oSheets = (Excel.Sheets)theWorkbook.Worksheets;

            foreach (Excel.Worksheet oWorkSheet in oSheets)
            {
                Console.WriteLine(" {0}",oWorkSheet.Name);
                Excel.PivotTables pivotTables1 =
                    (Excel.PivotTables)oWorkSheet.PivotTables();

                if (pivotTables1.Count > 0)
                {
                    for (int i = 1; i <= pivotTables1.Count; i++)
                    {
                        Console.WriteLine("  PivoteTable Refresh: {0}", pivotTables1.Item(i).Name);
                        pivotTables1.Item(i).RefreshTable();
                        
                    }
                }
                else
                {
                    Console.WriteLine("  !This worksheet contains no pivot tables.");
                }

            }

        }

        private static void refreshConnection(Excel.Workbook theWorkbook)
        {
            foreach (Microsoft.Office.Interop.Excel.WorkbookConnection i in theWorkbook.Connections)
            {
                Console.WriteLine("Connection refresh: {0}",i.Name);
                i.OLEDBConnection.BackgroundQuery = false;
                i.Refresh();
            }
        }


        static void Main(string[] args)
        {
            //http://commandline.codeplex.com/
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                // Values are available here
                if (options.Verbose) Console.WriteLine("Filename: {0}", options.InputFile);
                //Items?
                if (options.Verbose)
                {
                    Console.WriteLine("Items Count: {0}", options.Items.Count.ToString());
                    options.Items.ToList().ForEach(i => Console.Write("{0}\t", i));
                }
            }
            else
            {
                Console.WriteLine("required options not specified ... quiting.");
                return;
            }

            string txtLocation = Path.GetFullPath(options.InputFile);
            if (options.Verbose) Console.WriteLine("Input File Full Path: {0}", txtLocation);
            if (! File.Exists(txtLocation))
            {
                Console.WriteLine("Input File does not exist: {0}", txtLocation);
                return;
            }
           
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

            
            if (options.Visible) excel.Visible = true;

            if (options.All)
            {
                theWorkbook.RefreshAll();
            }
            else
            {
                if (options.Querytables) { refreshQueryTables(theWorkbook); }
                if (options.Connections) { refreshConnection(theWorkbook); }
                if (options.Pivottables) { refreshPivots(theWorkbook); }
            }
            //
            // To test:
            // XLRefreshC.exe -d  -f ..\..\..\Book1.xlsm -m  sheet1.showMessage
            //
            if (options.MacrosToRun != null)
            {
                if (options.Verbose) Console.WriteLine("Macro Count:{0}", options.MacrosToRun.Count().ToString());
                foreach (string macro in options.MacrosToRun)
                {
                    if (options.Verbose) Console.WriteLine("Macro :{0}", macro);
                    excel.Run(macro);
                }
            }
            else
            {
                if (options.Verbose) Console.WriteLine("no macro.");
            }


            if (options.Verbose) Console.WriteLine("shut it down!");
            excel.Calculate();
            if (options.Verbose) Console.WriteLine("Excel calculated");
            theWorkbook.Save();
            if (options.Verbose) Console.WriteLine("Workbook saved");
            theWorkbook.Close(true);
            if (options.Verbose) Console.WriteLine("Workbook closed");
            excel.Quit();
            if (options.Verbose) Console.WriteLine("Excel Quit");
            //Console.WriteLine("Press any key to close...");
            //Console.ReadLine();
            return;
        }
    }

// Define a class to receive parsed values
class Options {
  [Option('f', "file", Required = true,
    HelpText = "Input file to be processed.")]
  public string InputFile { get; set; }

  [OptionArray('m', "Macros", HelpText = "The worksheet macros to run. Example: -m sheet1.someMacro (sheet2.otherMacro)")]
    public string[] MacrosToRun { get; set; }

  [Option('d', "verbose", DefaultValue = false,
      HelpText = "Prints all messages to standard output.")]
  public bool Verbose { get; set; }

  [Option('v', "visible", DefaultValue = false,
    HelpText = "Shows Excel while update is running.")]
  public bool Visible { get; set; }

  [Option('p', "pivot-tables", DefaultValue = false,
      HelpText = "Refresh Pivot-tables.")]
  public bool Pivottables { get; set; }

  [Option('q', "query-tables", DefaultValue = false,
  HelpText = "Refresh query-tables. (Pre Excel 2013)")]
  public bool Querytables { get; set; }

  [Option('c', "connections", DefaultValue = false,
        HelpText = "Refresh External connections. (Excel 2013)")]
  public bool Connections { get; set; }

  [Option('a', "all", DefaultValue = false,
        HelpText = "Refreshes all external data ranges and PivotTable reports in the specified workbook.")]
  public bool All { get; set; }

  [ValueList(typeof(List<string>), MaximumElements = 6)]
    public IList<string> Items { get; set; }

  [ParserState]
  public IParserState LastParserState { get; set; }

  [HelpOption]
  public string GetUsage() {
    return HelpText.AutoBuild(this,
      (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
  }
}






}



