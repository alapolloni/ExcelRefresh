#ExcelRefresh.exe (XLRefresh in C#)
Command line utility to refresh Excel documents and their external connections, query tables and pivot tables.

[Download ExcelRefresh.exe](https://github.com/alapolloni/ExcelRefresh/blob/master/ExcelRefresh.exe?raw=true)

Example usage.  Use Window's Task Scheduler to run a .bat which updates Excel files and then copies them to a public location.  

As of Excel 2013, the query table method seems to have been replaced with an external connection method.  Leaving the query table just in case.

  -f, --file            Required. Input file to be processed.                          
                                                                                       
  -m, --Macros          The worksheet macros to run. Example: -m sheet1.someMacro (sheet2.otherMacro)                           
                                                                                       
  -d, --verbose         (Default: False) Prints all messages to standard output.                                                        
                                                                                       
  -v, --visible         (Default: False) Shows Excel while update is running.                                   
                                                                                       
  -p, --pivot-tables    (Default: False) Refresh Pivot-tables.                         
                                                                                       
  -q, --query-tables    (Default: False) Refresh query-tables. (Pre Excel 2013)        
                                                                                       
  -c, --connections     (Default: False) Refresh External connections. (Excel 2013)                                                          

#TODO/Notes#
##DONE##
 - The "DO you want to Save keeps popping up intermittently".  (DONE)
   - __Caused if an EXCEL.exe process is already running.__  Need to manually check in task manager and kill prior to running.
   - http://stackoverflow.com/questions/2123158/c-sharp-winforms-how-to-load-an-image-then-wait-a-few-seconds-then-play-a-mp3
  - which suggests application.doevents()
  - er...also added a Excel.Close(true). 
 - Add run --macros option (DONE)
 - ilmerge working (instead of distributing command.dll with .exe) (DONE)
    - from Debug dir
     - "C:\Program Files (x86)\Microsoft\ILMerge\ilmerge.exe" XLRefreshC.exe /out:ExcelRefresh.exe /target:exe CommandLine.dll /targetplatform:"v4,C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0
    - make a makefile
 - upload to github.com
 
##TODO##
 - Add do -a/--all option.

##Contribution
Based on Perl program/library originally written by [CTBROWN](http://cpansearch.perl.org/src/CTBROWN/Win32-Excel-Refresh-0.02/extras/XLRefresh.pl) 
 
