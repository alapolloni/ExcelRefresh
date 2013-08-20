XLRefreshC.exe (XLRefresh in C#)

 - The "DO you want to Save keeps popping up intermittently".  Test by adding sleeps...i just guessing.
   - !! Caused if an EXCEL.exe process is already running.  Need to manually check in task manager and kill prior to running.
   - http://stackoverflow.com/questions/2123158/c-sharp-winforms-how-to-load-an-image-then-wait-a-few-seconds-then-play-a-mp3
  - which suggests application.doevents()
  - er...also added a Excel.Close(true). 
 - Add run --macros option (DONE)
 - TODO
  - ilmerge get working (now need to distribute command.dll with .exe)
    - "C:\Program Files (x86)\Microsoft\ILMerge\ilmerge.exe" XLRefreshC.exe /out:ExcelRefresh.exe /target:exe CommandLine.dll /targetplatform:"v4,C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0
  - upload to github.com
  - Add do --all option.
 
