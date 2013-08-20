#this merges the CommandLine.dll and the XlRefresh.exe into one executable so that you don't have to distributre multiple files.
ExcelRefresh: bin/Debug/XlRefreshC.exe
	"C:\Program Files (x86)\Microsoft\ILMerge\ilmerge.exe" bin/Debug/XLRefreshC.exe /out:ExcelRefresh.exe /target:exe bin/Debug/CommandLine.dll /targetplatform:"v4,C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0
