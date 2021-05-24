on error resume next
Do
Dim objWMIService, objProcess, colProcessList
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("select * from win32_process where name='REDAgent.EXE'")
For Each objProcess In colProcessList
objProcess.Terminate
Set oshell = WScript.CreateObject ("WSCript.shell")
oshell.run "taskkill /F /IM checkrs.exe",0
oshell.run "taskkill /F /IM rscheck.exe",0
next
Loop
