on error resume next
if WScript.Arguments.length =0 Then
 
Set objShell = CreateObject("Shell.Application")
  
objShell.ShellExecute "wscript.exe", Chr(34) & _
  WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
else Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")
if WSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden") = 1 then
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowSuperHidden", "0", "REG_DWORD"
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden", "2", "REG_DWORD"
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", "1", "REG_DWORD"
WSHShell.RegWrite "HKCR\CLSID\{00000000-0000-0000-0000-000000000012}\Instance\InitPropertyBag\command", "显示所有的文件及扩展名（含系统）", "REG_SZ"
WSHShell.SendKeys "{F5}+{F10}e"
else
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowSuperHidden", "1", "REG_DWORD"
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden", "1", "REG_DWORD"
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", "0", "REG_DWORD"
WSHShell.RegWrite "HKCR\CLSID\{00000000-0000-0000-0000-000000000012}\Instance\InitPropertyBag\command", "隐藏所有的文件及扩展名（含系统）", "REG_SZ"
WSHShell.SendKeys "{F5}+{F10}e"
end if
end if
