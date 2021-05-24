on error resume next
if WScript.Arguments.length =0 Then
 
Set objShell = CreateObject("Shell.Application")
  
objShell.ShellExecute "wscript.exe", Chr(34) & _
  WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
else Const HKEY_CLASSES_ROOT = &H80000000
strComputer = "."  
Set StdOut = WScript.StdOut
Set objRegistry = GetObject("winmgmts:\\" & _   
strComputer & "\root\default:StdRegProv")  
strKeyPath = "CLSID\{00000000-0000-0000-0000-000000000012}\InProcServer32" 
strValueName = "ThreadingModel"   
strValue = "Apartment" 
objRegistry.GetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
If IsNull(strValue) Then 
intAnswer = Msgbox("你是否要注册“显示/隐藏所有的文件及扩展名（含系统文件）”右键菜单？"&vbCrlf&""&vbCrlf&"注意事项："&vbCrlf&""&vbCrlf&"1.目前本扩展功能完美支持 Vista/Win 7/Win 8 系统，XP 用户请不要使用！Win 10 用户请事先卸载 “入门”APP 后方可正常使用！"&vbCrlf&"2.注册过程绝对安全、无毒，不会破坏系统的正常功能，但是部分杀毒软件会报警提示，请直接忽略提示允许操作，否则会打断正常注册流程，从而导致注册失败，无法使用！"&vbCrlf&"3.本扩展功能依赖同目录下的“SuperHidden.vbe”文件，请不要对其更名、移动、删除，否则会导致功能失效！"&vbCrlf&"4.对于开启 UAC（用户账户控制）的用户，系统弹出对话框提示“是否确认更改”时，请点击“是”按钮继续，否则会导致功能失效！",vbYesNo+vbExclamation, "系统提示")
If intAnswer = vbNo Then 
wscript.quit
else VBSPath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_     
strComputer & "\root\default:StdRegProv")  
strKeyPath = "CLSID\{00000000-0000-0000-0000-000000000012}"   
oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
strKeyPath = "CLSID\{00000000-0000-0000-0000-000000000012}\InProcServer32"
oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
strValueName = ""
strValue = "%SystemRoot%\system32\shdocvw.dll"   
oReg.SetExpandedStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
strValueName = "ThreadingModel"   
strValue = "Apartment"
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
strKeyPath = "CLSID\{00000000-0000-0000-0000-000000000012}\Instance"   
oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
strValueName = "CLSID"   
strValue = "{3f454f0e-42ae-4d7c-8ea3-328250d6e272}"   
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
strKeyPath = "CLSID\{00000000-0000-0000-0000-000000000012}\Instance\InitPropertyBag"   
oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
strValueName = "CLSID"   
strValue = "{13709620-C279-11CE-A49E-444553540000}"    
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue    
strValueName = "method"   
strValue = "ShellExecute"   
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
strValueName = "Param1"   
strValue = VBSPath & "\SuperHidden.vbe"
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")
if WSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden") = 1 then
strValueName = "command"   
strValue = "隐藏所有的文件及扩展名（含系统）"   
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
else strValueName = "command"   
strValue = "显示所有的文件及扩展名（含系统）"   
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
end if
strKeyPath = "Directory\Background\shellex\ContextMenuHandlers\SuperHidden"   
oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
strValueName = ""   
strValue ="{00000000-0000-0000-0000-000000000012}"
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
msgbox "注册成功！",vbOKOnly+vbInformation,"系统提示"
end if
Else intAnswer = Msgbox("你是否要移除“显示/隐藏所有的文件及扩展名（含系统文件）”右键菜单？",vbYesNo+vbExclamation, "系统提示")
If intAnswer = vbNo Then 
wscript.quit
else Set oshell = WScript.CreateObject ("WSCript.shell")
oshell.run "cmd.exe /c reg delete HKCR\CLSID\{00000000-0000-0000-0000-000000000012} /f",0  
oshell.run "cmd.exe /c reg delete HKCR\Directory\Background\shellex\ContextMenuHandlers\SuperHidden /f",0
msgbox "移除成功！",vbOKOnly+vbInformation,"系统提示"
End If 
end if
end if