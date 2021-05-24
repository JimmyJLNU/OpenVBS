on error resume next
if WScript.Arguments.length =0 Then
 
Set objShell = CreateObject("Shell.Application")
  
objShell.ShellExecute "wscript.exe", Chr(34) & _
  WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
else test = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
Dim OperationRegistry
Set OperationRegistry=WScript.CreateObject("WScript.Shell")
OperationRegistry.RegWrite "HKCR\.swf\","ShockwaveFlash.ShockwaveFlash"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\","Shockwave Flash Object"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\DefaultIcon\",test&"\SAFlashPlayer.exe,5"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\CLSID","{D27CDB6E-AE6D-11cf-96B8-444553540000}"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\CurVer","ShockwaveFlash.ShockwaveFlash.18"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\Shell\","Open"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\Shell\Open\","打开(&O)"
OperationRegistry.RegWrite "HKCR\ShockwaveFlash.ShockwaveFlash\Shell\Open\Command\",test&"\SAFlashPlayer.exe ""%1"""
MsgBox "文件关联注册添加成功！",vbInformation+vbOKOnly,"系统提示"
End if