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
intAnswer = Msgbox("���Ƿ�Ҫע�ᡰ��ʾ/�������е��ļ�����չ������ϵͳ�ļ������Ҽ��˵���"&vbCrlf&""&vbCrlf&"ע�����"&vbCrlf&""&vbCrlf&"1.Ŀǰ����չ��������֧�� Vista/Win 7/Win 8 ϵͳ��XP �û��벻Ҫʹ�ã�Win 10 �û�������ж�� �����š�APP �󷽿�����ʹ�ã�"&vbCrlf&"2.ע����̾��԰�ȫ���޶��������ƻ�ϵͳ���������ܣ����ǲ���ɱ������ᱨ����ʾ����ֱ�Ӻ�����ʾ��������������������ע�����̣��Ӷ�����ע��ʧ�ܣ��޷�ʹ�ã�"&vbCrlf&"3.����չ��������ͬĿ¼�µġ�SuperHidden.vbe���ļ����벻Ҫ����������ƶ���ɾ��������ᵼ�¹���ʧЧ��"&vbCrlf&"4.���ڿ��� UAC���û��˻����ƣ����û���ϵͳ�����Ի�����ʾ���Ƿ�ȷ�ϸ��ġ�ʱ���������ǡ���ť����������ᵼ�¹���ʧЧ��",vbYesNo+vbExclamation, "ϵͳ��ʾ")
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
strValue = "�������е��ļ�����չ������ϵͳ��"   
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
else strValueName = "command"   
strValue = "��ʾ���е��ļ�����չ������ϵͳ��"   
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
end if
strKeyPath = "Directory\Background\shellex\ContextMenuHandlers\SuperHidden"   
oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
strValueName = ""   
strValue ="{00000000-0000-0000-0000-000000000012}"
oReg.SetStringValue HKEY_CLASSES_ROOT,strKeyPath,strValueName,strValue
msgbox "ע��ɹ���",vbOKOnly+vbInformation,"ϵͳ��ʾ"
end if
Else intAnswer = Msgbox("���Ƿ�Ҫ�Ƴ�����ʾ/�������е��ļ�����չ������ϵͳ�ļ������Ҽ��˵���",vbYesNo+vbExclamation, "ϵͳ��ʾ")
If intAnswer = vbNo Then 
wscript.quit
else Set oshell = WScript.CreateObject ("WSCript.shell")
oshell.run "cmd.exe /c reg delete HKCR\CLSID\{00000000-0000-0000-0000-000000000012} /f",0  
oshell.run "cmd.exe /c reg delete HKCR\Directory\Background\shellex\ContextMenuHandlers\SuperHidden /f",0
msgbox "�Ƴ��ɹ���",vbOKOnly+vbInformation,"ϵͳ��ʾ"
End If 
end if
end if