on error resume next
intAnswer = Msgbox("��ӭ��ʹ���ı���ȡС���ߣ�"&vbCrlf&" "&vbCrlf&"�����߿��Ը��ݹؼ��ʣ���ȡ����ͬ�е��ı����ݣ���������Զ������������û�ֱ���˵��ز鿴�������й���Ϣ���Ӷ��������߹���Ч�ʣ�" &vbCrlf&" "&vbCrlf&"�����ȷ������ťѡȡĿ���ļ���"&vbCrlf&" "&vbCrlf&"��ܰ��ʾ��"&vbCrlf&"����VBS�ű��༭�����Ŀ͹����ƣ������߲�֧�� Unicode ���룡"&vbCrlf&"ֻ֧�� UTF-8 �����Ӣ���ַ�����֧�������ַ���"&vbCrlf&"�����û��������������Ϊ���룬������ѡ��лл������⣡" ,vbOKCancel+vbExclamation, "�ı���ȡС����")
If intAnswer = vbCancel Then 
wscript.quit
Else Function BrowseForFile()
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
Dim tempName : tempName = fso.GetTempName()
Dim tempFile : Set tempFile = tempFolder.CreateTextFile(tempName & ".hta")
tempFile.Write _
        "<html>" & _
        "  <head>" & _
        "    <title>Browse</title>" & _
        "  </head>" & _
        "  <body>" & _
        "    <input type='file' id='f'>" & _
        "    <script type='text/javascript'>" & _
        "      var f = document.getElementById('f');" & _
        "      f.click();" & _
        "      var shell = new ActiveXObject('WScript.Shell');" & _
        "      shell.RegWrite('HKEY_CURRENT_USER\\Software\\MsgResp', f.value);" & _
        "      window.close();" & _
        "    </script>" & _
        "  </body>" & _
        "</html>"
tempFile.Close
shell.Run tempFolder & "\" & tempName & ".hta", 1, true
BrowseForFile = shell.RegRead("HKEY_CURRENT_USER\SOFTWARE\MsgResp")
shell.RegDelete "HKEY_CURRENT_USER\SOFTWARE\MsgResp"
End Function
file=BrowseForFile
dim read,str,f,m
f=inputbox("��������Ҫ���ҵĹؼ��֣�" ,"�ı���ȡС����") 
set open=createobject("scripting.filesystemobject")
m=split(file,"\")
str=m(ubound(m))
private Function TrimT(ByVal Text)
If Right(Text, 2) = vbcrlf Then TrimT = Left(Text, Len(Text) - 2)
End Function
private Function findstr(ByVal text,ByVal find)
For i = 1 To Len(Text)
If Mid(Text, i, Len(Find)) = Find Then j = 1: Exit For Else j = 0
Next
findstr=j
End Function
read=open.opentextfile(file).readall
for each s in split(read,vbcrlf)
if findstr(s,f)=1 then h=h&s&vbcrlf
next
h=trimt(h)
open.createtextfile("result.txt").write h
MsgBox "��ȡ��ɣ�" &vbCrlf&"" &vbCrlf& "��Ӧ���ı������Ѿ����浽��ǰĿ¼�µġ�result.txt���ļ��У�" ,vbOKOnly+vbInformation,"��ȡ�ɹ�" 
end if
