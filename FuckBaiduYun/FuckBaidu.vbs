on error resume next
Set argv = WScript.Arguments
If argv.Count = 0 Then 
intAnswer = Msgbox("��ӭ��ʹ�ðٶ����̷���Դ��г�ƽ⹤�ߣ�Ӱ��ר�ð棩��" &vbCrlf&" "&vbCrlf& "�����ȷ���������ѡ����Ҫ�������⴦������Դ������˲������ƽ⣡������ļ�����ֱ����ק��Ҫ�������⴦������Դ���������ϣ�"&vbCrlf& ""&vbCrlf& "ע�����" &vbCrlf&"������ֻ�����ڴ���Ӱ�ӡ���������Դ���мɲ�Ҫ�������ڴ����ĵ���ѹ����������ȣ�����һ���ļ��𻵣�����Ը���"&vbCrlf&""&vbCrlf&"���������ļ���ʹ�ðٶ����̷���Դ��г�ƽ⹤�ߣ�ͨ�ð棩��"&vbCrlf& ""&vbCrlf& "�ر����ѣ�" &vbCrlf& "���ڱ����������ޣ�������ֻ�ܴ����ָ�ʽӰ����Դ�����ʧЧ��ʹ�ðٶ����̷���Դ��г�ƽ⹤�ߣ�ͨ�ð棩��" ,vbOKCancel+vbExclamation, "�ٶ����̷���Դ��г�ƽ⹤�ߣ�Ӱ��ר�ð棩")
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
a = BrowseForFile
Set fso = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8 
set cop=fso.getfile(a)
cop.copy(a & ".movie") 
Dim objWMP, objSong
Set objWMP = CreateObject("WMPlayer.OCX")
Set objSong = objWMP.newMedia(a & ".movie")
objSong.setItemInfo "Description", "FuckBaidu"
Set objSong = Nothing
Set objWMP = Nothing
MsgBox "�ƽ�ɹ����Ѿ��������⴦�����ļ����浽����Ŀ¼��" & vbCrLf & vbCrLf &a & ".movie"&vbCrlf& ""&vbCrlf&"��Ϊ��ֹ���ٶȺ�г�������ɵ��ļ���׺Ϊ��.movie�����������޸�Ϊ��ȷ��ʽ���ٲ��ţ���" &vbCrlf& ""&vbCrlf& "���ڱ����������ޣ�������ֻ�ܴ����ָ�ʽӰ����Դ�������Ч��ʹ�ðٶ����̷���Դ��г�ƽ⹤�ߣ�ͨ�ð棩��" ,vbOKOnly+vbInformation,"ϵͳ��ʾ" 
end if
else Set fso = CreateObject("Scripting.FileSystemObject")
set cop=fso.getfile(argv(0))
cop.copy(argv(0) & ".movie") 
Set objWMP = CreateObject("WMPlayer.OCX")
Set objSong = objWMP.newMedia(argv(0) & ".movie")
objSong.setItemInfo "Description", "FuckBaidu"
Set objSong = Nothing
Set objWMP = Nothing
MsgBox "�ƽ�ɹ����Ѿ��������⴦�����ļ����浽����Ŀ¼��" & vbCrLf & vbCrLf &argv(0) & ".movie"&vbCrlf& ""&vbCrlf&"��Ϊ��ֹ���ٶȺ�г�������ɵ��ļ���׺Ϊ��.movie�����������޸�Ϊ��ȷ��ʽ���ٲ��ţ���"&vbCrlf& ""&vbCrlf& "���ڱ����������ޣ�������ֻ�ܴ����ָ�ʽӰ����Դ�������Ч��ʹ�ðٶ����̷���Դ��г�ƽ⹤�ߣ�ͨ�ð棩��" ,vbOKOnly+vbInformation,"ϵͳ��ʾ" 
end if