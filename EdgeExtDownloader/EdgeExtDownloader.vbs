ext_id=inputbox("Microsoft Edge ���������" &vbCrlf&"" &vbCrlf&"��С���߿��Դ�ֱ�Ӵ�΢��ٷ�����̵����ز����.crx���������߲����ק����չ����ҳ�棨edge://extensions/������ֱ�Ӱ�װ��"& vbCrLf&""& vbCrLf& "��������ID�ţ�" ,"Microsoft Edge ���������")
If ext_id=vbEmpty Then
wscript.quit
Else currentpath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
Set Post = CreateObject("Msxml2.ServerXMLHTTP")
Set Shell = CreateObject("Wscript.Shell")
extid=ext_id
msgbox "������ں�̨�����У����Ժ�" &vbCrlf&"" &vbCrlf& "���غ���ļ������浽��ǰĿ¼�¡�" ,vbOKOnly+vbInformation,"Microsoft Edge ���������" 
Post.Open "GET","http://jimmyjlnu.herokuapp.com/proxy/https://edge.microsoft.com/extensionwebstorebase/v1/crx?response=redirect&prod=chromiumcrx&prodchannel=&x=id%3D"&extid&"%26installsource%3Dondemand%26uc",0
Post.Send
Set aGet = CreateObject("ADODB.Stream")
aGet.Mode = 3
aGet.Type = 1
aGet.Open() 
aGet.Write(Post.responseBody)
aGet.SaveToFile currentpath &"\"&extid&".crx",2
msgbox "������سɹ���" &vbCrlf&"" &vbCrlf& "���غ���ļ��Ѿ����浽��ǰĿ¼�¡�" ,vbOKOnly+vbInformation,"Microsoft Edge ���������" 
end if