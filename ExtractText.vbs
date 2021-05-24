on error resume next
intAnswer = Msgbox("欢迎您使用文本提取小工具！"&vbCrlf&" "&vbCrlf&"本工具可以根据关键词，提取与其同行的文本内容，并将结果自动导出，方便用户直截了当地查看并整理有关信息，从而极大地提高工作效率！" &vbCrlf&" "&vbCrlf&"点击“确定”按钮选取目标文件！"&vbCrlf&" "&vbCrlf&"温馨提示："&vbCrlf&"由于VBS脚本编辑环境的客观限制，本工具不支持 Unicode 编码！"&vbCrlf&"只支持 UTF-8 编码的英文字符，不支持中文字符。"&vbCrlf&"否则会没有输出结果或中文为乱码，请慎重选择！谢谢您的理解！" ,vbOKCancel+vbExclamation, "文本提取小工具")
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
f=inputbox("请输入需要查找的关键字：" ,"文本提取小工具") 
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
MsgBox "提取完成！" &vbCrlf&"" &vbCrlf& "相应的文本内容已经保存到当前目录下的“result.txt”文件中！" ,vbOKOnly+vbInformation,"提取成功" 
end if
