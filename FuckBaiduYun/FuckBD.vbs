on error resume next
Set argv = WScript.Arguments
If argv.Count = 0 Then 
intAnswer = Msgbox("欢迎您使用百度网盘防资源和谐破解工具！" &vbCrlf&" "&vbCrlf& "点击“确定”浏览并选择需要做“特殊处理”的文件，即可瞬间完成破解！亦可在文件夹中直接拖拽需要做“特殊处理”的文件到本工具上！"&vbCrlf& ""&vbCrlf& "特别提醒："&vbCrlf& ""&vbCrlf&"本工具主要用于处理文档、压缩包以及软件，尽量不要处理用于影音资源，否则一旦文件损坏，后果自负！"&vbCrlf&""&vbCrlf&"处理影音资源请先使用百度网盘防资源和谐破解工具（影音专用版），如果无效，再使用本工具处理！",vbOKCancel+vbExclamation, "百度网盘防资源和谐破解工具（通用版）")
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
cop.copy(a & ".bd") 
Set FileTemp = FSO.OpenTextFile(a & ".bd",ForAppending,TristateFalse)  
FileTemp.Write "FuckBD" 
MsgBox "破解成功！已经被“特殊处理”的文件保存到以下目录：" & vbCrLf & vbCrLf &a & ".bd"&vbCrlf& ""&vbCrlf&"（为防止被百度和谐，新生成的文件后缀为“.bd”，请自行修改为正确格式后再使用！）" ,vbOKOnly+vbInformation,"系统提示" 
end if
else Set fso = CreateObject("Scripting.FileSystemObject")
set cop=fso.getfile(argv(0))
cop.copy(argv(0) & ".bd") 
Set FileTemp = FSO.OpenTextFile(argv(0) & ".bd",ForAppending,TristateFalse)  
FileTemp.Write "FuckBD" 
MsgBox "破解成功！已经被“特殊处理”的文件保存到以下目录：" & vbCrLf & vbCrLf &argv(0) & ".bd"&vbCrlf& ""&vbCrlf&"（为防止被百度和谐，新生成的文件后缀为“.bd”，请自行修改为正确格式后再使用！）" ,vbOKOnly+vbInformation,"系统提示" 
end if
