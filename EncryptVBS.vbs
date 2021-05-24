on error resume next
Set argv = WScript.Arguments
If argv.Count = 0 Then
intAnswer = Msgbox("欢迎您使用 VBS 脚本加密小工具！" &vbCrlf&" "&vbCrlf& "点击“确定”浏览并选择需要加密的 VBS 脚本，即可瞬间完成加密！" &vbCrlf&""&vbCrlf& "亦可在文件夹中直接拖拽需要加密的脚本文件到本文件上！",vbOKCancel+vbExclamation, "VBS 脚本加密小工具")
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
Randomize
pass = Int(Rnd*12)+20 '异或加密有效范围20-31，所以随机生成好了。
data = fso.OpenTextFile(a).ReadAll
data = "d=" & Chr(34) & ASCdata(data) & Chr(34)
data = data & vbCrLf & ":M=Split(D):For each O in M:N=N&chr(O):Next:execute N"
data = Replace(data, " ", ",")
fso.OpenTextFile(a & "_encrypted.vbe", 2, True).Write Encoder(EncHexXorData(data))
Function EncHexXorData(data)
EncHexXorData = "x=""" & EncHexXor(data) & """:For i=1 to Len(x) Step 2:s=s&Chr(CLng(""&H""&Mid(x,i,2)) Xor " & pass & "):Next:Execute Replace(s,"","","" "")"
End Function
Function Encoder(data) '加密3
Encoder = CreateObject("Scripting.Encoder").EncodeScriptFile(".vbs", data, 0, "VBScript")
End Function
Function EncHexXor(x) '加密2
For i = 1 To Len(x)
EncHexXor = EncHexXor & Hex(Asc(Mid(x, i, 1)) Xor pass)
Next
End Function
Function ASCdata(Data) '加密1
num = Len(data)
newdata = ""
For j = 1 To num
If j = num Then
newdata = newdata&Asc(Mid(data, j, 1))
Else
newdata = newdata&Asc(Mid(data, j, 1)) & " "
End If
Next
ASCdata = newdata
End Function
MsgBox "加密成功！已经加密的脚本文件保存到以下目录：" & vbCrLf & vbCrLf &a & "_encrypted.vbe",vbOKOnly+vbInformation,"系统提示"
end if
else Set fso = CreateObject("Scripting.FileSystemObject")
Randomize
pass = Int(Rnd*12)+20 '异或加密有效范围20-31，所以随机生成好了。
data = fso.OpenTextFile(argv(0), 1).ReadAll
data = "d=" & Chr(34) & ASCdata(data) & Chr(34)
data = data & vbCrLf & ":M=Split(D):For each O in M:N=N&chr(O):Next:execute N"
data = Replace(data, " ", ",")
fso.OpenTextFile(argv(0) & "_encrypted.vbe", 2, True).Write Encoder(EncHexXorData(data))
Function EncHexXorData(data)
EncHexXorData = "x=""" & EncHexXor(data) & """:For i=1 to Len(x) Step 2:s=s&Chr(CLng(""&H""&Mid(x,i,2)) Xor " & pass & "):Next:Execute Replace(s,"","","" "")"
End Function
Function Encoder(data) '加密3
Encoder = CreateObject("Scripting.Encoder").EncodeScriptFile(".vbs", data, 0, "VBScript")
End Function
Function EncHexXor(x) '加密2
For i = 1 To Len(x)
EncHexXor = EncHexXor & Hex(Asc(Mid(x, i, 1)) Xor pass)
Next
End Function
Function ASCdata(Data) '加密1
num = Len(data)
newdata = ""
For j = 1 To num
If j = num Then
newdata = newdata&Asc(Mid(data, j, 1))
Else
newdata = newdata&Asc(Mid(data, j, 1)) & " "
End If
Next
ASCdata = newdata
End Function
MsgBox "加密成功！已经加密的脚本文件保存到以下目录：" & vbCrLf  & vbCrLf & argv(0) & "_encrypted.vbe",vbOKOnly+vbInformation,"系统提示"
end if