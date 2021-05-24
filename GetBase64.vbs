on error resume next
Set argv = WScript.Arguments
If argv.Count = 0 Then
    data = Inputbox("欢迎使用 Base64 编码小工具！"& vbCrLf& " "& vbCrLf& "请输入源码：" ,"Base64 编码小工具")
    If data =vbEmpty Then
        WScript.Quit
    End If
    data = ConvertStringToUTF8(data)
Else
    data=""
    With CreateObject("ADODB.Stream")
        .Mode = 3: .Type = 1: .Open: .LoadFromFile argv(0)
        .Read(3)
        For j = 3 To .Size - 1
            tmp = AscB(.Read(1))
            If tmp < 16 Then
                tmp = "0" & Hex(tmp)
            Else tmp = Hex(tmp)
            End If
            data =data & tmp
        Next
       .Close
    End With
End If
 
newdata = ""
For j = 1 To Len(data) Step 2
    tmp = CLng("&H"&Mid(data,j,2))
    m = ""
    Do
        m = m & tmp mod 2
        tmp = tmp\2
    Loop Until tmp=0
    m = StrReverse(m)
    Do While Len(m)<8
        m = "0" & m
    Loop
    newdata = newdata & m
Next
 
res = ""
For j = 1 To Len(newdata) Step 6
    tmp = Mid(newdata, j, 6)
    num = 0
    For i = 1 To 6
        m = Mid(tmp, i, 1)
        if m = "" Then m = 0
        num = num * 2 + m
    Next
    If num <= 25 Then
        res = res & chr(65+num)
    ElseIf num <= 51 Then
        res = res & chr(71+num)
    ElseIf num <=61 Then
        res = res & chr(num-4)
    ElseIf num = 62 Then
        res = res & "+"
    ElseIf num = 63 Then
        res = res & "/"
    End If
Next
 
tmp = Len(res) mod 4
If tmp <>0 Then
    For j = 1 To 4-tmp
        res = res & "="
    Next
End If

Set fso=CreateObject("Scripting.FileSystemObject")

Set ts = fso.CreateTextFile("base64.txt", true)
ts.write res
WScript.Echo "Base64 编码是："&vbCrlf&" "& vbCrLf & res &vbCrlf&" "& vbCrLf& "Base64 编码将自动导出到同目录下的“base64.txt”文件中。"

'InputBox "结果如下","",res,"Base64 编码小工具"

Function ConvertStringToUTF8(TextString)
    Set objXmlDom = CreateObject("Microsoft.XMLDOM")
    Set objNode = objXmlDom.CreateElement("Binary")
    objNode.DataType = "bin.hex"
    With CreateObject("ADODB.Stream")
        .Type = 2:.Charset = "UTF-8":.Open
        .WriteText TextString
        .Position = 0:.Type = 1
        objNode.nodeTypedValue = .Read
        .Close
    End With
    ConvertStringToUTF8 = Mid(objNode.Text,7)
End Function
