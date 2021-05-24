ext_id=inputbox("Microsoft Edge 插件下载器" &vbCrlf&"" &vbCrlf&"该小工具可以从直接从微软官方插件商店下载插件（.crx），将离线插件拖拽到扩展管理页面（edge://extensions/）即可直接安装！"& vbCrLf&""& vbCrLf& "请输入插件ID号：" ,"Microsoft Edge 插件下载器")
If ext_id=vbEmpty Then
wscript.quit
Else currentpath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
Set Post = CreateObject("Msxml2.ServerXMLHTTP")
Set Shell = CreateObject("Wscript.Shell")
extid=ext_id
msgbox "插件正在后台下载中，请稍候！" &vbCrlf&"" &vbCrlf& "下载后的文件将保存到当前目录下。" ,vbOKOnly+vbInformation,"Microsoft Edge 插件下载器" 
Post.Open "GET","https://jimmyjlnu.herokuapp.com/proxy/https://edge.microsoft.com/extensionwebstorebase/v1/crx?response=redirect&prod=chromiumcrx&prodchannel=&x=id%3D"&extid&"%26installsource%3Dondemand%26uc",0
Post.Send
Set aGet = CreateObject("ADODB.Stream")
aGet.Mode = 3
aGet.Type = 1
aGet.Open() 
aGet.Write(Post.responseBody)
aGet.SaveToFile currentpath &"\"&extid&".crx",2
msgbox "插件下载成功！" &vbCrlf&"" &vbCrlf& "下载后的文件已经保存到当前目录下。" ,vbOKOnly+vbInformation,"Microsoft Edge 插件下载器" 
end if
