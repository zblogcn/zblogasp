<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    uEditorRemoteImage - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////



'注册插件
Call RegisterPlugin("uEditorRemoteImage","ActivePlugin_uEditorRemoteImage")
'挂口部分
Function ActivePlugin_uEditorRemoteImage()
	Call Add_Filter_Plugin("Filter_Plugin_UEditor_Config","uEditorRemoteImage")	
	Call Add_Action_Plugin("Action_Plugin_uEditor_getRemoteImage_Begin","uEditorGetRemoteImage")	
End Function

Function uEditorRemoteImage(s)
	s=Replace(s,",catchRemoteImageEnable: false",",catchRemoteImageEnable: true,"&_
	"catcherUrl: URL+""asp/getRemoteImage.asp"",catcherPath:"""&Split(Split(s,"imagePath:""")(1),"""")(0)&"""")
End Function
Function uEditorGetRemoteImage()
	dim uploadPath,PostTime
	Randomize
	PostTime=GetTime(Now())
	Dim strUPLOADDIR
	strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"\"&Year(GetTime(Now()))&"\"&Month(GetTime(Now()))
	CreatDirectoryByCustomDirectory(strUPLOADDIR)
	Dim l
	Dim strURL,i,j,k,t,strResponse(2),aryURL
	aryURL=Split(Trim(Request.Form("upfile")),uEditor_Split)  '得到地址数组
	
	Set objStream = Server.Createobject("adodb.stream") 
	Set objXmlhttp=Server.CreateObject("msxml2.serverxmlhttp")
	
	For l=0 To Ubound(aryURL)
		strURL=aryURL(l)   
		If Left(strURL,Len(BlogHost))<>BlogHost Then 
			If CheckRegExp(strURL,"^http.+?\.(jpe?g|gif|bmp|png)$") Then   '判断URL是否符合图片格式
				t=RandomFileName(Split(strURL,".")(Ubound(Split(strURL,"."))))  '得到重命名文件夹
				Dim objXmlHttp,objStream
				objXmlHttp.Open "GET",strURL,False
				objXmlHttp.Send
				If objXmlHttp.Status=200 Then  'HTTP状态码需要符合“200 OK”
					k=objXmlHttp.getResponseHeader("Content-Type")    
					If Instr(k,"image") Then '判断Content-type是否为图片
							CreatDirectoryByCustomDirectory BlogPath &  strUPLOADDIR &"\"
							objStream.Type =1 
							objStream.Open 
							objStream.Write objXmlhttp.ResponseBody 
							objStream.Savetofile BlogPath &  strUPLOADDIR &"\" & "\"&t,2 
							strResponse(0)=t&uEditor_Split&strResponse(0)  '记录保存位置
							strResponse(1)=strURL&uEditor_Split&strResponse(1)  '记录原地址					
							Dim uf
							Set uf=New TUpLoadFile
							uf.AuthorID=BlogUser.ID
							uf.AutoName=False
							uf.IsManual=True
							uf.FileSize=objStream.Size
							uf.FileName=t
							uf.UpLoad
							objStream.Close() 
							'Exit For
					End If
				End If
			End If
		End If
	Next
	strResponse(2)="OK"
	strResponse(0)=Left(Replace(Replace(strResponse(0),"\","/"),"'","\'"),Len(strResponse(0))-Len(uEditor_Split))
	strResponse(1)=Left(Replace(Replace(strResponse(1),"\","/"),"'","\'"),Len(strResponse(1))-Len(uEditor_Split))
	Response.Write "{'url':'"&strResponse(0)&"','tip':'"&strResponse(2)&"','srcUrl':'"&strResponse(1)& "'}"
	Set objStream=Nothing
	Set objXmlhttp=Nothing
	Response.End
End Function


%>