<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog(http://www.rainbowsoft.org)
'// 插件制作:    Zx.MYS
'// 备    注:    Wordpress数据导入程序  （改编自zx.asd的rss2.0导入插件）
'// 最后修改：   2010-7-11
'// 最后版本:    1.41
'///////////////////////////////////////////////////////////////////////////////
%>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=False %>
<% Server.ScriptTimeOut=1000000 %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%
Const VideoTemplate="<p><embed src=""<#swfpath#>"" quality=""high"" width=""480"" height=""400"" align=""middle"" allowScriptAccess=""always"" allowFullScreen=""true"" mode=""transparent"" type=""application/x-shockwave-flash""></embed></p><p><a href=""<#link#>"" target=""_blank"">在新窗口中打开:<#swfname#></a></p>"

Const ImageTemplate="<p><img src=""<#imgurl#>"" alt=""<#imgid#>"" title=""<#imgid#>""/></p>"

Const MusicTemplate="<div class=""audio-cover""><img src=""<#imgsrc#>"" alt=""<#imgalt#>""/></div><div class=""audio-player""><object width=""257"" height=""33"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""><param value=""http://www.xiami.com/widget/0_<#songid#>/singlePlayer.swf"" name=""movie""></param><param value=""transparent"" name=""wmode""></param><param value=""high"" name=""quality""></param><embed width=""257"" height=""33"" type=""application/x-shockwave-flash"" pluginspage=""http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"" quality=""high"" wmode=""transparent"" menu=""false"" src=""http://www.xiami.com/widget/0_<#songid#>/singlePlayer.swf"" /></object></div>"
Call System_Initialize()
Dim strTemp

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

BlogTitle="点点数据导入程序"

Dim intPageCount
intPageCount=CInt(Request.QueryString("pagecount"))
If intPageCount=0 Then intPageCount=100

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

			<div id="divMain">
<div class="divHeader"><%=BlogTitle%></div>
<div id="divMain2">
<%
'*********************************************************
' 目的：    
'*********************************************************
Function GetCategoryIDbyName(Name)

	GetCategory()

	Dim Category
	For Each Category in Categorys
		If IsObject(Category) Then
			If Name=Category.Name Then
				GetCategoryIDbyName=Category.ID
				Exit Function
			End If
		End If
	Next

End Function
'*********************************************************


'*********************************************************
' 目的：    
'*********************************************************
Function GetTagIDbyName(Name)

	GetTags()

	Dim Category
	For Each Category in Categorys
		If IsObject(Category) Then
			If Name=Category.Name Then
				GetCategoryIDbyName=Category.ID
				Exit Function
			End If
		End If
	Next

End Function
'*********************************************************

Dim introPos
Function GetIntro(ByVal Str)
introPos=instr(str,"<!--more-->")-1
If introPos>0 then
str=left(str,introPos)
Else
'str=left(str,200)
str=str
End if
GetIntro=str
End Function 
%>
<script runat="server" language="javascript">
function decodeurl(Source){
	try{
		return decodeURIComponent(Source);
	}catch(e){
		return "";
	}
}
//javascript就是好~~
</script>

<form border="1" name="edit" id="edit" method="post" enctype="multipart/form-data" action="tool.asp?act=FileUpload">
  <p>上传<b>点点导出的</b>的XML文件: </p><p><input type="file" id="edtFileLoad" name="edtFileLoad" size="20">  <input type="submit" class="button" value="提交" name="B1" onclick='' /> <input class="button" type="reset" value="重置" name="B2" /></p>
<br/>
说明:<b><br/>
1:如果XML文件较大（>=2MB），请将文件重命名为diandian.xml用FTP方式上传到插件文件夹下，然后重新打开本插件。使用HTTP方式上传大文件可能会出错！</b>
<br/>
2:无法导入置顶文章和评论。导入过程中原有地址可能会丢失。
<br/>
3:导入过程非常占用主机CPU，如果服务器受限或者因此挂掉，可以先在本地IIS导入后再上传数据库。
<%

Dim strAct
strAct=Request.QueryString("act")

Dim fso,ftpupload
Set fso = CreateObject("Scripting.FileSystemObject")
	ftpupload=fso.FileExists(Server.MapPath("diandian.xml"))
Set fso = Nothing

If strAct="FileUpload" or ftpupload Then

	Dim objXmlFile
	Dim objNodeList
	Dim i,j,a,b,c,failed 

	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile


	objUpLoadFile.AutoName=False
	objUpLoadFile.IsManual=True
	objUpLoadFile.FileSize=0
	objUpLoadFile.FileName="diandian.xml"
	objUpLoadFile.FullPath=BlogPath & "/zb_users/plugin/diandian/diandian.xml"

	
	If ftpupload Then
		c=True
	Else
		Call objUpLoadFile.UpLoad_Form()
		Call objUpLoadFile.SaveFile()
		c=True
	End If

	If c Then
		Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async=False
		objXmlFile.load(BlogPath & "/zb_users/plugin/diandian/diandian.xml")

		'加文章
		Dim objArticle
		Dim newXMLNode
		
		Response.Write "<br/> ------ <br/><p style='border-bottom:1px solid gray;'><b>导入记录:</b>Title|Alias|CategoryID|Level|PostTime|Tags</p>"
		If ftpupload Then Response.Write  "<p>读取插件文件夹下的diandian.xml，导入完成后将自动删除。</p>"
		
		
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode = 0 Then

				Set objNodeList = objXmlFile.documentElement.selectSingleNode("Posts").selectNodes("Post")
				j=objNodeList.length-1
				For i=0 To j
					

					Set objArticle=New TArticle
					objArticle.ID=0
					objArticle.CateID=0
					Dim donotPost
					Select Case objNodeList(i).SelectSingleNode("PostType").text
						Case "link"
							Response.Write "<p>……读取链接 " & i & "……"
							objArticle.Title=IIf(objNodeList(i).SelectSingleNode("Title").text="","分享链接",objNodeList(i).SelectSingleNode("Title").text)
							objArticle.Content="<a href='" & objNodeList(i).SelectSingleNode("Link").text & "' target='_blank' rel='nofollow'>" & IIf(objNodeList(i).SelectSingleNode("Title").text="","分享链接",objNodeList(i).SelectSingleNode("Title").text)
						Case "video"
							Response.Write "<p>……读取视频 " & i & "……"
							objArticle.Title=IIf(objNodeList(i).SelectSingleNode("VideoName").text="","分享视频",objNodeList(i).SelectSingleNode("VideoName").text)
							strTemp=VideoTemplate
							strTemp=Replace(strTemp,"<#swfpath#>",objNodeList(i).SelectSingleNode("VideoFlashUrl").text)
							strTemp=Replace(strTemp,"<#swfname#>",objNodeList(i).SelectSingleNode("VideoName").text)
							strTemp=Replace(strTemp,"<#link#>",objNodeList(i).SelectSingleNode("VideoVideoUrl").text)
							objArticle.Content=strTemp
						Case "audio"
							'Response.Write "<p>抱歉，ID为"& objNodeList(i).SelectSingleNode("ID").text & "的日志为音频（《"&objNodeList(i).SelectSingleNode("SongName").text&"》），无法导入</p>"
							objArticle.Title=IIf(objNodeList(i).SelectSingleNode("SongName").text="","分享音乐",objNodeList(i).SelectSingleNode("SongName").text)
							strTemp=MusicTemplate
							strTemp=Replace(strTemp,"<#imgsrc#>",objNodeList(i).SelectSingleNode("Cover").text)
							strTemp=Replace(strTemp,"<#imgalt#>",objNodeList(i).SelectSingleNode("SongName").text)
							strTemp=Replace(strTemp,"<#songid#>",objNodeList(i).SelectSingleNode("SongId").text)
							objArticle.Content=strTemp
							
							'donotPost=True
						Case "photo"
							Response.Write "<p>……读取图片 " & i & "……"
							Dim sImage,sj
							Set sImage=objNodeList(i).selectNodes("PhotoItem")
							For sj=0 To sImage.length-1
								strTemp=strTemp&ImageTemplate
								strTemp=Replace(strTemp,"<#imgurl#>",GetImage(sImage(sj).text,objXmlFile))
								strTemp=Replace(strTemp,"<#imgid#>",sImage(sj).text)
							Next
							objArticle.Content=strTemp
							If objNodeList(i).selectNodes("Title").length>0 Then objArticle.Title=objNodeList(i).SelectSingleNode("Title").text
						Case "text"
							Response.Write "<p>……读取文章 " & i & "……"
							
							If objNodeList(i).selectNodes("Title").length>0 Then objArticle.Title=objNodeList(i).SelectSingleNode("Title").text

							objArticle.Content=objNodeList(i).SelectSingleNode("Text").text
					End Select
					
					Dim objRegExp
					Set objRegExp=New RegExp
					objRegExp.Global=True
					objRegExp.Pattern="img(.+?)(id="".+?)"""
					objRegExp.IgnoreCase=True
					Dim SubMatches,SubMatch
					Set SubMatches=objRegExp.Execute(objArticle.Content)
					For Each SubMatch In SubMatches
						objArticle.Content=Replace(objArticle.Content,SubMatch.SubMatches(1),"src="""&GetImage(Split(SubMatch.SubMatches(1),"id=""")(1),objXmlFile))
					Next
					
					objArticle.Intro=closeHTML(GetIntro(objArticle.Content))
					If objNodeList(i).SelectSingleNode("Privacy").text<>0 Then objArticle.Level=2
					Dim t,strTags
					If objNodeList(i).selectNodes("Tags").length>0 Then
						Set t = objNodeList(i).SelectSingleNode("Tags").SelectNodes("Tag")
						If t.length>0 Then
							strTags=""
							strTags=t(0).text
							For b = 1 to t.length - 1
								strTags=strTags & "," & t(b).text
							Next
							objArticle.Tag=ParseTag(strTags)
						End If
					End If



					objArticle.AuthorID=BlogUser.ID
			
					
					If objNodeList(i).selectNodes("CreateTime").length>0 Then
						objArticle.PostTime=rfc_to_iso(CDbl(objNodeList(i).SelectSingleNode("CreateTime").text))
						Response.Write rfc_to_iso(objNodeList(i).SelectSingleNode("CreateTime").text)
						If IsDate(objArticle.PostTime) Then
							objArticle.PostTime=CDate(objArticle.PostTime)
						Else
							objArticle.PostTime=Now()
						End If
						
					Else
						objArticle.PostTime=Now()
					End If
					
					
					If objNodeList(i).selectNodes("Uri").length>0 Then
						objArticle.Alias=decodeurl(objNodeList(i).SelectSingleNode("Uri").text)
					Else
						objArticle.Alias=""
					End if
					
					If objArticle.Title="" Then objArticle.Title="未命名文章"

					Response.write "<br/>尝试导入日志:<b>" & objArticle.Title & "|" & objArticle.Alias & "|" & objArticle.CateID & "|" & objArticle.Level & "|" & objArticle.PostTime & "|" & strTags & "</b><br/>"
					
					If objArticle.Post Then
						Response.Write "导入日志成功！"
					End If
					
					'Response.Flush
					Set objArticle=Nothing

				Next

				Call SetBlogHint(Null,True,True)
				
				Response.Write "<br/><p><b>导入完成!</b></p>"
				Response.Write "<script type=""text/javascript"">alert(""导入完成,请进行[文件重建]."");</script>"

				Set objNodeList = Nothing
				
			Else
			
				Response.Write "<p>XML文件格式错误,请上传完好的XML格式文件（建议使用FTP上传），以下是错误信息：<br/>"
				Response.Write objXmlFile.parseError.errorCode & ":" & objXmlFile.parseError.reason & "<br/>位置:" & objXmlFile.parseError.line & "," & objXmlFile.parseError.linePos & "<br/>内容:" & TransferHTML(objXmlFile.parseError.srcText,"[html-format][enter][""]") & "</p>"

			End If
		End If
		
	End If
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(BlogPath & "/zb_users/plugin/diandian/diandian.xml")

	
End If

%>
</form>
</div>
    <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

'If Err.Number<>0 then
'  Call ShowError(0)
'End If



Function GetImage(ID,xmlObj)
	Dim oImageList
	Set oImageList=xmlObj.documentElement.selectSingleNode("Images").selectNodes("Image")
	Dim i
	For i=0 to oImageList.length-1
		If oImageList(i).selectSingleNode("Id").text=ID Then 
			GetImage=oImageList(i).selectSingleNode("Url").text
			Exit Function
		End If
	Next
	GetImage=""
End Function
%>
<script language="javascript" runat="server">
function rfc_to_iso(DataRFC){
var dateTimeObject = new Date(DataRFC);
var timezoneOffset = new Date().getTimezoneOffset()/60*-1;
var str = dateTimeObject.toLocaleString();
//var d=new Date(str);
//var shuchu=d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDay()+" "+d.getHours()+":"+d.getMinutes()+":"+d.getSeconds();
//Response.Write(str);
//Response.End()
var shuchu=str;
return(shuchu);
}</script>