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

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

BlogTitle="WXR数据导入程序"

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

<form border="1" name="edit" id="edit" method="post" enctype="multipart/form-data" action="tool.asp?act=FileUpload"><p>上传<b>符合Wordpress eXtended RSS标准</b>的XML文件: </p><p><input type="file" id="edtFileLoad" name="edtFileLoad" size="20">  <input type="submit" class="button" value="提交" name="B1" onclick='' /> <input class="button" type="reset" value="重置" name="B2" /></p>
<br/>
说明:<b><br/>1:如果XML文件较大（>=2MB），请将文件重命名为wpimport.xml并使用FTP方式上传到插件文件夹下，然后重新打开本插件。使用HTTP方式上传大文件可能会出错！</b>
<br/>
2.由于WP和Z-BLOG的过滤参数并不相同，一些评论的用户名或E-mail可能会被Z-BLOG过滤
，本插件会自动把被过滤的用户名转换成“匿名”，E-Mail转换成“null@null.com”。如果对Z-BLOG的过滤参数不满意，可以自行修改FUNCTION\c_function.asp 66行CheckRegExp函数后的几个正则表达式。
<br/>
3:WP中所有的Pages将视为文章导入。
<br/>
4:WP中所有的revision（包括autosave）将被丢弃。
<br/>
5:导入过程非常占用主机CPU，如果服务器受限或者因此挂掉，可以先在本地IIS导入后再上传数据库。
<%

Dim strAct
strAct=Request.QueryString("act")

Dim fso,ftpupload
Set fso = CreateObject("Scripting.FileSystemObject")
	ftpupload=fso.FileExists(BlogPath & "/zb_users/plugin/wpimport/wpimport.xml")
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
	objUpLoadFile.FileName="wpimport.xml"
	objUpLoadFile.FullPath=BlogPath & "/zb_users/plugin/wpimport/wpimport.xml"

	
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
		objXmlFile.load(BlogPath & "/zb_users/plugin/wpimport/wpimport.xml")
		
		'加文章
		Dim objArticle
		Dim newXMLNode
		
		Response.Write "<br/> ------ <br/><p style='border-bottom:1px solid gray;'><b>导入记录:</b>Title|Alias|CategoryID|Level|PostTime|Tags</p>"
		If ftpupload Then Response.Write  "<p>读取插件文件夹下的wpimport.xml，导入完成后将自动删除。</p>"
		
		
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode = 0 Then

				'读取所有的Category
				Dim objCategory
				Response.Write "<p><b>导入所有的Category</b></p>"
				Set objNodeList = objXmlFile.documentElement.selectSingleNode("channel").selectNodes("wp:category")
				j=objNodeList.length-1
				For i=0 To j
					Set objCategory=New TCategory
					objCategory.Name=objNodeList(i).SelectSingleNode("wp:cat_name").text
					objCategory.Alias=objNodeList(i).SelectSingleNode("wp:category_nicename").text
					objCategory.Post
				Next
				Response.Write "<p>已导入"&i&"个Category</p>"

				IsRunGetCategory=False
				Call GetCategory()


				'读取所有的Tags
				Dim objTag
				Response.Write "<p><b>导入所有的Tags</b></p>"
				Set objNodeList = objXmlFile.documentElement.selectSingleNode("channel").selectNodes("wp:tag")
				j=objNodeList.length-1
				For i=0 To j
					Set objTag=New TTag
					objTag.Name=objNodeList(i).SelectSingleNode("wp:tag_name").text
					objTag.Alias=objNodeList(i).SelectSingleNode("wp:tag_slug").text
					objTag.Post
				Next
				Response.Write "<p>已导入"&i&"个Tag</p>"

				IsRunGetTags=False
				Call GetTags()

			
				Set objNodeList = objXmlFile.documentElement.selectSingleNode("channel").selectNodes("item")
				j=objNodeList.length-1
				For i=0 To j
					Response.Write "<p>……读取文章 " & i & "……"

					Set objArticle=New TArticle
					objArticle.ID=0
					If objNodeList(i).SelectSingleNode("wp:post_type").text="post" Then
						objArticle.CateID=GetCategoryIDbyName(objNodeList(i).SelectSingleNode("category[@domain='category']").text)

						Dim t,strTags
						Set t = objNodeList(i).SelectNodes("category[@domain='post_tag']")
						If t.length>0 Then
							strTags=""
							strTags=t(0).text
							For b = 1 to t.length - 1
							strTags=strTags & "," & t(b).text
							Next
							objArticle.Tag=ParseTag(strTags)
						End If

					End If
					If objNodeList(i).SelectSingleNode("wp:post_type").text="page" Then
						objArticle.FType=1
					End If

					objArticle.AuthorID=BlogUser.ID
					
					If objNodeList(i).selectNodes("wp:status").length>0 Then
						Select Case objNodeList(i).SelectSingleNode("wp:status").text
						Case "private":
							objArticle.Level=2
						Case "draft"
							objArticle.Level=1
						Case Else
							If objNodeList(i).selectNodes("wp:comment_status").length=0 then
									objArticle.Level=4
							Else
								If objNodeList(i).SelectSingleNode("wp:comment_status").text <> "open" Then
									objArticle.Level=3
								Else
									objArticle.Level=4
								End if
							End if
						End Select
					End If
					
					If objNodeList(i).selectNodes("pubDate").length>0 Then
						objArticle.PostTime=rfc_to_iso(objNodeList(i).SelectSingleNode("pubDate").text)
						If IsDate(objArticle.PostTime) Then
						objArticle.PostTime=CDate(objArticle.PostTime)
						Else
							objArticle.PostTime=Now()
						End If
					Else
						objArticle.PostTime=Now()
					End If
					If objNodeList(i).selectNodes("title").length>0 Then
						objArticle.Title=objNodeList(i).SelectSingleNode("title").text
					Else
						objArticle.Title="没有标题"
					End If
					
					If objNodeList(i).selectNodes("wp:post_name").length>0 Then
						objArticle.Alias=decodeurl(objNodeList(i).SelectSingleNode("wp:post_name").text)
					Else
						objArticle.Alias=""
					End if
					
					If objNodeList(i).selectNodes("content:encoded").length>0 Then
						objArticle.Content=closeHTML(Replace(objNodeList(i).SelectSingleNode("content:encoded").text,vblf,"<br/>"))
						objArticle.Intro=closeHTML(GetIntro(objArticle.Content))
					End If

					Response.write "<br/>尝试导入日志:<b>" & objArticle.Title & "|" & objArticle.Alias & "|" & objArticle.CateID & "|" & objArticle.Level & "|" & objArticle.PostTime & "|" & strTags & "</b><br/>"
					
					If objArticle.Post Then
						Response.Write "导入日志成功！"
						'导入评论和TrackBack

							Set c = objNodeList(i).SelectNodes("wp:comment")
							a=c.length-1
							Response.Write "评论数：" & c.length & "</p>"
							For b=0 To a
								Dim objComment
								Set objComment=New TComment
								objComment.ID=0
								objComment.AuthorID=0
								objComment.log_ID=objArticle.ID
								if c(b).SelectNodes("wp:comment_author_url").length>0 then objComment.HomePage=c(b).SelectSingleNode("wp:comment_author_url").text
								If (Not CheckRegExp(objComment.HomePage,"[homepage]")) And (Not CheckRegExp("http://" & objComment.HomePage,"[homepage]")) Then objComment.HomePage=""
								objComment.Content=c(b).SelectSingleNode("wp:comment_content").text
								objComment.PostTime=c(b).SelectSingleNode("wp:comment_date").text
								objComment.Author=c(b).SelectSingleNode("wp:comment_author").text
								If Not CheckRegExp(objComment.Author,"[username]") Then objComment.Author="匿名"
								if c(b).SelectNodes("wp:comment_author_IP").length>0 then objComment.IP=c(b).SelectSingleNode("wp:comment_author_IP").text
								if c(b).SelectNodes("wp:comment_author_email").length>0 then objComment.Email=c(b).SelectSingleNode("wp:comment_author_email").text
								If Not CheckRegExp(objComment.Email,"[email]") Then objComment.Email="null@null.com"
								
								Response.Write "<p>尝试导入评论:<b>Author:"& objComment.Author & "|Email:" & objComment.Email & "</b><br/>"
									If objComment.Post Then
										Response.Write "导入评论成功！</p>"
									Else
										Response.Write "导入评论失败！</p>"
									End if
								Set objComment=Nothing																
							Next




					End If
					
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
	fso.DeleteFile(BlogPath & "/zb_users/plugin/wpimport/wpimport.xml")

	
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
%>
<script language="javascript" runat="server">
function rfc_to_iso(DataRFC){
var dateTimeObject = new Date(DataRFC);
var timezoneOffset = new Date().getTimezoneOffset()/60*-1;
var shuchu = dateTimeObject.toLocaleString();
return(shuchu);
}</script>