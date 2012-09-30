<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Devo Or Newer
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    批量管理文章插件 - 跳转页
'// 最后修改：   2009-6-16
'// 最后版本:    1.4.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="config.asp" -->
<%


'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>4 Then Call ShowError(6) 

If CheckPluginState("BatchArticles")=False Then Call ShowError(48)

BlogTitle="Batch Articles"
Call GetUser
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

	<style>#TagCloud span{display:inline;margin:0 4px 0 0;padding:0 0 0 0;line-height:160%;white-space: nowrap;}</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<script type="text/javascript">
	$(function(){
		BT_setOptions({openWait:100, closeWait:0, cacheEnabled:true});
	})
	ActiveLeftMenu("aArticleMng");</script>
</script>

<div id="divMain">
<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">批量管理文章插件 - 管理页面</div>
<div class="SubMenu">
	<a href="articleList.asp"><span class="m-left m-now">[批量文章管理]</span></a>
	<%=Response_Plugin_ArticleMng_SubMenu%>
	<a href="help.asp"><span class="m-right">[帮助和设置]</span></a>
</div>
<div id="divMain2">

<%

	Dim intPage,intCate,intLevel,intIstop,intUser,intTag,intTitle,tagFilter

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	if Request.QueryString("page")<>"" then intPage=Request.QueryString("page")
	if Request.QueryString("cate")<>"" then intCate=Request.QueryString("cate")
	if Request.QueryString("level")<>"" then intLevel=Request.QueryString("level")
	if Request.QueryString("istop")<>"" then intIstop=Request.QueryString("istop")
	if Request.QueryString("user")<>"" then intUser=Request.QueryString("user")
	if Request.QueryString("tag")<>"" then intTag=Request.QueryString("tag")
	if Request.QueryString("title")<>"" then intTitle=Request.QueryString("title")
	if Request.QueryString("tagfilter")<>"" then tagFilter=Request.QueryString("tagfilter")

	if Request.Form("cate")<>"" then intCate=Request.Form("cate")
	if Request.Form("level")<>"" then intLevel=Request.Form("level")
	if Request.Form("istop")<>"" then intIstop=Request.Form("istop")
	if Request.Form("user")<>"" then intUser=Request.Form("user")
	if Request.Form("tag")<>"" then intTag=Request.Form("tag")
	if Request.Form("title")<>"" then intTitle=Escape(Request.Form("title"))
	if Request.Form("tagfilter")<>"" then tagFilter=Escape(Request.Form("tagfilter"))

	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(intCate,"int",-1)
	Call CheckParameter(intLevel,"int",-1)
	Call CheckParameter(intIstop,"int",-1)
	Call CheckParameter(intUser,"int",-1)
	Call CheckParameter(intTag,"int",-1)
	Call CheckParameter(intTitle,"sql",-1)
	Call CheckParameter(tagFilter,"sql",-1)
	intTitle=vbsunescape(intTitle)
	intTitle=FilterSQL(intTitle)
	tagFilter=vbsunescape(tagFilter)
	tagFilter=FilterSQL(tagFilter)

	'调出 Tags 列表
	If UseTagMng Then
		Call GetTags
		Dim strTagSel,strTagCloud,strTagList,strTagEdit
		Dim strTagID,strTagName,strTagFilterSQL

		If tagFilter<>"-1" Then
			strTagFilterSQL = "WHERE "&ExportSearch("tag_Name",tagFilter)
		End If

		Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name] FROM [blog_Tag] "& strTagFilterSQL &" ORDER BY [tag_Name] ASC")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Do While Not objRS.eof
				If CStr(intTag) = CStr(objRS("tag_ID")) Then strTagSel = "selected"
				strTagID = objRS("tag_ID")
				strTagName = TransferHTML(Tags(objRS("tag_ID")).Name,"[html-format]")
				strTagCloud = strTagCloud & "<span><a href=""articleList.asp?tag="& strTagID &""">"& strTagName &"</a></span> "
				strTagList = strTagList & "<option value="""& strTagID &""" "& strTagSel &">"& strTagName &"</option>"
				strTagEdit = strTagEdit & "<option value="""& strTagID &""" >"& strTagName &"</option>"
				strTagSel = ""
				objRS.MoveNext
			Loop
		End If
		objRS.Close
		Set objRS=Nothing
	End If

	Response.Write "<form id=""edit"" method=""post"" enctype=""application/x-www-form-urlencoded"" action=""articleList.asp""><p>"
	'Response.Write "<p>"&ZC_MSG158&": {"& intCate &"/"& intLevel &"/"& intIstop &"/"& intTitle &"/"& intUser &"}</p><p>"

	Response.Write ZC_MSG012&" <select class=""edit"" size=""1"" id=""cate"" name=""cate"" style=""width:120px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "
	Dim Category
	For Each Category in Categorys
		If IsObject(Category) Then
			Response.Write "<option value="""&Category.ID&""" "
			If Category.ID=intCate Then Response.Write "selected"
			Response.Write ">"&TransferHTML(Category.Name,"[html-format]")&"</option>"
		End If
	Next
	Response.Write "</select> "

	Response.Write ZC_MSG061&" <select class=""edit"" size=""1"" id=""level"" name=""level"" style=""width:90px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "

	For i=LBound(ZVA_Article_Level_Name)+1 to Ubound(ZVA_Article_Level_Name)
			Response.Write "<option value="""&i&""" "
			If i=intLevel Then Response.Write "selected"
			Response.Write ">"&ZVA_Article_Level_Name(i)&"</option>"
	Next
	Response.Write "</select> "

	Response.Write ZC_MSG051&" <select class=""edit"" size=""1"" id=""istop"" name=""istop"" style=""width:90px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "
			Response.Write "<option value=""1"""
			If intIstop = 1 Then Response.Write "selected"
			Response.Write ">是</option>"
			Response.Write "<option value=""0"""
			If intIstop = 0 Then Response.Write "selected"
			Response.Write ">否</option>"
	Response.Write "</select> "

	Dim User
	If CheckRights("Root")=True Then
	Response.Write ZC_MSG003&" <select class=""edit"" size=""1"" id=""user"" name=""user"" style=""width:90px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "
	For Each User in Users
		If IsObject(User) And User.Level<=2 Then
				Response.Write "<option value="""&User.ID&""" "
				If intUser = User.ID Then Response.Write "selected"
				Response.Write ">"&TransferHTML(User.Name,"[html-format]")&"</option>"
		End If
	Next
	Response.Write "</select></p><p> "
	End If

	If UseTagMng Then
		Dim TagFilterValue
		If tagFilter="-1" Then : TagFilterValue="" : Else : TagFilterValue=tagFilter : End If
		Response.Write ""& ZC_MSG138 &" <input id=""tagfilter"" name=""tagfilter"" style=""width:60px;"" type=""text"" value="""& TagFilterValue &""" /> <select class=""edit"" size=""1"" id=""tag"" name=""tag"" style=""width:80px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "
			Response.Write strTagList
		Response.Write "</select> "
	End If

	Dim TitleValue
	If intTitle="-1" Then : TitleValue="" : Else : TitleValue=intTitle : End If
	Response.Write " "&ZC_MSG224&" <input id=""title"" name=""title"" style=""width:100px;"" type=""text"" value="""& TitleValue &""" /> "
	Response.Write "<input type=""submit"" class=""button"" value="""&ZC_MSG087&""" title=""执行筛选""> "

	If UseTagMng Then
		If UseTagCloud Then
			Dim TagCloudDisplay
			If tagFilter="-1" Then : TagCloudDisplay="none" : Else : TagCloudDisplay="block" : End If
			Response.Write " <span onclick=""$('#TagCloud').toggle('normal');"" style=""cursor:pointer;color:navy;"">[点此展开TagCloud]</span></p><p id=""TagCloud"" style=""display:"& TagCloudDisplay &";"">"
			'Response.Write " <span onclick=""if(document.getElementById('TagCloud').style.display == 'none'){document.getElementById('TagCloud').style.display = 'block'}else{document.getElementById('TagCloud').style.display = 'none'};"" style=""cursor:pointer;color:navy;"">[点此展开TagCloud]</span></p><p id=""TagCloud"" style=""display:"& TagCloudDisplay &";"">"
			Response.Write strTagCloud
		End If
	End If

	Response.Write "</p></form>"



	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL="WHERE ([log_Level]>0) AND ([log_Type]=0) "

	If CheckRights("Root")=False Then
		strSQL= strSQL & "AND [log_AuthorID] = " & BlogUser.ID
	Else

		If intUser<>-1 Then
			strSQL= strSQL & " AND [log_AuthorID] = " & intUser
		End If

	End If

	If intCate<>-1 Then
		strSQL= strSQL & " AND [log_CateID] = " & intCate
	End If

	If intLevel<>-1 Then
		strSQL= strSQL & " AND [log_Level] = " & intLevel
	End If

	If intTag<>-1 Then
		strSQL= strSQL & " AND " & ExportSearch("log_Tag","{"& intTag &"}")
	End If

	If intIstop<>-1 Then
		if intIstop = "1" then
		strSQL= strSQL & " AND [log_Istop]=1"
		end if
		if intIstop = "0" then
		strSQL= strSQL & " AND [log_Istop]=0"
		end if
	End If

	If intTitle<>"-1" Then

		Dim aryTitle,itemTitle

		intTitle=Replace(intTitle,"　"," ")
		aryTitle=Split(intTitle," ")

		For Each itemTitle In aryTitle

			itemTitle=Trim(itemTitle)

			If itemTitle<>"" Then

				strSQL = strSQL & "AND ("&ExportSearch("log_Title",itemTitle)&" OR "&ExportSearch("log_Intro",itemTitle)&" OR "&ExportSearch("log_Content",itemTitle)&" )"

			End If

		Next
		
	End If

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
	Response.Write "<tr><td>"& ZC_MSG076 &"</td><td>"& ZC_MSG012 &"</td><td>"& ZC_MSG003 &"</td><td>"& ZC_MSG061 &"</td><td>"& ZC_MSG051 &"</td><td>"& ZC_MSG075 &"</td><td>"& ZC_MSG060 &"</td><td>"& ZC_MSG047 &"</td><td align='center'><a href='' onclick='BatchSelectAll();return false'>"& ZC_MSG229 &"</a></td></tr>"
	objRS.Open("SELECT * FROM [blog_Article] "& strSQL &" ORDER BY [log_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"

			Response.Write "<td>" & objRS("log_ID") & "</td>"

			Dim Cate
			For Each Cate in Categorys
				If IsObject(Cate) Then
					If Cate.ID=objRS("log_CateID") Then
						Response.Write "<td><a href=""articleList.asp?cate="& Cate.ID &""" title="""& Cate.Name &""">" & Left(Cate.Name,8) & "</a></td>"
					End If
				End If
			Next

			
			'Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("log_AuthorID") Then
						Response.Write "<td><a href=""articleList.asp?user="& User.ID &""" title="""& User.Name &""">" & Left(User.Name,10) & "</a></td>"
					End If
				End If
			Next

			Dim Istop
				if objRS("log_IsTop")=True then
					Istop = "<font color=""green""><b>&nbsp;√&nbsp;</b></font> "
				else
					Istop = ""
				end if

			Response.Write "<td><a href=""articleList.asp?level="& objRS("log_Level") &""" title="""& ZVA_Article_Level_Name(objRS("log_Level")) &""">" & ZVA_Article_Level_Name(objRS("log_Level")) & "</a></td>"
			Response.Write "<td>" & Istop & "</td>"
			Response.Write "<td>" & FormatDateTime(objRS("log_PostTime"),vbShortDate) & "</td>"

			Response.Write "<td>"
			If Len(objRS("log_Title"))>22 Then
				Response.Write "<a href=""../../../zb_system/view.asp?id=" & objRS("log_ID") & """ target=""_blank"" title=""" & objRS("log_Title") & """>" & Left(objRS("log_Title"),21) & ".." & "</a>"
			Else
				Response.Write "<a href=""../../../zb_system/view.asp?id=" & objRS("log_ID") & """ target=""_blank"" title=""" & objRS("log_Title") & """>" & objRS("log_Title") & "</a>"
			End If
			Response.Write " <a id=""mylink"&objRS("log_ID")&""" href=""$div"&objRS("log_ID")&"tip?width=125"" class=""betterTip"" title=""关键词(Tags)""><img src=""tag.png""  alt=""TagIcon"" align=""absbottom"" ></a><div id=""div"&objRS("log_ID")&"tip"" style=""display:none;"">"& Export_ArticleTag(objRS("log_Tag")) &"</div>"
			Response.Write "</td>"

			Response.Write "<td align=""center""><a href=""../../cmd.asp?act=ArticleEdt&type="& ZC_BLOG_WEBEDIT &"&id=" & objRS("log_ID") & """>[可视]</a></td>"
			Response.Write "<td align=""center"" ><input type=""checkbox"" name=""edtDel"" id=""edtDel"" value="""&objRS("log_ID")&"""/></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	objRS.Close
	Set objRS=Nothing

	Response.Write "</table>"

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"articlelist.asp?cate="&ReQuest("cate")&"&level="&ReQuest("level")&"&istop="&ReQuest("istop")&"&title="&Escape(ReQuest("title"))&"&user="&ReQuest("user")&"&tag="&ReQuest("tag")&"&tagfilter="&Escape(ReQuest("tagfilter"))&"&amp;page=")

	Response.Write "<br/><p>注意: 批量执行会消耗一定时间, 请耐心等候. 一次执行文章数量过多还可能导致系统资源过载,  请谨慎使用.</p>"

	'批量设置
	Response.Write "<form id=""frmBatch"" method=""post"" action=""articleMng.asp?act=BasicEdit&page="&Request.QueryString("page")&"&cate="&ReQuest("cate")&"&level="&ReQuest("level")&"&istop="&ReQuest("istop")&"&user="&ReQuest("user")&"&tag="&ReQuest("tag")&"&title="&Escape(ReQuest("title"))&"""><p><input type=""hidden"" id=""edtBatch"" name=""edtBatch"" value=""""/>"

	Response.Write "<select name=""MoveCata"" id=""MoveCata""><option value=""-1"">转移到分类...</option>"
	For Each Category in Categorys
		If IsObject(Category) Then
			Response.Write "<option value="""&Category.ID&""" "
			Response.Write ">"&TransferHTML(Category.Name,"[html-format]")&"</option>"
		End If
	Next
	Response.Write "</select>&nbsp;"

	Response.Write "<select name=""EdtLevel"" id=""EdtLevel""><option value=""-1"">类型设置为...</option>"
	For i=LBound(ZVA_Article_Level_Name)+1 to Ubound(ZVA_Article_Level_Name)
			Response.Write "<option value="""&i&""" "
			Response.Write ">"&ZVA_Article_Level_Name(i)&"</option>"
	Next
	Response.Write "</select>&nbsp;"

	Response.Write "<select name=""EdtIstop"" id=""EdtIstop""><option value=""-1"">置顶设置为...</option>"
			Response.Write "<option value=""1"">是 (置顶)</option>"
			Response.Write "<option value=""0"">否 (不置顶)</option>"
	Response.Write "</select>&nbsp;"

	If CheckRights("Root")=True Then
	Response.Write "<select name=""EdtUser"" id=""EdtUser""><option value=""-1"">用户更改为...</option> "
	For Each User in Users
		If IsObject(User) Then
				Response.Write "<option value="""&User.ID&""" "
				Response.Write ">"&TransferHTML(User.Name,"[html-format]")&"</option>"
		End If
	Next
	Response.Write "</select>&nbsp;"
	End If

	Response.Write "&nbsp;删除:<input type=""checkbox"" name=""BatchDel"" id=""BatchDel"" value=""True"" onclick=""BatchDelEnabled();return true""/></p><p>"

	If UseTagMng Then
		Response.Write "<select name=""AddTag"" id=""AddTag"" style=""width:120px;""><option value=""-1"">增加标签...</option>"
			Response.Write strTagEdit
		Response.Write "</select>&nbsp;"

		If UseTagHint Then strTagEdit = strTagList

		Response.Write "<select name=""RmvTag"" id=""RmvTag"" style=""width:120px;""><option value=""-1"">删除标签...</option>"
			Response.Write strTagEdit
		Response.Write "</select>&nbsp;"
	End If

	Response.Write "<input class=""button"" type=""submit"" onclick='BatchDeleteAll(""edtBatch"");' value=""将选择的文章提交批量执行"" id=""btnPost""/></p><form>"


	Response.Write "<hr/>" & ZC_MSG042 & ": " & strPage


Function Export_ArticleTag(ByVal strTagsCode)
	Call GetTags
	Dim t,i,s

	If strTagsCode<>"" Then

		strTagsCode=Replace(strTagsCode,"}","")
		t=Split(strTagsCode,"{")
		GetTagsbyTagIDList strTagsCode
		For i=LBound(t) To UBound(t)
			If t(i)<>"" Then
				s=s & "<p>" & Tags(t(i)).Name & "</p>"
			End If
		Next

	End If

	Export_ArticleTag=s

End Function

Function ExportPageBar(PageNow,PageAll,PageLength,Url)

If PageAll=0 Then
	Exit Function
End if

Dim s
Dim i

'Dim PageNow
'Dim PageAll
'Dim PageLength
Dim PageFrist
Dim PageLast
Dim PagePrevious
Dim PageNext
Dim PageBegin
Dim PageEnd

PageFrist = 1
PageLast = PageAll

PageBegin = PageNow
PageEnd = PageBegin + PageLength - 1

If PageEnd > PageAll Then
	PageEnd = PageAll
	PageBegin = PageAll - PageLength + 1
	If PageBegin < 1 Then
		PageBegin = 1
	End If
End If

s=s &"<a href='"&Url & PageFrist &"'>["& "&lt;&lt;" &"]</a> "

For i=PageBegin To PageEnd
	If i=PageNow Then
		s=s &"["& Replace(ZC_MSG036,"%s",i) &"] "
	Else
		s=s &"<a href='"&Url & i  &"'>["& Replace(ZC_MSG036,"%s",i) &"]</a> "
	End If
Next

s=s &"<a href='"&Url & PageLast  &"'>["& "&gt;&gt;" &"]</a> "

ExportPageBar=s

End Function
%>
</div>
</div>

<script language="JavaScript" type="text/javascript">

function BatchDelEnabled() {
	var checkBatchDel=document.getElementById('BatchDel');

	if(checkBatchDel.checked==true){
		document.getElementById('MoveCata').disabled=true;
		document.getElementById('EdtLevel').disabled=true;
		document.getElementById('EdtIstop').disabled=true;
		document.getElementById('EdtUser').disabled=true;
		document.getElementById('AddTag').disabled=true;
		document.getElementById('RmvTag').disabled=true;
		alert("注意: 您选择了批量删除功能, 此操作不可恢复, 请小心提交!");
	}
	else{
		document.getElementById('MoveCata').disabled=false;
		document.getElementById('EdtLevel').disabled=false;
		document.getElementById('EdtIstop').disabled=false;
		document.getElementById('EdtUser').disabled=false;
		document.getElementById('AddTag').disabled=false;
		document.getElementById('RmvTag').disabled=false;
	}
}

$(document).ready(function(){ 

	//斑马线
	var tables=document.getElementsByTagName("table");
	var b=false;
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");

		cells[0].className="color1";
		for (var i = 1; i < cells.length; i++){
			if(b){
				cells[i].className="color2";
				b=false;
			}
			else{
				cells[i].className="color3";
				b=true;
			};
		};
	}

});

</script>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>

