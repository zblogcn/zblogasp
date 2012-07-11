<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 
'// 作    者:    ZSXSOFT
'// 版权所有:    ZSXSOFT
'// 技术支持:   zsx@zsxsoft.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    edit_ueditor.asp
'// 开始时间:    2012.7.5
'// 最后修改:
'// 备    注:    编辑页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Edit_ueditor_Begin in Action_Plugin_Edit_ueditor_Begin
	If Not IsEmpty(sAction_Plugin_Edit_ueditor_Begin) Then Call Execute(sAction_Plugin_Edit_ueditor_Begin)
Next

'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("ArticleEdt") Then Call ShowError(6)

Dim EditArticle

Set EditArticle=New TArticle

If Not IsEmpty(Request.QueryString("id")) Then
	If EditArticle.LoadInfobyID(Request.QueryString("id")) Then
		If EditArticle.AuthorID<>BlogUser.ID Then
			If CheckRights("Root")=False Then
				Call ShowError(6)
			End If
		End If
	Else
		Call ShowError(9)
	End If
Else
	EditArticle.AuthorID=BlogUser.ID
End If

	On Error Resume Next
BlogTitle=EditArticle.HtmlUrl
EditArticle.Content=UBBCode(EditArticle.Content,"[link][email][font][code][face][image][flash][typeset][media][autolink][key][link-antispam]")

If Err.Number=0 Then

	EditArticle.Title=TransferHTML(EditArticle.Title,"[html-japan]")
	EditArticle.Content=TransferHTML(EditArticle.Content,"[html-japan]")
	EditArticle.Intro=TransferHTML(EditArticle.Intro,"[html-japan]")

	EditArticle.Title=TransferHTML(EditArticle.Title,"[html-format]")
	'EditArticle.Content=TransferHTML(EditArticle.Content,"[textarea]")
	'EditArticle.Intro=TransferHTML(EditArticle.Intro,"[textarea]")

Else

	GetCategory()
	GetUser()

	EditArticle.Title=EditArticle.Title
	EditArticle.Content=TransferHTML(EditArticle.Content,"[&]")
	EditArticle.Intro=TransferHTML(EditArticle.Intro,"[&]")

End If
If Request.QueryString("type")="tags" Then
	Response.Write "$(""#ajaxtags"").html("""
	Dim objRS
	Set objRS=objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] ORDER BY [tag_Name] ASC")
	If (Not objRS.bof) And (Not objRS.eof) Then
		Do While Not objRS.eof
			If InStr(EditArticle.Tag,"{"& objRS("tag_ID") & "}")>0 Then
				Response.Write "<a href='#' class='selected'>"& TransferHTML(Tags(objRS("tag_ID")).Name,"[html-format]") &"</a> "
			Else
				Response.Write "<a href='#'>"& TransferHTML(Tags(objRS("tag_ID")).Name,"[html-format]") &"</a> "
			End If
			objRS.MoveNext
		Loop
	End If
	objRS.Close
	Set objRS=Nothing
	Response.Write """);$(""#ulTag"").tagTo(""#edtTag"");"
	Response.End
End If
Err.Clear

BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & IIf(Request.QueryString("type")="Page",ZC_MSG329,ZC_MSG047)

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../script/common.js" type="text/javascript"></script>
    <link rel="stylesheet" type="text/css" href="ueditor/themes/default/ueditor.css"/>
	<link rel="stylesheet" href="../CSS/jquery.bettertip.css" type="text/css" media="screen">
	<script language="JavaScript" src="../script/jquery.bettertip.pack.js" type="text/javascript"></script>
	<script language="JavaScript" src="../script/jquery.tagto.js" type="text/javascript"></script>

    <script type="text/javascript" charset="utf-8" src="ueditor/editor_config.asp"></script>
    <script type="text/javascript" charset="utf-8" src="ueditor/editor_all_min.js"></script>
	<script language="JavaScript" type="text/javascript">var loaded=false;var editor = new baidu.editor.ui.Editor();</script>

	<title><%=BlogTitle%></title>
</head>
<body>
<div id="divMain">
<div class="Header"><%=IIf(Request.QueryString("type")="Page",ZC_MSG329,ZC_MSG047)%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleEdt_SubMenu & "</div>"
%>
<div class="form">
<% Call GetBlogHint() %>
<div id="divClick" style="display:none;"><a href="#" onClick="document.getElementById('divClick').style.display='none';document.getElementById('divAdv').style.display='block';document.getElementById('divFileSnd').style.display='block';document.getElementById('divIntro').style.display='block';Advanced();return false;"><%=GetSettingFormNameWithDefault("ZC_MSG316","Advanced Option&gt;&gt;")%><span style="font-size: 1.5em; vertical-align: -1px;">»</span></a></div>
<form id="edit" name="edit" method="post">
	<input type="hidden" name="edtID" id="edtID" value="<%=EditArticle.ID%>">
	<p><%=ZC_MSG060%>:<input type="text" name="edtTitle" id="edtTitle" style="width:367px;"  onblur="if(this.value=='') this.value='<%=ZC_MSG099%>'" onFocus="if(this.value=='<%=ZC_MSG099%>') this.value=''" value="<%=EditArticle.Title%>" />
<%
On Error Resume Next
BlogTitle=EditArticle.Alias
If Err.Number=0 Then
%>
	&nbsp;<%=ZC_MSG147%>:<input type="text" style="width:220px;" name="edtAlias" id="edtAlias" value="<%=TransferHTML(EditArticle.Alias,"[html-format]")%>"> .<%=ZC_STATIC_TYPE%>
<%
End If
Err.Clear
%>
	</p>
<%
Err.Clear
On Error Resume Next
If Request.QueryString("type")<>"Page" Then
BlogTitle=EditArticle.Tag

If Err.Number=0 Then
%>
	<p><%=ZC_MSG012%>:<select style="width:140px;" class="edit" size="1" id="cmbCate" onChange="edtCateID.value=this.options[this.selectedIndex].value"><option value="0"></option>
<%
	'GetCategory()
	Dim aryCateInOrder : aryCateInOrder=GetCategoryOrder()
	Dim m,n
	For m=0 To Ubound(aryCateInOrder)
		If Categorys(aryCateInOrder(m)).ParentID=0 Then
			Response.Write "<option value="""&Categorys(aryCateInOrder(m)).ID&""" "
			If EditArticle.CateID=Categorys(aryCateInOrder(m)).ID Then Response.Write "selected=""selected"""
			Response.Write ">"&TransferHTML( Categorys(aryCateInOrder(m)).Name,"[html-format]")&"</option>"

			For n=0 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(n)).ParentID=Categorys(aryCateInOrder(m)).ID Then
					Response.Write "<option value="""&Categorys(aryCateInOrder(n)).ID&""" "
					If EditArticle.CateID=Categorys(aryCateInOrder(n)).ID Then Response.Write "selected=""selected"""
					Response.Write ">&nbsp;┄ "&TransferHTML( Categorys(aryCateInOrder(n)).Name,"[html-format]")&"</option>"
				End If
			Next
		End If
	Next
%>
	</select><input type="hidden" name="edtCateID" id="edtCateID" value="<%=EditArticle.CateID%>">
	
 <%
 End If
 Else
%><input type="hidden" name="edtCateID" id="edtCateID" value="0"><%
End If
%>&nbsp;<%=ZC_MSG003%>:<select style="width:100px;" class="edit" size="1" id="cmbUser" onChange="edtAuthorID.value=this.options[this.selectedIndex].value"><option value="0"></option>
<%
	GetUser()
	Dim User
	For Each User in Users
		If IsObject(User) Then
			If CheckRights("Root")=True Then
				Response.Write "<option value="""&User.ID&""" "
				If User.ID=EditArticle.AuthorID Then
					Response.Write "selected=""selected"""
				End If
				Response.Write ">"&TransferHTML(User.Name,"[html-format]")&"</option>"
			Else
				If User.ID=EditArticle.AuthorID Then
					Response.Write "<option value="""&User.ID&""" "
					Response.Write "selected=""selected"""
					Response.Write ">"&TransferHTML(User.Name,"[html-format]")&"</option>"
				End If
			End If
		End If
	Next

%>
	</select><input type="hidden" name="edtAuthorID" id="edtAuthorID" value="<%=EditArticle.AuthorID%>">
<%
If Request.QueryString("type")<>"Page" Then
	%>&nbsp;<%=ZC_MSG138%>:<input type="text" style="width:313px;" name="edtTag" id="edtTag" value="<%=TransferHTML(EditArticle.TagToName,"[html-format]")%>"> <a href="" style="cursor:pointer;" onClick="if(document.getElementById('ulTag').style.display=='none'){document.getElementById('ulTag').style.display='block';if(loaded==false){$.getScript('edit_ueditor.asp?type=tags<%if EditArticle.id<>0  then response.write "&id="&EditArticle.ID%>');loaded=true;}}else{document.getElementById('ulTag').style.display='none'};return false;"><%=ZC_MSG139%><span style="font-size: 1.5em; vertical-align: -1px;"></span></a>
	<ul id="ulTag" style="display:none;">
    <span id="ajaxtags"><%=ZC_MSG326%></span>

	&nbsp;&nbsp;(<%=ZC_MSG296%>)</ul></p>
<%
Else
%>
<input name="edtAutoList" id="edtAutoList" type="checkbox" value="True" checked="checked"/><label for="edtAutoList">自动加入导航条</label>

<%
End If
Err.Clear
%>
<div id="divAdv" style="display:block;">
<p><%=ZC_MSG061%>:<select style="width:140px;" class="edit" size="1" id="cmbArticleLevel" onChange="edtLevel.value=this.options[this.selectedIndex].value">
<%
	Dim ArticleLevel
	Dim i:i=0
	For Each ArticleLevel in ZVA_Article_Level_Name
		Response.Write "<option value="""& i &""" "
		If EditArticle.Level=i Then Response.Write "selected=""selected"""
		Response.Write ">"& Replace(ZVA_Article_Level_Name(i),ZC_MSG048,ZC_MSG330) &"</option>"
		i=i+1
	Next
%>
	</select><input type="hidden" name="edtLevel" id="edtLevel" value="<%=EditArticle.Level%>" />
<%
Err.Clear
On Error Resume Next
BlogTitle=EditArticle.Istop
If Request.QueryString("type")<>"Page" Then
If Err.Number=0 Then

%>
&nbsp;<%=ZC_MSG051%>
<%If EditArticle.Istop Then%>
<input type="checkbox" name="edtIstop" id="edtIstop" value="True" checked=""/>
<%Else%>
<input type="checkbox" name="edtIstop" id="edtIstop" value="True"/>
<%End If%>
<%

End If
Else
%><input type="hidden" name="edtIstop" id="edtIstop" value=""/><%
End If
Err.Clear
%>
	&nbsp;<%=ZC_MSG062%>:<input type="text" name="edtYear" id="edtYear" style="width:35px;" value="<%=Year(EditArticle.PostTime)%>" />-<input type="text" name="edtMonth" id="edtMonth" style="width:25px;" value="<%=Month(EditArticle.PostTime)%>" />-<input type="text" name="edtDay" id="edtDay" style="width:25px;" value="<%=Day(EditArticle.PostTime)%>" />-<input type="text" name="edtTime" id="edtTime" style="width:50px;" value="<%= Hour(EditArticle.PostTime)&":"&Minute(EditArticle.PostTime)&":"&Second(EditArticle.PostTime)%>" />
	<%
Err.Clear
%>
&nbsp;<%=ZC_MSG324%>:<select style="width:150px;" class="edit" size="1" id="cmbTemplate" onChange="edtTemplate.value=this.options[this.selectedIndex].value">
<%
	'Response.Write "<option value="""">"&ZC_MSG325&"</option>"

	Dim aryFileList

	aryFileList=LoadIncludeFiles("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		Dim j,t
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If EditArticle.TemplateName=t Then
				Response.Write "<option value="""&t&""" selected=""selected"">"&t&"</option>"
			Else
				Response.Write "<option value="""&t&""">"&t&"</option>"
			End If
		Next
	End If

	If EditArticle.TemplateName="" Then
	%><option value="" selected="selected"><%=ZC_MSG325%>(SINGLE)</option><%
	Else
	%><option value=""><%=ZC_MSG325%>(SINGLE)</option><%
	End If
%>
</select><input type="hidden" name="edtTemplate" id="edtTemplate" value="<%=EditArticle.TemplateName%>" />

</div>

<%
If Response_Plugin_Edit_Form<>"" Then
%>
<div><%=Response_Plugin_Edit_Form%></div>
<%
End If
%>


<!-- <div id="divFileSnd">
<%If CheckRights("FileSnd") Then%>
	<p id="filesnd"><iframe frameborder="0" height="56" marginheight="0" marginwidth="0" scrolling="no" width="100%" src="../cmd.asp?act=FileSnd"></iframe></p>
<%Else%>
<%End If%>
</div> -->

<div id="divContent" style="clear:both;">
<p><%=ZC_MSG055%>:(<span id="timemsg"></span><span id="msg2"></span><span id="msg"></span><SCRIPT LANGUAGE="JavaScript" src="c_autosaverjs.asp?act=edit&type=ueditor"></SCRIPT>)<br/></p>

<textarea id="ueditor" name="txaContent"><%=EditArticle.Content%></textarea>
	
</div>
<p></p>
<div id="divAutoIntro" class="anti_normal" style="display:<%If EditArticle.ID=0 And EditArticle.Intro="" Then Response.Write "block" Else Response.Write "none"%>;" onClick="try{this.style.display='none';document.getElementById('divIntro').style.display='block';AutoIntro();}catch(e){}"><p><a title="<%=ZC_MSG297%>" href="javascript:AutoIntro()">[<%=ZC_MSG310%>]</a></p></div>
<div id="divIntro" style="display:<%If EditArticle.Intro="" Then Response.Write "none" Else Response.Write "block"%>;">
<p><%=ZC_MSG016%>:<a title="<%=ZC_MSG297%>" href="javascript:AutoIntro()">[<%=ZC_MSG310%>]</a></p>
<textarea id="ueditor2" name="txaIntro"><%=EditArticle.Intro%></textarea>

	
</div>
<%
If Response_Plugin_Edit_Form2<>"" Then
%>
<div><%=Response_Plugin_Edit_Form2%></div>
<%
End If
%>
	<p><input class="button" type="submit" value="<%=ZC_MSG087%>" id="btnPost" onclick='return checkArticleInfo();' /></p>

</form>
</div>

			</div>
<script type="text/javascript">

	$(document).ready(function(){

		editor.render('ueditor');

		var editor2 = new baidu.editor.ui.Editor({
			toolbars:[['Source', 'Undo', 'Redo']],
			autoClearinitialContent:false,
			wordCount:false,
			elementPathEnabled:false,
			autoFloatEnabled:false,
			autoHeightEnabled:false,
			minFrameHeight:200
		});

		editor2.render('ueditor2');

	});

	var str10="<%=ZC_MSG115%>";
	var str11="<%=ZC_MSG116%>";
	var str12="<%=ZC_MSG117%>";

	function checkArticleInfo(){
		document.getElementById("edit").action="../cmd.asp?act=ArticlePst&webedit=ueditor<%=IIF(Request.QueryString("type")="Page","&type=Page","")%>";

<%If Request.QueryString("type")="" Then%>
		if(document.getElementById("edtCateID").value==0){
			alert(str10);
			return false
		}
<%End If%>
		if(!editor.getContent()){
			alert(str11);
			return false
		}
	}

	function AddKey(i) {
		var strKey=document.getElementById("edtTag").value;
		var strNow=","+i

		if(strKey==""){
			strNow=i
		}

		if(strKey.indexOf(strNow)==-1){
			strKey=strKey+strNow;
		}
		document.getElementById("edtTag").value=strKey;
	}
	function DelKey(i) {
		var strKey=document.getElementById("edtTag").value;
		var strNow="{"+i+"}"
		if(strKey.indexOf(strNow)!=-1){

			strKey=strKey.substring(0,strKey.indexOf(strNow))+strKey.substring(strKey.indexOf(strNow)+strNow.length,strKey.length)

		}
		document.getElementById("edtTag").value=strKey;
	}

	function AutoIntro() {
		editor2.setContent(editor.getContent().replace(/<[^>]+>/g, "").substring(0,200) )
	}

	function Advanced(){
		$("div.normal").css("display","block");
		$("div.anti_normal").css("display","none");
	}


</script>
</body>

</html>
<%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>