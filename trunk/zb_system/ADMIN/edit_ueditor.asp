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

GetCategory()
GetUser()

Dim IsPage

IsPage=Request.QueryString("type")="Page"

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

		'ajax tags

		If Request.QueryString("type")="tags" Then
			Response.Write "$(""#ajaxtags"").html("""
			Dim objRS
			Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name] FROM [blog_Tag] ORDER BY [tag_Count] DESC")
			If (Not objRS.bof) And (Not objRS.eof) Then
				Do While Not objRS.eof
					If InStr(EditArticle.Tag,"{"& objRS("tag_ID") & "}")>0 Then
						Response.Write "<a href='#' class='selected'>"& TransferHTML(objRS("tag_Name"),"[html-format]") &"</a> "
					Else
						Response.Write "<a href='#'>"& TransferHTML(objRS("tag_Name"),"[html-format]") &"</a> "
					End If
					objRS.MoveNext
				Loop
			End If
			objRS.Close
			Set objRS=Nothing
			Response.Write """);$(""#ulTag"").tagTo(""#edtTag"");"
			Response.End
		End If

BlogTitle=EditArticle.HtmlUrl
EditArticle.Content=UBBCode(EditArticle.Content,"[link][email][font][code][face][image][flash][typeset][media][autolink][key][link-antispam]")

EditArticle.Title=TransferHTML(EditArticle.Title,"[html-japan]")
EditArticle.Content=TransferHTML(EditArticle.Content,"[html-japan]")
EditArticle.Intro=TransferHTML(EditArticle.Intro,"[html-japan]")

EditArticle.Title=TransferHTML(EditArticle.Title,"[html-format]")

Call GetTagsbyTagIDList(EditArticle.Tag)

BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & IIf(IsPage,ZC_MSG329,ZC_MSG047)

%>
<!--#include file="admin_header.asp"-->
	<link rel="stylesheet" type="text/css" href="ueditor/themes/default/ueditor.css"/>
	<script type="text/javascript" src="../script/jquery.tagto.js"></script>
	<script type="text/javascript" charset="utf-8" src="ueditor/editor_config.asp"></script>
	<script type="text/javascript" charset="utf-8" src="ueditor/editor_all_min.js"></script>
	<script type="text/javascript">
	var loaded=false;
	var editor = new baidu.editor.ui.Editor();
	var editor2 = new baidu.editor.ui.Editor({
		toolbars:[['Source', 'Undo', 'Redo']],
		autoClearinitialContent:false,
		wordCount:false,
		elementPathEnabled:false,
		autoFloatEnabled:false,
		autoHeightEnabled:false,
		minFrameHeight:200
	});</script>
<!--#include file="admin_top.asp"-->
<%If IsPage=False Then%>
<script type="text/javascript">ActiveLeftMenu("aArticleEdt");</script>
<%Else%>
<script type="text/javascript">ActiveLeftMenu("aPageMng");</script>
<%End If%>
                <div id="divMain">
<%	Call GetBlogHint()	%>
<div class="divHeader2"><%=IIf(IsPage,ZC_MSG329,ZC_MSG047)%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleEdt_SubMenu & "</div>"
%>
                  <div id="divMain2">
                    <form id="edit" name="edit" method="post" action="">
<div id="divEditLeft">

<div id="divEditTitle">
                      <input type="hidden" name="edtID" id="edtID" value="<%=EditArticle.ID%>" />
<!-- title -->
                      <p><span class='editinputname'><%=ZC_MSG060%>:</span>
                        <input type="text" name="edtTitle" id="edtTitle" style="width:400px"  onblur="if(this.value=='') this.value='<%=ZC_MSG099%>'" onfocus="if(this.value=='<%=ZC_MSG099%>') this.value=''" value="<%=EditArticle.Title%>" /></p>
                        <%
On Error Resume Next
If Err.Number=0 Then
%>
<!-- alias -->
                        <p><span class='editinputname'><%=ZC_MSG147%>:</span>
                        <input type="text" style="width:400px" name="edtAlias" id="edtAlias" value="<%=TransferHTML(EditArticle.Alias,"[html-format]")%>" />
                        .<%=ZC_STATIC_TYPE%>
<%
End If
Err.Clear
%>
                      </p>

<!-- tags( -->
<% If Request.QueryString("type")<>"Page" Then %>
                        <p><span class='editinputname' style='padding-right:2px;'><%=ZC_MSG138%>:</span>
                        <input type="text" style="width:400px;" name="edtTag" id="edtTag" value="<%=TransferHTML(EditArticle.TagToName,"[html-format]")%>" />
                        <a href="" style="cursor:pointer;" onclick="if(document.getElementById('ulTag').style.display=='none'){document.getElementById('ulTag').style.display='block';if(loaded==false){$.getScript('edit_ueditor.asp?type=tags');loaded=true;}}else{document.getElementById('ulTag').style.display='none'};return false;"><%=ZC_MSG139%><span style="font-size: 1.5em; vertical-align: -1px;"></span></a></p>
                      <ul id="ulTag" style="display:none;">
                        <li><span id="ajaxtags"><%=ZC_MSG326%></span>&nbsp;&nbsp;(<%=ZC_MSG296%>)</li>
                      </ul>
                      
<% End If %>
<!-- )tags -->



</div>





<% If Response_Plugin_Edit_Form<>"" Then %>
<div id="divEditForm1"><%=Response_Plugin_Edit_Form%></div>
<% End If %>
                      
                      <hr/>
                      <div id="divContent" style="clear:both;">
						<!-- <p><span class='editinputname'><%=ZC_MSG055%>:</span></p> -->
                        <textarea id="ueditor" name="txaContent" rows="50" cols="250"><%=EditArticle.Content%></textarea>
						<p style="text-align:right;"><span><a title="<%=ZC_MSG297%>" href="" onclick="try{document.getElementById('divIntro').style.display='block';AutoIntro();return false;}catch(e){}">[<%=ZC_MSG310%>]</a></span><span class='editinputname'></span><span id="timemsg"></span><span id="msg2"></span><span id="msg"></span><span class='editinputname'></span><script type="text/javascript" src="c_autosaverjs.asp?act=edit&amp;type=ueditor"></script></p>
                      </div>
                      <hr/>




                      <div id="divIntro" style="display:<%If EditArticle.Intro="" Then Response.Write "none" Else Response.Write "block"%>;">
                        <p><span class='editinputname'><%=ZC_MSG016%>:<!--<a title="<%=ZC_MSG297%>" href="javascript:AutoIntro()">[<%=ZC_MSG310%>]</a>--></span></p>
                        <textarea id="ueditor2" name="txaIntro" rows="50" cols="250"><%=EditArticle.Intro%></textarea>
                      </div>

<% If Response_Plugin_Edit_Form2<>"" Then %>
<div id="divEditForm2"><%=Response_Plugin_Edit_Form2%></div>
<% End If %>



</div><!-- divEditLeft -->

<div id="divEditRight">


<div id="divEditPost">

<p>
<input class="button" type="submit" value="<%=ZC_MSG087%>" id="btnPost" onclick='return checkArticleInfo();' />
</p>



                        <p>
<!-- cate -->
                      <%
Err.Clear
On Error Resume Next
If Request.QueryString("type")<>"Page" Then

If Err.Number=0 Then
%>
                      <span class='editinputname' style="cursor:pointer;" onclick="$(this).next().toggleClass('hidden');"><%=ZC_MSG012%>:</span>
                        <select style="width:140px;" class="edit" size="1" id="cmbCate" onchange="edtCateID.value=this.options[this.selectedIndex].value">
                          <option value="0"></option>
                          <%
	'GetCategory()
	Dim aryCateInOrder : aryCateInOrder=GetCategoryOrder()
	Dim m,n
	For m=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
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
                        </select>
                        <input type="hidden" name="edtCateID" id="edtCateID" value="<%=EditArticle.CateID%>" />
<%
 End If
 Else
%>
                        <input type="hidden" name="edtCateID" id="edtCateID" value="0" />
                        <%
End If
%>
<!-- cate -->
                      </p>



                        <p>
<!-- template( -->
                          <span class='editinputname' style="cursor:pointer;" onclick="$(this).next().toggleClass('hidden');"><%=ZC_MSG324%>:</span>
                          <select style="width:150px;" class="edit" size="1" id="cmbTemplate" onchange="edtTemplate.value=this.options[this.selectedIndex].value">
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
	%>
                            <option value="" selected="selected"><%=ZC_MSG325%>(SINGLE)</option>
                            <%
	Else
	%>
                            <option value=""><%=ZC_MSG325%>(SINGLE)</option>
                            <%
	End If
%>
                          </select>
                          <input type="hidden" name="edtTemplate" id="edtTemplate" value="<%=EditArticle.TemplateName%>" />
<!-- )template -->
                      </p>



                        <p>
<!-- level -->
                          <span class='editinputname' style="cursor:pointer;" onclick="$(this).next().toggleClass('hidden');"><%=ZC_MSG061%>:</span><select class="edit" style="width:140px;" size="1" id="cmbArticleLevel" onchange="edtLevel.value=this.options[this.selectedIndex].value">
                            <%
	Dim ArticleLevel
	Dim i:i=0
	For Each ArticleLevel in ZVA_Article_Level_Name
		Response.Write "<option value="""& i &""" "
		If EditArticle.Level=i Then Response.Write "selected=""selected"""
		Response.Write ">"& ZVA_Article_Level_Name(i) &"</option>"
		i=i+1
	Next
%>
                          </select>
                          <input type="hidden" name="edtLevel" id="edtLevel" value="<%=EditArticle.Level%>" />
                          <%
Err.Clear
%>
<!-- )level -->
                      </p>



                        <p>
<!-- user( -->

                        <span class='editinputname' style="cursor:pointer;" onclick="$(this).next().toggleClass('hidden');"><%=ZC_MSG003%>:</span><select style="width:100px;" size="1" id="cmbUser" onchange="edtAuthorID.value=this.options[this.selectedIndex].value">
                          <option value="0"></option>
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
                        </select>
                        <input type="hidden" name="edtAuthorID" id="edtAuthorID" value="<%=EditArticle.AuthorID%>" />
<!-- )user -->
                      </p>




                        <p>
<!-- date( -->
                          <span class='editinputname' style="cursor:pointer;" onclick="$(this).next().toggleClass('hidden');"><%=ZC_MSG062%>:</span><span><input type="text" name="edtYear" id="edtYear" style="width:30px;" value="<%=Year(EditArticle.PostTime)%>" /><span>-</span><input type="text" name="edtMonth" id="edtMonth" style="width:15px;" value="<%=Month(EditArticle.PostTime)%>" /><span>-</span><input type="text" name="edtDay" id="edtDay" style="width:15px;" value="<%=Day(EditArticle.PostTime)%>" /><span>-</span><input type="text" name="edtTime" id="edtTime" style="width:50px;" value="<%= Hour(EditArticle.PostTime)&":"&Minute(EditArticle.PostTime)&":"&Second(EditArticle.PostTime)%>" /></span>
<!-- )date -->

                      </p>



                        <p>
<!-- Istop( -->
<% If Request.QueryString("type")<>"Page" Then %>
                          <span class='editinputname'><%=ZC_MSG051%>
                          <%If EditArticle.Istop Then%>
                          <input type="checkbox" name="edtIstop" id="edtIstop" value="True" checked=""/>
                          <%Else%>
                          <input type="checkbox" name="edtIstop" id="edtIstop" value="True"/></span>
                          <%End If%>
<%Else%>
                          <input type="hidden" name="edtIstop" id="edtIstop" value=""/>
<% End If %>
<!-- )Istop -->
                      </p>



<% If Request.QueryString("type")="Page" Then %>
                      <label for="edtAutoList">自动加入导航条 </label><input name="edtAutoList" id="edtAutoList" type="checkbox" value="" />
                      
<% End If %>


                      </p>









</div>


</div><!-- divEditRight -->



                    </form>
                  </div>
</div>
<script type="text/javascript">
// <![CDATA[

	editor.render('ueditor');
	editor2.render('ueditor2');

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
		editor2.setContent(editor.getContent().replace(/<[^>]+>/g, "").substring(0,200) );
		$("#divIntro").show()
	}

	function Advanced(){
		$("div.normal").css("display","block");
		$("div.anti_normal").css("display","none");
	}
// ]]>
</script>
<!--#include file="admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>