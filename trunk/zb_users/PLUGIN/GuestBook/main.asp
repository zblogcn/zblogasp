<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("GuestBook")=False Then Call ShowError(48)
BlogTitle="留言本"
Dim objConfig,a
Set objConfig=New TConfig
objConfig.Load("GuestBook")
If objConfig.Exists("v")=False Then
	objConfig.Write "v","1.0"
	objConfig.Write "g",0
	objConfig.Save
End If
If Request.QueryString("act")="save" Then
	a=CStr(Request.Form("id"))
	If a="0" Then
		Dim objArticle
		Set objArticle=New TArticle
		objArticle.FType=ZC_POST_TYPE_PAGE
		objArticle.AuthorID=BlogUser.ID
		objArticle.Content="欢迎给我留言"
		objArticle.Title="留言本"
		objArticle.Intro="欢迎给我留言"
		If objArticle.Post Then
			Call SetBlogHint_Custom("留言本生成完成！<a href="""&GetCurrentHost&"zb_system/cmd.asp?act=ArticleEdt&type=Page&webedit=ueditor&id="&a&""">点击这里去修改提示文字。</a>")
		End If
		a=objArticle.ID
	End If
	objConfig.Write "g",a
	objConfig.Save
	Call SaveToFile(BlogPath&"guestbook.asp",LoadFromFile(Server.MapPath("guestbook.asp"),"utf-8"),"utf-8",false)
	Call MakeBlogReBuild()
End If
Dim objRS
Set objRs=objConn.Execute("SELECT [log_ID] FROM [blog_Comment] WHERE [log_ID]=0")
If Not objRs.Eof Then Call SetBlogHint_Custom("检测到有1.8的留言未升级！请在下面指定一个页面后点击“迁移留言本”将1.8的留言升级到2.0！")
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">设定留言本页面</span></a><a href="b.asp"><span class="m-left">迁移留言本</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
<form id="form1" name="form1" method="post" action="?act=save">
<label for="id"></label>
<select name="id" size="1" id="id" style="width:100%">
<option value="0">新建页面</option>
<%

Set objRs=objConn.Execute("SELECT [log_ID],[log_Title] FROM [blog_Article] WHERE [log_Type]=0")
Do Until objRs.Eof
%>
<option value="<%=objRs("log_ID")%>"<%=IIf(CStr(objConfig.Read("g"))=CStr(objRs("log_ID"))," selected=""selected"" ","")%>><%=objRs("log_Title")%>（ID=<%=objRs("log_ID")%>）</option>
<%
objRs.MoveNext
Loop
Set objRs=Nothing%>
</select><br/><br/>
<input name="" type="submit" class="button" value="保存"/>
</form>
<p>若需要自定义标题、显示的内容以及模板，请选择“页面管理”。</p>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
