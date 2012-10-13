<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'该插件代码其乱无比，还是别看为好。
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\function.asp"-->
<%
Dim l

l=FilterSQL(Request.QueryString("name"))

Call System_Initialize()

Call CheckReference("")

If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("ZBDK")=False Then Call ShowError(48)

BlogTitle=title
Dim objRs,a,b,c,d,e,objC
Select Case Request("act")
	Case "open"
		ac
	Case "rename"
		objConn.Execute "UPDATE [blog_Config] SET [conf_Name]='"&FilterSQL(Request.QueryString("edit"))&"' WHERE [conf_Name]='"&l&"'"
		l=FilterSQL(Request.QueryString("edit"))
		ac
	Case "readleft"
		readleft
		response.end
	case "del"
		objConn.Execute "DELETE FROM [blog_Config] WHERE [conf_Name]='"&l&"'"
		Response.Write "删除成功"
		Response.End
	case "new"
		Set objRs=objConn.Execute ("SELECT * FROM [blog_Config] WHERE [conf_Name]='"&l&"'")
		If objRs.Eof Then
			objConn.Execute "INSERT INTO [blog_Config] VALUES ('"&l&"','')"
			ac
		Else
			ac
		End If
	case "e_del"
		Set objC=New TConfig
		objC.Load Request.Form("name2")
		objC.Remove Request.Form("name1")
		objC.Save
		l=FilterSQL(Request.Form("name2"))
		ac
	case "e_edit"
		Set objC=New TConfig
		objC.Load Request.Form("name2")
		objC.Write Request.Form("name1"),Request.Form("post")
		objC.Save
		l=FilterSQL(Request.Form("name2"))
		ac	
End Select
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="../css/BlogConfig.css" type="text/css" media="screen"/>
<link rel="stylesheet" href="../css/jquery.contextMenu.css" type="text/css" media="screen"/>
<script type="text/javascript" src="../script/jquery.contextMenu.js"></script>
<script type="text/javascript" src="../script/BlogConfig.js"></script>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=ZBDK.submenu(1)%></a> </div>
          <div id="divMain2"> <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <div class="DIVBlogConfig">
              <div class="DIVBlogConfignav" name="tree" id="tree">
                <ul>
                  <%ReadLeft%>
                </ul>
                <script type="text/javascript">$(document).ready(function() {$("#tree ul li").contextMenu({menu:'treemenu'},function(action, el, pos) {run(action,$(el).find("a").attr("id"))});});</script></div>
              <div id="content" class="DIVBlogConfigcontent">
                <div class="DIVBlogConfigcontentbody">请选择</div>
              </div>
              <div class="clear"></div>
            </div>
          </div>
        </div>
        <ul id="treemenu" class="contextMenu">
          <li class="open"> <a href="#open">打开</a> </li>
          <li class="rename"> <a href="#rename">重命名</a> </li>
          <li class="del"> <a href="#del">删除</a> </li>
        </ul>
        <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
<%
		Function ac
			Dim m,n,h
			m=l
			If m="BlogConfig" Then m=""
			Response.Write "<div class=""DIVBlogConfigtop""><span id=""name"">"&m & "</span><a href=""javascript:;"" onclick=""run2('new','"& m&"')"">新建</a></div>"
			Set objRs=objConn.Execute("SELECT [conf_Name] AS A,[conf_Value] AS B FROM [blog_Config] WHERE [conf_Name]='"&l&"'")	
			Response.Write "<table width=""100%"" style='padding:0px;' cellspacing='0' cellpadding='0' id=""configt""><tr><th width=""25%"">名称</th><th>内容</th><th width=""10%""></th></tr>"
			If Not objRs.Eof Then
				a=objRs("B")
				b=split(a,meta_split_string_2)
				If UBound(b)<=0 Then Response.Write "</table>":Response.End
				c=split(b(0),meta_split_string_1)
				d=split(b(1),meta_split_string_1)
				For e=1 To Ubound(c)
					n=TransferHTML(vbsunescape(d(e)),"[textarea]")
					h=TransferHTML(vbsunescape(c(e)),"[textarea]")
					Response.Write "<tr><td><input type='hidden' value='"&e&"'/><span id=""txt"&e&""">"&h&"</span></td><td onclick=""$('#ta"&e&"').show();$('#show"&e&"').hide()""><span id=""show"&e&""">"&n&"</span><textarea id=""ta"&e&""" style=""display:none;width:100%"">"&n&"</textarea></td><td><a href=""javascript:;"" onclick=""run2('edit','"&e&"','"&m&"')""><img src=""../../../../ZB_SYSTEM/image/admin/page_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='if( window.confirm(""单击“确定”继续。单击“取消”停止。"")){run2(""del"","""&e&""","""&m&""")};' href=""javascript:;"" onclick=""run2('del','"&e&"','"&m&"')""><img src=""../../../../ZB_SYSTEM/image/admin/delete.png"" alt=""删除"" title=""删除"" width=""16"" /></a></td></tr>"
				Next
			End If
			Response.Write "</table>"
			Response.End
	End Function

Function ReadLeft
	Set objRs=objConn.Execute("SELECT [conf_Name] FROM [blog_Config]")
	Do Until objRs.Eof
		Response.Write "<li><a id="""&objRs("conf_Name")&""" href=""javascript:;"" onclick=""run('open','"&objRs("conf_Name")&"')"">" & objRs("conf_Name") & "</a></li>"
		objRs.MoveNext
	Loop
	Response.Write "<li><a id=""BlogConfig"" href=""javascript:;"" onClick=""run('open','BlogConfig')"">新建</a></li>"
End Function%>