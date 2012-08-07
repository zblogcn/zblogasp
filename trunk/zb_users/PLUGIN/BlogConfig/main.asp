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
<!-- #include file="function.asp"-->
<%
Dim l
l=FilterSQL(Request.QueryString("name"))
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("BlogConfig")=False Then Call ShowError(48)
BlogTitle="配置管理"
Dim objRs,a,b,c,d,e,objC
Select Case Request("act")
	Case "open"
		ac
	Case "rename"
		objConn.Execute "UPDATE [blog_Config] SET [conf_Name]='"&FilterSQL(Request.QueryString("edit"))&"' WHERE [conf_Name]='"&l&"'"
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
	case "e_new"
		Set objC=New TConfig
		objC.Load l
	
	case "e_edit"
		Set objC=New TConfig
		objC.Load Request.Form("name2")
		objC.Write Request.Form("name1"),Request.Form("post")
		objC.Save
		
'		objC.Write Request.QueryString("newname"),
	
End Select
Call SetBlogHint_Custom(" 提示:本插件相当于Windows内的注册表编辑器，使用本插件请谨慎操作。一旦修改失误，很可能导致插件或博客无法打开！")
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="jquery.contextMenu.css" />
<style type="text/css">
table {
	table-layout: fixed;
	overflow:hidden
}
</style>
<script type="text/javascript" src="jquery.contextmenu.js"></script>
<script type="text/javascript">
var n;
n=false;
function read(a,b,c){
	console.log("open"+b);
	$("#content").html("Loading");
	$.get("main.asp",c,function(data){$("#content").html(data)});
	readleft();
	n=false;
}
function run2(a,b,c){
	var json={"act":"e_"+a,"name":b}
	switch (a){

	case "new":
		var j;
		j=$("#configt tr").last().children("td:first").children("input").attr("value");
		console.log(j);
		if(j=="NaN"){j=0}
		j=j+1;
		if(n==true){$("#configt").append("<tr><td></td><td>请保存后再新建</td><td></td></tr>");return false}
		$("#configt").append("<tr><td><input type='hidden' value='"+(j)+"'/><input type='text' id='txt"+(j)+"'></td><td><textarea id='ta"+(j)+" ' style='width:100%'/></td><td><a href='javascript:;' onclick='run2(\"edit\","+j+",$(this).parents(\"#content\").children(\"#name\").html())'><img src=\"../../../ZB_SYSTEM/image/admin/page_edit.png\" alt=\"编辑\" title=\"编辑\" width=\"16\" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp</td>");
		n=true;
//		$("#configt").append("<tr><td><input type=\"text\" id=\""+(tmp+1)+"\"/></td><td>-.-</td><td></tr>");
	break;
	case "edit":
		json["post"]=$("#ta"+b).attr("value");
		json["name1"]=$("#txt"+b).html();
		json["name2"]=$("#name").html()
//		json["test"]=b;
		console.log(json);
		$.post("main.asp",json,function(){read(json.act,json.name1,{"act":"open","name":json.name2})})
	break;
	}
}
function readleft(){
	var c={"act":"readleft"};
	console.log("readleft");
	$("#content").html("Loading");
	$.get("main.asp",c,function(data){$("#tree ul").html(data);$("#tree ul li").contextMenu({menu:'treemenu'},function(action, el, pos) {run(action,$(el).find("a").attr("id"))});})
	
}

function n(a){
		var c={"act":"new"}
		var d=prompt("请输入项名");
		if(d!=""&&d!=null){
			c["name"]=d;
			read(a,d,c);
		}else{return false}
}
	
function run(a,b){
	var c={"act":a,"name":b};
	if(b=="BlogConfig"){n(b);return false}
	switch (a){
		case "open":
		read(a,b,c)
		break;
		case "rename":
			var d=prompt("请输入新项名");
			if(d!=""&&d!=null){
				if(confirm("确定要把"+b+"改为"+d+"吗？\n\n请注意，盲目修改名字可能会导致某个插件或整个博客无法打开！")){
					c["edit"]=d;
					read(a,b,c)
				}else{return false}
			}else{return false}
		break;
		case "del":
			if(window.confirm("单击“确定”继续。单击“取消”停止。")){
				read(a,b,c);
			}
		break;

	}


}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">[管理] </span></a> </div>
          <div id="divMain2"> <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <div style="height:900px">
              <div style="float:left;width:10%;" name="tree" id="tree">
                <ul>
                  <%ReadLeft%>
                  
                </ul>
                <script type="text/javascript">$(document).ready(function() {$("#tree ul li").contextMenu({menu:'treemenu'},function(action, el, pos) {run(action,$(el).find("a").attr("id"))});});</script> 
              </div>
              <div id="content" style="float:right;width:88%;" >请选择</div>
            </div>
          </div>
        </div>
        <ul id="treemenu" class="contextMenu">
          <li class="open"> <a href="#open">打开</a> </li>
          <li class="rename"> <a href="#rename">重命名</a> </li>
          <li class="del"> <a href="#del">删除</a> </li>
        </ul>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
		Function ac
			Dim m
			
			m=l
			If m="BlogConfig" Then m=""
			Response.Write "<span id=""name"">"&m & "</span><a href=""javascript:;"" onclick=""run2('new','"& m&"')"">新建</a>"
			Set objRs=objConn.Execute("SELECT [conf_Name] AS A,[conf_Value] AS B FROM [blog_Config] WHERE [conf_Name]='"&l&"'")		
			Response.Write "<table width=""100%"" style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0' id=""configt""><tr><th width=""25%"">名称</th><th>内容</th><th width=""10%""></th></tr>"
			If Not(objRs.Eof) Then
				a=objRs("B")
				
				b=split(a,meta_split_string_2)
				If UBound(b)<=0 Then Response.Write "</table>":Response.End
				c=split(b(0),meta_split_string_1)
				d=split(b(1),meta_split_string_1)
				For e=1 To Ubound(c)
					Response.Write "<tr><td><input type='hidden' value='"&e&"'/><span id=""txt"&e&""">"&vbsunescape(c(e))&"</span></td><td onclick=""$('#ta"&e&"').show();$('#show"&e&"').hide()""><span id=""show"&e&""">"&vbsunescape(d(e))&"</span><textarea id=""ta"&e&""" style=""display:none;width:100%"">"&TransferHTML(vbsunescape(d(e)),"[textarea]")&"</textarea></td><td><a href=""javascript:;"" onclick=""run2('edit','"&e&"','"&m&"')""><img src=""../../../ZB_SYSTEM/image/admin/page_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm(""单击“确定”继续。单击“取消”停止。"");' href=""javascript:;"" onclick=""run2('del','"&e&"','"&m&"')""><img src=""../../../ZB_SYSTEM/image/admin/delete.png"" alt=""删除"" title=""删除"" width=""16"" /></a></td></tr>"
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
	End Function
%>