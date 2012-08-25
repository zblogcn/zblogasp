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
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("LoadCounter")=False Then Call ShowError(48)
BlogTitle="LoadCounter"
Select Case Request.QueryString("type")
	Case "header"
		Dim a
		Set a=New TCounter
		If a.LoadInfoById(Request.QueryString("id")) Then Response.Write TransferHTML(TransferHTML(a.AllRequestHeader,"[html-format]"),"[enter]")
		Response.End
	Case "postdata"
		Set a=New TCounter
		If a.LoadInfoById(Request.QueryString("id")) Then Response.Write TransferHTML(a.PostData,"[html-format]")
		Response.End
	Case "cleanlog"
		objConn.Execute "DELETE FROM [blog_Counter]"
		Response.Redirect "list.asp"
End Select
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
      
    
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
<a href="list.asp?type=cleanlog" onclick="return confirm('是否真要清除日志？')"><span class="m-left">清除所有日志</span></a>
  </div>
  <div id="divMain2">
 <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
    <%
	Call ExportCounterlist(Request("page"),Request("ip"),Request("agent"),vbsescape(Request("content")),vbsescape(Request("name")))
	%>

  </div>
</div>
<script type="text/javascript">
var a,k
$(document).ready(function(e) {
	$("a[name='httpheader']").bind("click",function(){
		var THIS=this;
		$.get("list.asp",{"id":$(this).attr("_temp"),"type":"header"},function(data){$(THIS).html(data)})
	})

	$("a[name='postdata']").bind("click",function(){
		var THIS=this
		$.get("list.asp",{"id":$(this).attr("_temp"),"type":"postdata"},function(data){$(THIS).html(data)})
	})

});
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%

Function ExportCounterlist(intPage,intCate,intLevel,intTitle,name)
	Dim i
	Dim objRS
	Dim strSQL,strPage
	Dim intPageAll
	Call CheckParameter(intPage,"int",1)
	intCate=vbsescape(intCate)
	intLevel=vbsescape(intLevel)
	name=vbsescape(name)
	Call CheckParameter(intTitle,"sql",-1)
	Call CheckParameter(name,"sql",-1)
	Dim tmp,tmp2
	tmp=TransferHTML(name,"[html-format]")
	tmp2=TransferHTML(intTitle,"[html-format]")
	Response.Write "<form id=""edit"" class=""search"" method=""post"" enctype=""application/x-www-form-urlencoded"" action=""list.asp"">"
	Response.Write "IP:<input type='text' name='ip' id='ip' value="""&TransferHTML(intCate,"[html-format]")&"""/>"
	Response.Write "  User-Agent:<input type='text' name='agent' id='agent' value='"&TransferHTML(intLevel,"[html-format]")&"'/>  日志类型： <input type='text' name='name' id='name' value='"&IIf(tmp="-1","",tmp)&"'/>  <input type=""submit"" class=""button"" value="""&ZC_MSG087&""">  "


	Response.Write "<br/><br/>PostData&AllHttp <input id=""content"" name=""content"" style=""width:70%"" type=""text"" value="""&IIf(tmp2="-1","",tmp)&""" /> "
	Response.Write ""
	Response.Write "</form>"
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	If intCate<>"" Then
		If ZC_MSSQL_ENABLE=False Then
			strSQL= strSQL & " AND InStr(1,[coun_IP],'"&intCate&"',0)<>0"
		Else
			strSQL= strSQL & " AND CHARINDEX('"&intCate&"',[coun_IP])<>0"
		End iF
	End If
	If intLevel<>"" Then
		iF zc_mssql_enable=false then
			strSQL= strSQL & " AND InStr(1,LCase([coun_Agent]),LCase('"&intLevel&"'),0<>0)"
		else
			strSQL= strSQL & " AND CHARINDEX([coun_Agent],'"&intLevel&"')<>0)"
		end if
	End If
	If intTitle<>"-1" Then
		If ZC_MSSQL_ENABLE=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([coun_AllRequestHeader]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([coun_Content]),LCase('" & intTitle &"'),0)<>0))"
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & intTitle &"',[coun_AllRequestHeader]))<>0) OR (CHARINDEX('" & intTitle &"',[coun_Content])<>0)"
		End If
	End If
	
	If Name<>"-1" Then
		If ZC_MSSQL_ENABLE=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([coun_logName]),LCase('" & Name &"'),0)<>0) "
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & Name &"',[coun_logName]))<>0)"
		End If
	End If
	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
	Response.Write "<tr><td>"& ZC_MSG076 &"</td><td>IP</td><td>类型</td><td>操作者</td><td>操作时间</td><td>操作内容</td><td>方法及URL</td><td>HTTP头</td><td>POSTDATA</td></tr>"
	If strsql<>"" then strsql="WHERE 1=1 "&strsql
	objRS.Open("SELECT * FROM [blog_Counter] "& strSQL &" ORDER BY [coun_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount
	
	If (Not objRS.bof) And (Not objRS.eof) Then
	
		For i=1 to objRS.PageSize
			If objRs.Eof Then Exit For
			Response.Write "<tr>"
			Response.Write "<td>" & objRS("coun_ID") & "</td>"
			Response.Write "<td>" & vbsunescape(objRS("coun_IP")) & "</td>"
			Response.Write "<td>" & vbsunescape(objRS("coun_logName")) & "</td>"
			Call GetUsersbyUserIDlist(objRS("coun_UserID"))
			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("coun_UserID") Then
						Response.Write "<td>" & User.Name & "</td>"
					End If
				End If
			Next
			Response.Write "<td>" & objRS("coun_PostTime") & "</td>"
			Response.Write "<td>" & TransferHTML(vbsunescape(objRS("coun_Content")),"[html-format]") & "</td>"
			Response.Write "<td>" & TransferHTML(vbsunescape(objRS("coun_URL")),"[html-format]") & "</td>"
			Response.Write "<td style='word-break: break-all;'><a href=""javascript:void(0)"" name=""httpheader"" _temp="""&objRS("coun_ID")&""">[查看]</a></td>"
			Response.Write "<td style='word-break: break-all;'><a href=""javascript:void(0)"" name=""postdata"" _temp="""&objRS("coun_ID")&""">[查看]</a></td>"
			objRs.MoveNext
		Next
	End If
	Response.Write "</table> "
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"list.asp?page="&Request("page")&"&ip="&Request("ip")&"&agent="&Request("agent")&"&content="&vbsescape(Request("content"))&"&name="&vbsescape(Request("name")))

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage

	Response.Write "</p></div>"
	objRS.Close
	Set objRS=Nothing
End Function
%>
