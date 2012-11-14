<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBDK")=False Then Call ShowError(48)
BlogTitle=title
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
		Response.Redirect "main.asp"
End Select
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
      
    
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%=ZBDK.submenu(2)%>
  </div>
  <div id="divMain2">
 <script type="text/javascript">ActiveTopMenu("zbdk");</script>
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
		$.get("main.asp",{"id":$(this).attr("_temp"),"type":"header"},function(data){$(THIS).html(data)})
	})

	$("a[name='postdata']").bind("click",function(){
		var THIS=this
		$.get("main.asp",{"id":$(this).attr("_temp"),"type":"postdata"},function(data){$(THIS).html(data)})
	})

});
</script>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->

<%

Function ExportCounterlist(intPage,strIP,strAgent,strAQH,name)
	Dim i
	Dim objRS
	Dim strSQL,strPage
	Dim intPageAll
	Call CheckParameter(intPage,"int",1)
	strIP=vbsescape(strIP)
	strAgent=vbsescape(strAgent)
	name=vbsescape(name)
	Call CheckParameter(strAQH,"sql",-1)
	Call CheckParameter(name,"sql",-1)
	Dim tmp,tmp2
	tmp=TransferHTML(name,"[html-format]")
	tmp2=TransferHTML(strAQH,"[html-format]")
	Response.Write "<form id=""edit"" class=""search"" method=""post"" enctype=""application/x-www-form-urlencoded"" action=""main.asp"">"
	Response.Write "IP:<input type='text' name='ip' id='ip' value="""&vbsunescape(TransferHTML(strIP,"[html-format]"))&"""/>"
	Response.Write "  User-Agent:<input type='text' name='agent' id='agent' value='"&vbsunescape(TransferHTML(strAgent,"[html-format]"))&"'/>  日志类型： <input type='text' name='name' id='name' value='"&vbsunescape(IIf(tmp="-1","",tmp))&"'/>  <input type=""submit"" class=""button"" value="""&ZC_MSG087&""">  "
	Response.Write "<br/><br/>PostData&AllHttp <input id=""content"" name=""content"" style=""width:70%"" type=""text"" value="""&vbsunescape(IIf(tmp2="-1","",tmp2))&""" /> "
	Response.Write ""
	Response.Write "</form>"
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	
	If strIP<>"" Then
		strSQL = strSQL & ExportSearch(strIP,"coun_IP")
	End If
	
	If strAgent<>"" Then
		strSQL = strSQL & ExportSearch(strAgent,"coun_Agent")
	End If
	
	
	If strAQH<>"-1" Then
		strSQL = strSQL & ExportSearch(strAQH,"coun_AllRequestHeader") & " OR (1=1 " & ExportSearch(strAQH,"coun_Content")  & ")"
	End If
	
	If Name<>"-1" Then
		strSQL = strSQL & ExportSearch(Name,"coun_LogName")
	End If
	
	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"" height=""40"">"
	Response.Write "<tr><td>"& ZC_MSG076 &"</td><td>IP</td><td>类型</td><td>操作者</td><td>操作时间</td><td>操作内容</td><td>方法及URL</td><td>HTTP头</td><td>POSTDATA</td></tr>"
	If strsql<>"" then strsql="WHERE 1=1 "&strsql
	Response.Write strsql
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
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"main.asp?ip="&Request("ip")&"&agent="&Request("agent")&"&content="&vbsescape(Request("content"))&"&name="&vbsescape(Request("name"))&"&page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage

	Response.Write "</p></div>"
	objRS.Close
	Set objRS=Nothing
End Function

Function ExportSearch(name,field)
	ExportSearch=IIf(ZC_MSSQL_ENABLE,"AND ( (CHARINDEX('" & name &"',["&field&"]))<>0)","AND (InStr(1,LCase(["&field&"]),LCase('" & Name &"'),0)<>0) ")
End Function
%>