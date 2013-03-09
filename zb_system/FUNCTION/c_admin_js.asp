<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_admin_js.asp
'// 开始时间:    
'// 最后修改:    
'// 备    注:    后台ajax辅助
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.ContentType="application/x-javascript" %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<% 
Call LoadGlobeCache

Dim s
Dim f
f=Request.QueryString("act")
If f<>"" Then

	If f="tags" Then

		Call OpenConnect()

		Set BlogUser = New TUser
		If BlogUser.Verify()=True Then
			If CheckRights("ArticleEdt")=True Then

				'ajax tags
				Response.Write "$(""#ajaxtags"").html("""
				Dim objRS
				Set objRS=objConn.Execute("SELECT TOP 50 [tag_ID],[tag_Name] FROM [blog_Tag] ORDER BY [tag_Count] DESC,[tag_ID] ASC")
				If (Not objRS.bof) And (Not objRS.eof) Then
					Do While Not objRS.eof
						Response.Write "<a href='#'>"& TransferHTML(objRS("tag_Name"),"[html-format]") &"</a> "
						objRS.MoveNext
					Loop
				End If
				objRS.Close
				Set objRS=Nothing
				Response.Write """);$(""#ulTag"").tagTo(""#edtTag"");"

			End If
		End If

		Call CloseConnect()

	End If
	Response.End
End If
%>
