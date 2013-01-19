<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%

Call System_Initialize()


Dim ac_config,id
Set ac_config=New TConfig
ac_config.Load "AdminColor"
id=ac_config.read("ColorID")

If Request.QueryString("color")<>"" Then

	ac_config.Write "ColorID",Request.QueryString("color")
	ac_config.Save
	Response.Redirect bloghost & "zb_system/cmd.asp?act=admin"

End If

If id="" Then id=0
If id="random" Then
	Randomize
	id=Int((Ubound(BlodColor) - LBound(BlodColor) + 1) * Rnd + LBound(BlodColor))
End If

Response.Expires=0
Response.ContentType = "text/css"

Dim c
c=LoadFromFile(BlogPath & "zb_system\CSS\admin2.css","utf-8")

c=Replace(c,"url(../","url(../../../zb_system/")



Dim c1,c2,c3,c4,c5
c1="#1d4c7d"
c2="#3a6ea5"
c3="#b0cdee"
c4="#3399cc"
c5="#d60000"

c=Replace(c,c1,BlodColor(id))
c=Replace(c,c2,NormalColor(id))
c=Replace(c,c3,LightColor(id))
c=Replace(c,c4,HighColor(id))
c=Replace(c,c5,AntiColor(id))
c=Replace(c,"../IMAGE","../../../zb_system/IMAGE")

c=c & "#admin_color{line-height: 2.5em;font-size: 0.5em;letter-spacing: -0.1em;}"
Response.Write(c) 

Call System_Terminate()
%>