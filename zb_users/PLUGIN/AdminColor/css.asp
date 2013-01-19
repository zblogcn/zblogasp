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



Response.Expires=0
Response.ContentType = "text/css"

Dim c
c=LoadFromFile(BlogPath & "zb_system\CSS\admin2.css","utf-8")

c=Replace(c,"url(../image/","url(../../../zb_system/image/")

Dim c1,c2,c3,c4,c5
c1="#1d4c7d"
c2="#3a6ea5"
c3="#b0cdee"
c4="#3399cc"
c5="#d60000"

c=Replace(c,c1,BlodColor(1))
c=Replace(c,c2,NormalColor(1))
c=Replace(c,c3,LightColor(1))
c=Replace(c,c4,HighColor(1))
c=Replace(c,c5,AntiColor(1))

Response.Write(c) 

Call System_Terminate()
%>