<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% Response.Charset="UTF-8" %>
<% Response.Expires=0 %>
<% Response.ContentType = "text/css" %>
<!-- #include file="../../../c_option.asp" -->
<!-- #include file="../../../../zb_system/function/c_function.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\..\..\plugin\p_config.asp" -->


<%
Response.Write("@import url("""& GetCurrentHost & "zb_users/theme" & "/" & ZC_BLOG_THEME & "/style/" & ZC_BLOG_CSS & ".css" & """);") 


	Call OpenConnect()

	Call GetConfigs()

	Dim c
	Set c = New TConfig
	c.Load("WhitePage")

	If c.Exists("custom_bgcolor")=True Then Response.Write "body{background-color:" & c.Read("custom_bgcolor") & ";}"
	If c.Exists("custom_headtitle")=True Then Response.Write "#BlogTitle,#BlogSubTitle{text-align:" & c.Read("custom_headtitle") & ";}"
	If c.Exists("custom_pagewidth")=True Then
		if c.Read("custom_pagewidth")=1000 then
			Response.Write "#divAll{width:1000px;}#divMiddle{width:940px;padding:0 30px;}#divSidebar{width:240px;padding:0 0 0 20px;}#divMain{width:670px;padding:0 0 20px 0;}#divTop{padding-top:30px;}body{font-size:15px;}"
		end if 
	End If
	If c.Exists("custom_pagetype")=True Then	
		if c.Read("custom_pagetype")=1 then
			if c.Read("custom_pagewidth")=1000 then
				Response.Write "#divAll{background:url('../style/default/bg1000-1.png') no-repeat 50% top;}#divPage{background:url('../style/default/bg1000-2.png') no-repeat 50% bottom;}#divMiddle{background:url('../style/default/bg1000-3.png') repeat-y 50% 50%;}"
			end if
		end if 
		if c.Read("custom_pagetype")=2 then
			Response.Write "#divAll{box-shadow: 0 0 5px #666;background-color:white;border-radius: 0px;}"
			Response.Write "#divAll{background:white;}#divPage{background:none;}#divMiddle{background:none;}"
		end if 
		if c.Read("custom_pagetype")=3 then
			Response.Write "#divAll{box-shadow: 0 0 5px #666;background-color:white;border-radius: 5px;}"
			Response.Write "#divAll{background:white;}#divPage{background:none;}#divMiddle{background:none;}"
		end if
		if c.Read("custom_pagetype")=4 then
			Response.Write "#divAll{box-shadow:none;background-color:white;border-radius: 0;}"
			Response.Write "#divAll{background:white;}#divPage{background:none;}#divMiddle{background:none;}"
			Response.Write "#divTop{padding-top:30px;}"
		end if
	End If	
	
	c.Save
	Set c =Nothing




	Call CloseConnect()

%>