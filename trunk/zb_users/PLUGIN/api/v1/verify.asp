<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../c_option.asp" -->
<!-- #include file="../../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../p_config.asp" -->
<!-- #include file="../JSON.asp" -->
<%
'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 验证key和连接
' 参数: ?keyid=0839609affdbb82c9884fc05a2dfcd18&keysecretmd5=a9b0772d4900afd67178b14579b008ca&post=111
'*********************************************************
	Call System_Initialize()
	Response.ContentType="application/json"

	Dim objConfig,data_export,errcode,body_array(4),request_array(3),msg,ret:errcode="0":msg="true":ret=0
	Set objConfig=New TConfig:Set data_export = jsObject()
	objConfig.Load("api")

	Verify()	'参数keyid，keysecret md5加密后的keysecretmd5，非空任意值post
	
	data_export("errcode") = errcode
	data_export("msg") = msg
	If errcode<>0 Then ret=1 End If:data_export("ret") = ret
	If ret=0 Then data_export("body") = body_array End If
	data_export("timestamp") = DateDiff("s", "01/01/1970 00:00:00", Now())
	data_export.Flush
	
	Function Verify()
		request_array(0) = Request("keyid")
		request_array(1) = Request("keysecretmd5")
		request_array(2) = Request("post")

		If request_array(0)<>objConfig.Read("id") Then
			errcode="001":msg="keyid is wrong."
		ElseIf request_array(1)<>md5(objConfig.Read("secret")) Then
			errcode="002":msg="keysecret is wrong."
		ElseIf request_array(2)<>"" Then
			body_array("0")=Array("blog_title",ZC_BLOG_TITLE)
			body_array("1")=Array("blog_subtitle",ZC_BLOG_SUBTITLE)
			body_array("2")=Array("blog_url",ZC_BLOG_HOST)
			body_array("3")=Array("blog_language",ZC_BLOG_LANGUAGE)
			body_array("4")=Array("blog_version",ZC_BLOG_VERSION)
		End If
	End Function
%>