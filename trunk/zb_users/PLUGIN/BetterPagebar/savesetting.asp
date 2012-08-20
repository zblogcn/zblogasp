<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%

Call System_Initialize()

Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
If CheckPluginState("BetterPagebar")=False Then Call ShowError(48)

	Dim blnPagebar_AlwaysShow
	
	Dim strPagebar_FristPage
	Dim strPagebar_LastPage
	Dim strPagebar_PrvePage
	Dim strPagebar_NextPage
	Dim strPagebar_FristPage_Tip
	Dim strPagebar_LastPage_Tip
	Dim strPagebar_PrvePage_Tip
	Dim strPagebar_NextPage_Tip	
	Dim strPagebar_Extend
	
	blnPagebar_AlwaysShow=Request.Form("AlwaysShow")
	If IsEmpty(blnPagebar_AlwaysShow) Then blnPagebar_AlwaysShow=False
	
	strPagebar_FristPage=Trim(Request.Form("FristPage"))
	strPagebar_LastPage=Trim(Request.Form("LastPage"))
	strPagebar_PrvePage=Trim(Request.Form("PrvePage"))
	strPagebar_NextPage=Trim(Request.Form("NextPage"))
	strPagebar_FristPage_Tip=Trim(Request.Form("FristPage_Tip"))
	strPagebar_LastPage_Tip=Trim(Request.Form("LastPage_Tip"))
	strPagebar_PrvePage_Tip=Trim(Request.Form("PrvePage_Tip"))
	strPagebar_NextPage_Tip=Trim(Request.Form("NextPage_Tip"))	
	strPagebar_Extend=Trim(Request.Form("Extend"))
	Dim c
	Set c = New TConfig
		c.Load("BetterPagebar")
		c.Write "BetterPagebar_AlwaysShow",blnPagebar_AlwaysShow
		c.Write "BetterPagebar_FristPage",strPagebar_FristPage
		c.Write "BetterPagebar_LastPage",strPagebar_LastPage
		c.Write "BetterPagebar_PrvePage",strPagebar_PrvePage
		c.Write "BetterPagebar_NextPage",strPagebar_NextPage
		c.Write "BetterPagebar_FristPage_Tip",strPagebar_FristPage_Tip
		c.Write "BetterPagebar_LastPage_Tip",strPagebar_LastPage_Tip
		c.Write "BetterPagebar_PrvePage_Tip",strPagebar_PrvePage_Tip
		c.Write "BetterPagebar_NextPage_Tip",strPagebar_NextPage_Tip
		c.Write "strPagebar_Extend",strPagebar_Extend
		c.Save
	Set c=Nothing	


If Err.Number<>0 then
  Call ShowError(0)
End If

Call SetBlogHint(True,True,Empty)

Response.Redirect "main.asp?s"

%>
