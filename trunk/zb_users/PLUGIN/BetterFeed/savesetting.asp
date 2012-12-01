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
If CheckPluginState("BetterFeed")=False Then Call ShowError(48)
	
	Dim strCopyright_message
	Dim blnAddreadmoreinfeed
	Dim strReadmore_message
	Dim blnAddcommentinfeed
	Dim strComment_message
	
	Dim blnCommentinfeed
	Dim intCommentinfeed_limit
	Dim strCommentinfeed_before
	Dim strCommentinfeed_layout
	Dim strCommentinfeed_after
	
	Dim blnRelatedpostinfeed
	Dim intRelatedpostinfeed_limit
	Dim strRelatedpostinfeed_before
	Dim strRelatedpostinfeed_layout
	Dim strRelatedpostinfeed_after
	Dim strRelatedpostinfeed_sub
	Dim strOtherinfeed
	


	strCopyright_message=closeHTML(Replace(Replace(Request.Form("Copyright_message"),vbCr,""),vbLf,""))
	
	blnAddreadmoreinfeed=Request.Form("addreadmoreinfeed")
	If IsEmpty(blnAddreadmoreinfeed) Then blnAddreadmoreinfeed=False

	strReadmore_message=closeHTML(Replace(Replace(Request.Form("Readmore_message"),vbCr,""),vbLf,""))
	
	blnAddcommentinfeed=Request.Form("addcommentinfeed")
	If IsEmpty(blnAddcommentinfeed) Then blnAddcommentinfeed=False

	strComment_message=closeHTML(Replace(Replace(Request.Form("Comment_message"),vbCr,""),vbLf,""))
	
	blnCommentinfeed=Request.Form("Commentinfeed")
	If IsEmpty(blnCommentinfeed) Then blnCommentinfeed=False

	intCommentinfeed_limit=Request.Form("Commentinfeed_limit")	
	strCommentinfeed_before=Replace(Replace(Request.Form("Commentinfeed_before"),vbCr,""),vbLf,"")
	strCommentinfeed_layout=closeHTML(Replace(Replace(Request.Form("Commentinfeed_layout"),vbCr,""),vbLf,""))
	strCommentinfeed_after=Replace(Replace(Request.Form("Commentinfeed_after"),vbCr,""),vbLf,"")
	
	blnRelatedpostinfeed=Request.Form("Relatedpostinfeed")
	If IsEmpty(blnRelatedpostinfeed) Then blnRelatedpostinfeed=False

	intRelatedpostinfeed_limit=Request.Form("Relatedpostinfeed_limit")
	strRelatedpostinfeed_before=Replace(Replace(Request.Form("Relatedpostinfeed_before"),vbCr,""),vbLf,"")
	strRelatedpostinfeed_layout=Replace(Replace(Request.Form("Relatedpostinfeed_layout"),vbCr,""),vbLf,"")
	strRelatedpostinfeed_after=Replace(Replace(Request.Form("Relatedpostinfeed_after"),vbCr,""),vbLf,"")
	strRelatedpostinfeed_sub=Replace(Replace(Request.Form("Relatedpostinfeed_sub"),vbCr,""),vbLf,"")
	
	strOtherinfeed=Replace(Replace(Request.Form("Otherinfeed"),vbCr,""),vbLf,"")
	
	Dim c
	Set c = New TConfig
		c.Load("BetterFeed")
		c.Write "BetterFeed_Copyright_message",strCopyright_message
		c.Write "BetterFeed_Addreadmoreinfeed",blnAddreadmoreinfeed
		c.Write "BetterFeed_Readmore_message",strReadmore_message
		c.Write "BetterFeed_Addcommentinfeed",blnAddcommentinfeed
		c.Write "BetterFeed_Comment_message",strComment_message
		c.Write "BetterFeed_Commentinfeed",blnCommentinfeed
		c.Write "BetterFeed_Commentinfeed_limit",intCommentinfeed_limit
		c.Write "BetterFeed_Commentinfeed_before",strCommentinfeed_before
		c.Write "BetterFeed_Commentinfeed_layout",strCommentinfeed_layout
		c.Write "BetterFeed_Commentinfeed_after",strCommentinfeed_after
		c.Write "BetterFeed_Relatedpostinfeed",blnRelatedpostinfeed
		c.Write "BetterFeed_Relatedpostinfeed_limit",intRelatedpostinfeed_limit
		c.Write "BetterFeed_Relatedpostinfeed_before",strRelatedpostinfeed_before
		c.Write "BetterFeed_Relatedpostinfeed_layout",strRelatedpostinfeed_layout
		c.Write "BetterFeed_Relatedpostinfeed_after",strRelatedpostinfeed_after
		c.Write "BetterFeed_Relatedpostinfeed_sub",strRelatedpostinfeed_sub
		c.Write "BetterFeed_Otherinfeed",strOtherinfeed
		c.Save
	Set c=Nothing

If Err.Number<>0 then
  Call ShowError(0)
End If

Call SetBlogHint(True,True,Empty)

Response.Redirect "main.asp?s"
%>
