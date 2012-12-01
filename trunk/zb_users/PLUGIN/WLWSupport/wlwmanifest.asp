<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<% Response.ContentType = "text/xml" %><?xml version="1.0" encoding="utf-8" ?>
<manifest xmlns="http://schemas.microsoft.com/wlw/manifest/weblog">

  <options>
    <clientType>Metaweblog</clientType>
	<supportsKeywords>Yes</supportsKeywords>
	<supportsNewCategories>No</supportsNewCategories> 
	<supportsNewCategoriesInline>No</supportsNewCategoriesInline> 
	<supportsCommentPolicy>No</supportsCommentPolicy> 
	<supportsSlug>Yes</supportsSlug> 
	<supportsExcerpt>Yes</supportsExcerpt> 
	<supportsEmbeds>Yes</supportsEmbeds> 
	<supportsScripts>Yes</supportsScripts> 
	<supportsEmptyTitles>No</supportsEmptyTitles> 
	<requiresHtmlTitles>No</requiresHtmlTitles> 
	<supportsPostAsDraft>Yes</supportsPostAsDraft> 
	<supportsCustomDate>Yes</supportsCustomDate> 
	<supportsFileUpload>Yes</supportsFileUpload> 
	<supportsCategories>Yes</supportsCategories> 
	<supportsMultipleCategories>No</supportsMultipleCategories> 
  </options>
  
</manifest>