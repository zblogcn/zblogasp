<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% Response.Charset="UTF-8" %>
<% Response.Expires=0 %>
<% Response.ContentType = "text/css" %>
<!-- #include file="../../../c_option.asp" -->
<%
Response.Write("@import url("""& ZC_BLOG_HOST & "zb_users/theme" & "/" & ZC_BLOG_THEME & "/style/" & ZC_BLOG_CSS & ".css" & """);") 
%>