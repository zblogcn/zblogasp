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
<rsd version="1.0">
    <service>
        <engineName>Z-Blog</engineName>
        <engineLink>http://www.rainbowsoft.org/</engineLink>
        <homePageLink><%=ZC_BLOG_HOST%></homePageLink>
        <apis>
            <api name="MetaWeblog" preferred="true" apiLink="<%=ZC_BLOG_HOST%>zb_system/xml-rpc/index.asp" blogID="1"/>
            <api name="Blogger" preferred="false" apiLink="<%=ZC_BLOG_HOST%>zb_system/xml-rpc/index.asp" blogID="1"/>
        </apis>
    </service>
</rsd>