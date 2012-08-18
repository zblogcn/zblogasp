<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

Dim a

Set a = New TArticle
'a.AuthorID=1
a.id=0
a.title="留言本"
a.FType=1
a.Content="<p>这是我的留言本,欢迎给我留言.</p>"
a.Level=4
a.Alias="guestbook"
Response.write a.post
Set a=Nothing


Set a = New TArticle
'a.AuthorID=1
a.id=0
a.Title="Hello, world!"
a.FType=0
a.Content="<p>欢迎使用Z-Blog,这是程序自动生成的文章.您可以删除或是编辑它,在没有进行&quot;文件重建&quot;前,无法打开该文章页面的,这不是故障:)</p><p>系统总共生成了一个&quot;留言本&quot;页面,和一个&quot;Hello, world!&quot;文章,祝您使用愉快!</p><p>默认管理员账号和密码为:zblogger.</p>"
a.Intro=a.Content
a.Level=4
Response.write a.post
Set a=Nothing


%>
<br/><%=RunTime()%>ms<br/>
<%

Call System_Terminate()

%>