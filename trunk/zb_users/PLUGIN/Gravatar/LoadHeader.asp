<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize
Dim c,m,h:Set c=CreateObject("scripting.filesystemobject")
Dim a:a=BlogPath& "zb_users\avatar\"
Dim e:e=BlogUser.ID&".png"

if b(a&e) And Request.QueryString("act")<>"refresh" Then
	Response.Redirect GetCurrentHost&"zb_users/avatar/"&e
Else
	Set m=New TConfig:m.Load "Gravatar":h=m.Read("c")
	if h="" then z
	h=Replace(h,"<#article/comment/emailmd5#>",MD5(Bloguser.Email))
	If Request.QueryString("act")="refresh" Then
		Call SetBlogHint(True,True,Empty)
		f h,a&e,"main.asp"
	Else
		f h,a&e,GetCurrentHost&"zb_users/avatar/"&e
	End If
	
End If
Function b(n):b=c.FileExists(n):End Function
sub z()
	Response.EnD
end sub
sub f(k,l,n)
	on error resume next
	dim u,v,w
	set u=server.createobject("msxml2.serverxmlhttp")
	u.open "GET",k
	u.send
	If err.number<>0 then exit sub
	v=u.ResponseBody 
	set w=server.createObject("Adodb.Stream") 
	w.Type = 1 
	w.Open 
	w.Write v 
	w.SaveToFile l,2 
	w.Close() 
	Set w=nothing 
	Set u=nothing
	response.Redirect n
end sub
set c=nothing
%>