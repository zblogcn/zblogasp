<%
Call RegisterPlugin("LockZBlog","ActivePlugin_LockZBlog")

Function ActivePlugin_LockZBlog()
	Call Add_Action_Plugin("Action_Plugin_System_Initialize","LockZBlog")
	Call Add_Action_Plugin("Action_Plugin_Default_Begin","LockZBlog")
End Function

Function LockZBlog()
	If Request.QueryString("act")="passwordvaild" Then
		Response.Cookies("LockZBlog")=Request("password")
		Response.Cookies("LockZBlog").Path=CookiesPath
		'Response.Cookies("LockZBlog").Path=CookiesPath
'		Response.Cookies("LockZBlog").
	End If
	If Request.Cookies("LockZBlog")<>"password" Then
%>
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>请输入密码</title>
    </head>
    
    <body>
    <form action="<%=BlogHost%>default.asp?act=passwordvaild" method="post">
    请输入密码<input type="password" value="" name="password" /><input  type="submit" value="提交"/>
    </form>
    </body>
    </html>

<%
	Response.End
	End If
End Function

%>