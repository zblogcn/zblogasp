<%
on error resume next
dim db,conn,myconn
db=WP_DATA_PATH
Set Conn = Server.CreateObject("ADODB.Connection")
MyConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
Conn.Open MyConn

If Err Then
    err.Clear
    Set Conn = Nothing
    'Call SetBlogHint_Custom("!! 数据库连接错误,你可以修改include.asp中的数据库路径.")
    Response.Redirect "admin_setting.asp"
    Response.End
End If
%>