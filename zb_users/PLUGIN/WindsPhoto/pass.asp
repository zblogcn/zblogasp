<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备   注:     WindsPhoto
'// 最后修改：   2010.6.10
'// 最后版本:    2.7.1
'///////////////////////////////////////////////////////////////////////////////
%>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%Call System_Initialize
Call WindsPhoto_Initialize
%>

<%
pa = Request.Form("pase")
If pa = "" Then
    response.Write "<script>alert('对不起,密码不可以为空!');history.back();</script>"
    response.End
End If
Dim TypeName, classid, classname, pl, hot
If IsNumeric(Request.QueryString("typeid")) = FALSE Then
    response.Write "<script>alert('对不起,参数错误,请重新再试!');history.back();</script>"
Else
    typeid = CInt(Request.QueryString("typeid"))
End If
Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select * from WindsPhoto_zhuanti where id="&typeid
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, objConn, 1, 3
If Rs("pass") = "" Or IsNull(Rs("pass")) = TRUE Then
    Response.redirect "album.asp?typeid="&typeid
    Response.End
Else
    ps = Rs("pass")
End If
If ps <> pa Then
    response.Write "<script>alert('对不起,密码错误,请重新再试!');history.back();</script>"
    response.End
Else
    Response.cookies("'"&typeid&"'") = pa
    Response.redirect "album.asp?typeid="&typeid
    Response.End
End If
%>