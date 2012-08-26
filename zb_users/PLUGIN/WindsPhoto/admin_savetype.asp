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
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="data/conn.asp" -->
<!-- #include file="function.asp" -->

<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "管 理 相 册"

%>

<%
If Request.Form("zt") = "" Then
    zt = Request.QueryString("zt")
Else
    zt = Request.Form("zt")
End If

If Trim(Request.Form("name")) = "" Then
    Call SetBlogHint_Custom("!! 相册名不能为空")
    Response.Redirect"admin_main.asp"
Else
    Set rs = Server.CreateObject("ADODB.Recordset")
    If zt = "editzhuanti" Then
        sql = "SELECT * FROM zhuanti where id="&Request.Form("typeid")
        rs.Open sql, Conn, 1, 3
        rs("name") = Request.Form("name")
        rs("time1") = Request.Form("fabu")
        rs("data") = Request.Form("riqi")
        rs("js") = Request.Form("js")
        rs("pass") = Request.Form("pass")
        rs("view") = Request.Form("view")
        
        If Request.Form("ordered") = "" then
          sql2 = "select * from zhuanti order by ordered,id asc"
          Set rs2 = Server.CreateObject("ADODB.Recordset")
          rs2.Open sql2, conn, 1, 1
          ordered = rs2.RecordCount + 1
        Else
          ordered = Request.Form("ordered")
        end If
        
        rs("ordered") = ordered
        
        rs.update
        rs.Close
        Set rs = Nothing      
                
        Call SaveSortList()
        
        Call SetBlogHint_Custom("√ 编辑相册分类成功.")

        Response.Redirect"admin_main.asp"
    End If

    If zt = "addzhuanti" Then
        sql = "SELECT * FROM zhuanti where (id is null)"
        rs.Open sql, Conn, 1, 3
        rs.addnew
        rs("name") = Request.Form("name")
        rs("time1") = Request.Form("fabu")
        rs("data") = Request.Form("riqi")
        rs("js") = Request.Form("js")
        rs("pass") = Request.Form("pass")
        rs("view") = Request.Form("view")
        
        If Request.Form("ordered") = "" then
          sql2 = "select * from zhuanti order by ordered,id asc"
          Set rs2 = Server.CreateObject("ADODB.Recordset")
          rs2.Open sql2, conn, 1, 1
          ordered = rs2.RecordCount + 1
        Else
          ordered = Request.Form("ordered")
        end If
        
        rs("ordered") = ordered
        
        rs.update
        rs.Close
        Set rs = Nothing
        
        Call SaveSortList()        
       
        Call SetBlogHint_Custom("√ 添加相册分类成功.")
        
        Response.Redirect "admin_main.asp"
    End If

End If

If zt = "delzhuanti" Then
    conn.Execute "delete from zhuanti where id="&Request.Form("typeid")
    conn.Execute "delete from desktop where zhuanti="&Request.Form("typeid")
    
    conn.Close
    Set conn = Nothing
    
    Call SaveSortList()
    
    Call SetBlogHint_Custom("√ 删除相册分类成功.")
    
    Response.Redirect "admin_main.asp"
End If
%>
<%
Call System_Terminate()

If Err.Number<>0 Then
    Call ShowError(0)
End If
%>