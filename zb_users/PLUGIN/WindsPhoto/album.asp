<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 spirit 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备    注:    WindsPhoto
'// 最后修改：   2011.8.22
'// 最后版本:    2.7.3
'///////////////////////////////////////////////////////////////////////////////
%>
<%' Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp" -->
<%Call System_Initialize
Call WindsPhoto_Initialize%><%

Dim TypeName, hot, data, data1, js, p
If IsNumeric(Request.QueryString("typeid")) = FALSE Then
    response.Write "<script>alert('对不起,参数错误!');history.back();</script>"
Else
    typeid = CInt(Request.QueryString("typeid"))
End If
If Request.QueryString("mo") <>"" And IsNumeric(Request.QueryString("mo")) = FALSE Then
    response.Write "<script>alert('对不起,参数错误!');history.back();</script>"
Else
    If Request.QueryString("mo") <>"" Then
        mo = CInt(Request.QueryString("mo"))
    End If
End If

Set temprs = objConn.Execute("select name,hot,data,time1,js,pass,[view] FROM WindsPhoto_zhuanti where id="&typeid)
If temprs.EOF Or temprs.bof Then
    response.Write "<script>alert('对不起,该相册不存在');history.back();</script>"
    response.End
End If

If temprs(6) = 1 And mo = 0 Then
    Response.Redirect "album.asp?typeid="&typeid&"&mo=1"
End If
TypeName = temprs(0)
hot = temprs(1)
data = temprs(2)
data1 = temprs(3)
js = temprs(4)
p = temprs(5)
Set temprs = Nothing

Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select * FROM WindsPhoto_zhuanti where id="&typeid
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, objConn, 1, 3
ps = Rs("pass")
If Rs("hot") = "" Or IsNull(Rs("hot")) = TRUE Then
    Rs("hot") = "1"
Else
    Rs("hot") = Rs("hot") + 1
End If
rs.update
Rs.Close
Set Rs = Nothing
%>

<%
Dim objArticle
Set objArticle = New TArticle
'objArticle.
objArticle.FType=ZC_POST_TYPE_PAGE
If GetTemplate("TEMPLATE_WP_ALBUM")<>empty Then
    objArticle.template = "WP_ALBUM"
End If


objArticle.Title = WP_ALBUM_NAME &"-"& TypeName
objArticle.Content=GetPhoto()
If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then 
	objArticle.Build

    Dim Html
    Html = objArticle.html
    Dim AddedHtml
	AddedHtml="<link rel=""alternate"" type=""application/rss+xml"" href="""& WP_SUB_DOMAIN &"rss.asp?id="&typeid&""" title=""订阅我的相册"" />" & vbCrLf
	AddedHtml = AddedHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"images/windsphoto.css"" type=""text/css"" media=""screen"" />" & VBCRLF
    AddedHtml = AddedHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/windsphoto.js""></script>" & VBCRLF & "</head>"

    If WP_SCRIPT_TYPE = "1" Then
        Html = Replace(Html, "</head>", HighSlide_Code & AddedHtml)
    ElseIf WP_SCRIPT_TYPE = "2" Then
        Html = Replace(Html, "</head>", GreyBox_Code & AddedHtml)
    ElseIf WP_SCRIPT_TYPE = "3" Then
        Html = Replace(Html, "</head>", Lightbox_Code & AddedHtml)
    ElseIf WP_SCRIPT_TYPE = "4" Then
        Html = Replace(Html, "</head>", Thickbox_Code & AddedHtml)
    Else
        Html = Replace(Html, "</head>", AddedHtml)
    End If

    Html = Replace(Html, ">Powered By", ">Powered By <a href='http://photo.wilf.cn/' target='_blank' title='WindsPhoto官方网站'>WindsPhoto</a> &")
    Call ClearGlobeCache
    Call LoadGlobeCache
    Response.Write Html
End If
Set objArticle = Nothing
%>

<%
Function LightBox_Code
    Dim innerHtml
    innerHtml = ""
    innerHtml = innerHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"script/LightBox/lightbox.css"" type=""text/css"" media=""screen"" />" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"">" & VBCRLF
    innerHtml = innerHtml & "var lightBoxM = " & ZC_IMAGE_WIDTH & ";" & VBCRLF
    innerHtml = innerHtml & "var lightBoxL = """& WP_SUB_DOMAIN &"script/LightBox/lightbox-ico-loading.gif"";" & VBCRLF
    innerHtml = innerHtml & "var lightBoxP = """& WP_SUB_DOMAIN &"script/LightBox/lightbox-btn-prev.gif"";" & VBCRLF
    innerHtml = innerHtml & "var lightBoxN = """& WP_SUB_DOMAIN &"script/LightBox/lightbox-btn-next.gif"";" & VBCRLF
    innerHtml = innerHtml & "var lightBoxC = """& WP_SUB_DOMAIN &"script/LightBox/lightbox-btn-close.gif"";" & VBCRLF
    innerHtml = innerHtml & "var lightBoxB = """& WP_SUB_DOMAIN &"script/LightBox/lightbox-blank.gif"";" & VBCRLF
    innerHtml = innerHtml & "</script>" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/LightBox/lightbox.pack.js""></script>" & VBCRLF
    LightBox_Code = innerHtml
End Function

Function GreyBox_Code
    Dim innerHtml
    innerHtml = ""
    innerHtml = innerHtml & "<script type=""text/javascript"">" & VBCRLF
    innerHtml = innerHtml & "var GB_ROOT_DIR =  """& WP_SUB_DOMAIN &"script/greybox/"";" & VBCRLF
    innerHtml = innerHtml & "</script>" & VBCRLF
    innerHtml = innerHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"script/greybox/gb_styles.css"" type=""text/css"" media=""screen"" />" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/greybox/AJS.js""></script>" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/greybox/AJS_fx.js""></script>" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/greybox/gb_scripts.js""></script>" & VBCRLF
    GreyBox_Code = innerHtml
End Function

Function ThickBox_Code
    Dim innerHtml
    innerHtml = ""
    innerHtml = innerHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"script/thickbox/thickbox.css"" type=""text/css"" media=""screen"" />" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"">" & VBCRLF
    innerHtml = innerHtml & "var tb_pathToImage = """& WP_SUB_DOMAIN &"script/thickbox/loadingAnimation.gif"";" & VBCRLF
    innerHtml = innerHtml & "</script>" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/thickbox/thickbox.pack.js""></script>" & VBCRLF
    ThickBox_Code = innerHtml
End Function

Function HighSlide_Code
    Dim innerHtml
    innerHtml = ""
    innerHtml = innerHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"script/highslide/highslide.css"" type=""text/css"" media=""screen"" />" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/highslide/highslide.packed.js""></script>" & VBCRLF
    innerHtml = innerHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/highslide/highslide2.js""></script>" & VBCRLF
    innerHtml = innerHtml & "<div id='controlbar' class='highslide-overlay controlbar'><a href='#' class='previous' onclick='return hs.previous(this)' title='上一张'></a><a href='#' class='next' onclick='return hs.next(this)' title='下一张'></a><a href='#' class='highslide-move' onclick='return false' title='拖动'></a><a href='#' class='close' onclick='return hs.close(this)' title='关闭'></a></div>" & VBCRLF
    HighSlide_Code = innerHtml
End Function

Function Gallery_Style
    Dim innerHtml
    innerHtml = ""
    If WP_SCRIPT_TYPE = "1" Then
        innerHtml = innerHtml & "rel='class='highslide' onclick='return hs.expand(this)'"
    ElseIf WP_SCRIPT_TYPE = "2" Then
        innerHtml = innerHtml & "rel='gb_imageset[nice_pics]'"
    ElseIf WP_SCRIPT_TYPE = "4" Then
        innerHtml = innerHtml & "class='thickbox' rel='gallery-plants'"
    Else
    End If
    Gallery_Style = innerHtml
End Function

Function GetType()
    sqlo2 = "select id FROM WindsPhoto_desktop where zhuanti="&typeid&" order by id asc"
    Set rso2 = Server.CreateObject("ADODB.RecordSet")
    rso2.Open sqlo2, objConn, 1, 1
    If rso2.EOF Or rso2.bof Then
        sm = 0
    Else
        sm = rso2.RecordCount
    End If
    rso2.Close
    Set rso2 = Nothing
    GetType = "<div>"&TypeName&"</div>"
End Function

Function GetPhoto()

    '判断是否需要密码
    pss = Request.cookies("'"&typeid&"'")
    If Len(ps)>0 Then
        If pss<>ps Then
            GetPhoto = GetPhoto&"<div><span style='color:#808000;font-size:14px;'>对不起，该相册为加密相册，如果你有密码的话，请提供查看密码：</a></span><form name='form' method='post' action='pass.asp?typeid="&typeid&"'><input type='password' name='pase'><input type='submit' name='Submit' value='确定'></form></div>"
            Exit Function '中止Function
        End If
    End If
    '取得相册信息
    sqlo = "select id FROM WindsPhoto_desktop where zhuanti="&typeid&" order by id asc"
    Set rso = Server.CreateObject("ADODB.RecordSet")
    rso.Open sqlo, objConn, 1, 1
    If rso.EOF Or rso.bof Then
        sm = 0
    Else
        sm = rso.RecordCount
    End If
    rso.Close
    Set rso = Nothing

    GetPhoto = GetPhoto&"<p style='margin:0;line-height:1px'>&nbsp;</p>" 'firefox果然很变态
    GetPhoto = GetPhoto& VBCRLF&"<table border='0' cellpadding='0' cellspacing='0' width='100%' height='60'>"& VBCRLF
    GetPhoto = GetPhoto&"<tr><td>本相册共有"&sm&"张照片</td><td colspan='2' align='right'><a href='album.asp?typeid="&typeid&"&mo=2'>以缩略图形式查看</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='album.asp?typeid="&typeid&"&mo=1'>以列表形式查看</a></td></tr>"& VBCRLF
    GetPhoto = GetPhoto&"<tr><td>发布日期："&data1&"(浏览:"&hot&")</td>"
    GetPhoto = GetPhoto&"<td>拍摄日期："&data&"</td>"
    GetPhoto = GetPhoto&"<td align='right'>"
    If p<>"" Then
        GetPhoto = GetPhoto&"<span style=""color:red;"">需权限浏览</span>"
    Else
        GetPhoto = GetPhoto&"公开相册"
    End If
    GetPhoto = GetPhoto&"</td></tr></table>"& VBCRLF
    '介绍
    If js<>"" Then
        GetPhoto = GetPhoto&"<div style='margin:10px 0 5px 0;'>"&js&"</div>"& VBCRLF
    End If
    '相片
    Dim ipagecount
    Dim ipagecurrent
    Dim irecordsshown
    If request.querystring("page") = "" Then
        ipagecurrent = 1
    Else
        ipagecurrent = CInt(request.querystring("page"))
    End If
    If mo<>1 Then
        GetPhoto = GetPhoto&"<table width='100%' border='0' cellspacing='0' cellpadding='5'>"& VBCRLF
        If WP_ORDER_BY = "0" Then
            sql = "SELECT surl,url,name,jj,id FROM WindsPhoto_desktop where zhuanti="&typeid&" ORDER BY id asc"
        Else
            sql = "SELECT surl,url,name,jj,id FROM WindsPhoto_desktop where zhuanti="&typeid&" ORDER BY id desc"
        End If
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, objConn, 1, 1
        rs.pagesize = WP_SMALL_PAGERCOUNT
        ipagecount = rs.pagecount
        If ipagecurrent > ipagecount Then ipagecurrent = ipagecount End If
        If ipagecurrent < 1 Then ipagecurrent = 1 End If
        If ipagecount = 0 Then
            GetPhoto = GetPhoto&"<tr><td align='center'><img src='images/nopic.jpg' /></tr></td>"& VBCRLF
        Else
            rs.absolutepage = ipagecurrent
            irecordsshown = 0
            Do While irecordsshown<WP_SMALL_PAGERCOUNT And Not rs.EOF
                GetPhoto = GetPhoto&"<tr align='center'>"& VBCRLF
                If Not rs.EOF Then
                    surl = rs("surl")
                    url = rs("url")
                    If InStr(surl, "photo.163.com") Or InStr(surl, "126.net") Or InStr(surl, "photo.sina.com") Or InStr(surl, "photos.baidu.com") Then surl = "stealink.asp?" & surl End If
                    If InStr(url, "photo.163.com") Or InStr(url, "126.net") Or InStr(url, "photo.sina.com") Or InStr(url, "photos.baidu.com") Then url = "stealink.asp?" & url End If
                    GetPhoto = GetPhoto&"<td width='33%'><a href='"&url&"' "& Gallery_Style &" title='"&rs("jj")&"'><img class='wp_small' src='"&surl&"' title='"&rs("jj")&"' alt='"&rs("name")&"' onload='WindsPhotoResizeImage(this,"&WP_SMALL_WIDTH&","&WP_SMALL_HEIGHT&")' /></a><br />"&rs("name")&"</td>"& VBCRLF
                    irecordsshown = irecordsshown + 1
                    rs.movenext
                End If

                If Not rs.EOF Then
                    surl = rs("surl")
                    url = rs("url")
                    If InStr(surl, "photo.163.com") Or InStr(surl, "126.net") Or InStr(surl, "photo.sina.com") Or InStr(surl, "photos.baidu.com") Then surl = "stealink.asp?" & surl End If
                    If InStr(url, "photo.163.com") Or InStr(url, "126.net") Or InStr(url, "photo.sina.com") Or InStr(url, "photos.baidu.com") Then url = "stealink.asp?" & url End If
                    GetPhoto = GetPhoto&"<td width='33%'><a href='"&url&"' "& Gallery_Style &" title='"&rs("jj")&"'><img class='wp_small' src='"&surl&"' title='"&rs("jj")&"' alt='"&rs("name")&"' onload='WindsPhotoResizeImage(this,"&WP_SMALL_WIDTH&","&WP_SMALL_HEIGHT&")' /></a><br />"&rs("name")&"</td>"& VBCRLF
                    irecordsshown = irecordsshown + 1
                    rs.movenext
                End If

                If Not rs.EOF Then
                    surl = rs("surl")
                    url = rs("url")
                    If InStr(surl, "photo.163.com") Or InStr(surl, "126.net") Or InStr(surl, "photo.sina.com") Or InStr(surl, "photos.baidu.com") Then surl = "stealink.asp?" & surl End If
                    If InStr(url, "photo.163.com") Or InStr(url, "126.net") Or InStr(url, "photo.sina.com") Or InStr(url, "photos.baidu.com") Then url = "stealink.asp?" & url End If
                    GetPhoto = GetPhoto&"<td width='33%'><a href='"&url&"' "& Gallery_Style &" title='"&rs("jj")&"'><img class='wp_small' src='"&surl&"' title='"&rs("jj")&"' alt='"&rs("name")&"' onload='WindsPhotoResizeImage(this,"&WP_SMALL_WIDTH&","&WP_SMALL_HEIGHT&")' /></a><br />"&rs("name")&"</td>"& VBCRLF
                    irecordsshown = irecordsshown + 1
                    rs.movenext
                End If
                GetPhoto = GetPhoto&"</tr>"
            Loop
        End If
        GetPhoto = GetPhoto&"</table></tr></table>"& VBCRLF
        '分页
        If ipagecount >1 Then
            GetPhoto = GetPhoto&"<div class=""post pagebar"">"
            'GetPhoto = GetPhoto&"<span class=""other-page"">"&ipagecount&"页中的第"&ipagecurrent&"页</span>"
            GetPhoto = GetPhoto&"<a title='首页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page=1'>1</a>"
            If ipagecurrent = 1 Then
                GetPhoto = GetPhoto&"<span class=""other-page"">«</span>"
            Else
                GetPhoto = GetPhoto&"<a title='上一页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&ipagecurrent -1&"'>«</a>"
            End If
            
            If ipagecount>ZC_PAGEBAR_COUNT Then
                a=ipagecurrent-Cint((ZC_PAGEBAR_COUNT-1)/2)
                b=ipagecurrent+ZC_PAGEBAR_COUNT-Cint((ZC_PAGEBAR_COUNT-1)/2)-1
                If a<=1 Then
                    a=1:b=ZC_PAGEBAR_COUNT
                End If
                If b>=ipagecount Then
                    b=ipagecount:a=ipagecount-ZC_PAGEBAR_COUNT+1
                End If
            Else
                a=1:b=ipagecount
            End If

			For i = a to b
                'ipagenow = ipagenow + 1
                If ipagecurrent = i Then
                    GetPhoto=GetPhoto&"<span class=""now-page"">"&i&"</span>"
                Else
                    GetPhoto=GetPhoto&"<a href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&i&"'>"&i&"</a>"
                End If
            Next

            If ipagecount>ipagecurrent Then
                GetPhoto = GetPhoto&"<a title='下一页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&ipagecurrent + 1&"'>»</a>"
            Else
                GetPhoto = GetPhoto&"<span class=""other-page"">»</span>"
            End If

            GetPhoto = GetPhoto&"<a title='尾页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&ipagecount&"'>"&ipagecount&"</a></div>"
        End If
        rs.Close
        Set rs = Nothing

    Else

        If WP_ORDER_BY = "0" Then
            sql = "SELECT url,name,id,jj FROM WindsPhoto_desktop where zhuanti="&typeid&" ORDER BY id asc"
        Else
            sql = "SELECT url,name,id,jj FROM WindsPhoto_desktop where zhuanti="&typeid&" ORDER BY id desc"
        End If
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, objConn, 1, 1
        rs.pagesize = WP_LIST_PAGERCOUNT
        ipagecount = rs.pagecount
        If ipagecurrent > ipagecount Then ipagecurrent = ipagecount End If
        If ipagecurrent < 1 Then ipagecurrent = 1 End If
        If ipagecount = 0 Then
            GetPhoto = GetPhoto&"<p align='center'><img src='images/nopic.jpg'></p>"& VBCRLF
        Else
            rs.absolutepage = ipagecurrent
            irecordsshown = 0
            Do While irecordsshown<WP_LIST_PAGERCOUNT And Not rs.EOF
                url = rs("url")
                If InStr(url, "photo.163.com") Or InStr(url, "126.net") Or InStr(url, "photo.sina.com") Or InStr(url, "photos.baidu.com") Then url = "stealink.asp?" & url End If            
                GetPhoto = GetPhoto&"<p align='center'><a href='"&url&"' " & Gallery_Style &" title='"&rs("jj")&"'><img class='wp_list' alt='"&rs("name")&"' src='"&url&"' onload='WindsPhotoResizeImage(this,"&WP_LIST_WIDTH&","&WP_LIST_HEIGHT&")' /></a><br />"&rs("name")&"</p>"& VBCRLF
                irecordsshown = irecordsshown + 1
                rs.movenext
            Loop
        End If

        '分页
        If ipagecount >1 Then
            GetPhoto = GetPhoto&"<div class=""post pagebar"">"
            'GetPhoto = GetPhoto&"<span class=""other-page"">"&ipagecount&"页中的第"&ipagecurrent&"页</span>"
            GetPhoto = GetPhoto&"<a title='首页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page=1'>1</a>"
            If ipagecurrent = 1 Then
                GetPhoto = GetPhoto&"<span class=""other-page"">«</span>"
            Else
                GetPhoto = GetPhoto&"<a title='上一页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&ipagecurrent -1&"'>«</a>"
            End If

            If ipagecount>ZC_PAGEBAR_COUNT Then
                a=ipagecurrent-Cint((ZC_PAGEBAR_COUNT-1)/2)
                b=ipagecurrent+ZC_PAGEBAR_COUNT-Cint((ZC_PAGEBAR_COUNT-1)/2)-1
                If a<=1 Then
                    a=1:b=ZC_PAGEBAR_COUNT
                End If
                If b>=ipagecount Then
                    b=ipagecount:a=ipagecount-ZC_PAGEBAR_COUNT+1
                End If
            Else
                a=1:b=ipagecount
            End If

			For i = a to b
                'ipagenow = ipagenow + 1
                If ipagecurrent = i Then
                    GetPhoto=GetPhoto&"<span class=""now-page"">"&i&"</span>"
                Else
                    GetPhoto=GetPhoto&"<a href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&i&"'>"&i&"</a>"
                End If
            Next

            If ipagecount>ipagecurrent Then
                GetPhoto = GetPhoto&"<a title='下一页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&ipagecurrent + 1&"'>»</a>"
            Else
                GetPhoto = GetPhoto&"<span class=""other-page"">»</span>"
            End If

            GetPhoto = GetPhoto&"<a title='尾页' href='album.asp?typeid="&typeid&"&mo="&request.querystring("mo")&"&page="&ipagecount&"'>"&ipagecount&"</a></div>"
        End If
        rs.Close
        Set rs = Nothing

    End If
End Function
%>

<%

If Err.Number<>0 then
	Call ShowError(0)
End If
%>