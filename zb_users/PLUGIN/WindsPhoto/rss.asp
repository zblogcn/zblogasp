<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 spirit 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备   注:     WindsPhoto RSS
'// 最后修改：   2010.6.10
'// 最后版本:    2.7.1
'///////////////////////////////////////////////////////////////////////////////
%>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!--#include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" --><%Call System_Initialize%>

<%
Dim sSQL, rs, sCrLf, sXmlClear, sRssHead, sRssEnd
sCrLf = Chr(13) & Chr(10)

sXmlClear = "<?xml version='1.0' encoding='utf-8' standalone='yes'?>" & sCrLf
'sXmlClear = sXmlClear & "<?xml-stylesheet type=""text/xsl"" href='" & ZC_BLOG_HOST & "css/rss.xslt'?>" & sCrLf
'类似flickr的rss格式
sRssHead = "<rss version=""2.0"" xmlns:media=""http://search.yahoo.com/mrss/"" xmlns:dc=""http://purl.org/dc/elements/1.1/"">" & sCrLf
sRssHead = sRssHead & "<channel>" & sCrLf
sRssHead = sRssHead & "<title>"& WP_ALBUM_NAME &"</title>" & sCrLf
sRssHead = sRssHead & "<link>"& WP_SUB_DOMAIN &"</link>" & sCrLf
sRssHead = sRssHead & "<description>Powered by WindsPhoto</description>" & sCrLf
sRssHead = sRssHead & "<language>zh-cn</language>" & sCrLf
sRssHead = sRssHead & "<generator>WindsPhoto RSS</generator>" & sCrLf& sCrLf

sRssEnd = "</channel></rss>"

Response.CharSet = "utf-8"
Response.Contenttype = "text/xml"

Response.Write sXmlClear
Response.Write sRssHead

Set rs = server.CreateObject("adodb.recordset")
If Request.QueryString("id") = "" or IsNumeric(Request.QueryString("id")) = FALSE Then
    sql = "select top 20 * from WindsPhoto_desktop order by id desc"
Else
    id = Request.QueryString("id")
    sql = "select top 20 * from WindsPhoto_desktop where zhuanti="&id&" order by id desc"    
End If

rs.Open sql, objconn, 1, 1
Do While Not rs.EOF
    surl = rs("surl")
    url = rs("url")
    typeid = rs("zhuanti")
    jj = rs("jj")
    link = WP_SUB_DOMAIN & "album.asp?typeid=" & typeid
    If rs("name")<>"" then name = rs("name") Else name = "未命名" end If
    If rs("itime")<>"" then itime=ParseDateForRFC822(rs("itime")) Else itime=ParseDateForRFC822(now()) end If   
    If Left(url, 4)<>"http" Then url = WP_SUB_DOMAIN & url End If
    If InStr(url, "photo.163.com") Or InStr(url, "photo.sina.com") Or InStr(url, "126.net") Or InStr(url, "photos.baidu.com") Then url = WP_SUB_DOMAIN &"stealink.asp?" & url End If
    If Left(surl, 4)<>"http" Then surl = WP_SUB_DOMAIN & surl End If
    If InStr(surl, "photo.163.com") Or InStr(surl, "photo.sina.com") Or InStr(surl, "126.net") Or InStr(surl, "photos.baidu.com") Then surl = WP_SUB_DOMAIN &"stealink.asp?" & surl End If
    
    Set rsp = Server.CreateObject("ADODB.RecordSet")
    sqlp = "select pass from WindsPhoto_zhuanti where id="&typeid&""
    rsp.Open sqlp, objConn, 1, 1
    p = rsp("pass")
    If p<>"" Then
        name = name &"[已加密]"
        url = WP_SUB_DOMAIN &"images/nopass.gif"
        surl = WP_SUB_DOMAIN &"images/nopass.gif"
        jj = "本图片存在于加密相册内，输入密码方可查看"
    End If
    rsp.Close
    Set rsp = Nothing

    response.Write "<item>"&sCrLf
    response.Write "<title><![CDATA["&name&"]]></title>"&sCrLf    
    response.write "<pubDate>" & itime & "</pubDate>"&sCrLf   '不输出日期了?日期输出不是正确的吗?
    response.Write "<link>"& link &"</link>"&sCrLf
    response.Write "<description><![CDATA[<p align=center><a href="""& link &""" target=""_blank""><img src="""& surl &"""></a></p>"& jj &"]]></description>"&sCrLf'
    response.Write "<media:content url="""& url &""" type=""image/jpeg"" />"&sCrLf
    response.Write "<media:title>" & name & "</media:title>"&sCrLf
    response.Write "<media:text type=""html""><![CDATA["& trim(rs("jj")) &"]]></media:text>"&sCrLf   '使用cdata框住图片html文字说明,不然会出错
    response.Write "<media:thumbnail url="""& surl &""" />"&sCrLf

    response.Write "</item>"&sCrLf&sCrLf

    rs.movenext
Loop

rs.Close
Set rs = Nothing


Response.Write sRssEnd


'转换时间格式
Function ParseDateForRFC822(dtmDate)

    Dim dtmDay, dtmWeekDay, dtmMonth, dtmYear
    Dim dtmHours, dtmMinutes, dtmSeconds
    Dim TimeZone
    TimeZone = "+0800"

    Select Case WeekDay(dtmDate)
        Case 1:dtmWeekDay="Sun"
        Case 2:dtmWeekDay="Mon"
        Case 3:dtmWeekDay="Tue"
        Case 4:dtmWeekDay="Wed"
        Case 5:dtmWeekDay="Thu"
        Case 6:dtmWeekDay="Fri"
        Case 7:dtmWeekDay="Sat"
    End Select

    Select Case Month(dtmDate)
        Case 1:dtmMonth="Jan"
        Case 2:dtmMonth="Feb"
        Case 3:dtmMonth="Mar"
        Case 4:dtmMonth="Apr"
        Case 5:dtmMonth="May"
        Case 6:dtmMonth="Jun"
        Case 7:dtmMonth="Jul"
        Case 8:dtmMonth="Aug"
        Case 9:dtmMonth="Sep"
        Case 10:dtmMonth="Oct"
        Case 11:dtmMonth="Nov"
        Case 12:dtmMonth="Dec"
    End Select

    dtmYear = Year(dtmDate)
    dtmDay = Right("00" & Day(dtmDate),2)

    dtmHours = Right("00" & Hour(dtmDate),2)
    dtmMinutes = Right("00" & Minute(dtmDate),2)
    dtmSeconds = Right("00" & Second(dtmDate),2)

    ParseDateForRFC822 = dtmWeekDay & ", " & dtmDay &" " & dtmMonth & " " & dtmYear & " " & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & " " & TimeZone

End Function
%>