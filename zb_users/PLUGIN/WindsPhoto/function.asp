<%
'********************************************************************************
'函数名：ISO8601
'功  能：转换时间格式
'参  数：时间
'返  回：
'********************************************************************************
Function ISO8601(DateTime)
    Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond

    DateTime = DateAdd("h", -8, DateTime)
    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    DateWeek = Weekday(DateTime)
    DateSecond = Second(DateTime)
    If Len(DateMonth)<2 Then DateMonth = "0"&DateMonth
    If Len(DateDay)<2 Then DateDay = "0"&DateDay
    If Len(DateMinute)<2 Then DateMinute = "0"&DateMinute
    If Len(DateHour)<2 Then DateHour = "0"&DateHour
    If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
    ISO8601 = Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&"Z"
End Function

'********************************************************************************
'函数名：RemoveHTML
'功  能：过滤HTML
'参  数：HTML字符串
'返  回：
'********************************************************************************
Function RemoveHTML(strHTML)
    Dim objRegExp, Match, Matches
    Set objRegExp = New Regexp
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    '取闭合的<>
    objRegExp.Pattern = "<.+?>"
    '进行匹配
    Set Matches = objRegExp.Execute(strHTML)
    ' 遍历匹配集合，并替换掉匹配的项目
    For Each Match in Matches
    strHtml=Replace(strHTML,Match.Value,"")
    Next
    RemoveHTML=strHTML
    Set objRegExp = Nothing
End Function

'*********************************************************
'函数名：GetPhotoIndex
'功   能：输出相册首页主体部分
'参   数：
'返   回：
'*********************************************************
Function GetPhotoIndex()
    Dim ipagecount
    Dim ipagecurrent
    Dim strorderBy
    Dim irecordsshown
    If request.querystring("page") = "" Then
        ipagecurrent = 1
    Else
        ipagecurrent = CInt(request.querystring("page"))
    End If

    Dim tpid, sql, sql2, sql3
    Set rso = Server.CreateObject("ADODB.RecordSet")
    sql = "select count(*) as C from desktop"
    If WP_ORDER_BY = "0" Then
        sql2 = "select * from zhuanti where pass<>'no' order by ordered,id asc"
    Else
        sql2 = "select * from zhuanti where pass<>'no' order by ordered,id asc "
    End If
    rso.Open sql, Conn, 3, 3
    sm = rso("c")
    rso.Close
    Set rso = Nothing
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    rs2.pagesize = WP_INDEX_PAGERCOUNT
    rs2.Open sql2, Conn, 1, 1
    ipagecount = rs2.pagecount
    If ipagecurrent > ipagecount Then ipagecurrent = ipagecount
    If ipagecurrent < 1 Then ipagecurrent = 1
    If ipagecount = 0 Then
        GetPhotoIndex = GetPhotoIndex&"<div><p align='center'>没有任何相册</p>"
    Else
        rs2.absolutepage = ipagecurrent
        irecordsshown = 0
        GetPhotoIndex = GetPhotoIndex&"<p>所有相册情况，截止到"&Now()&" 共有"&rs2.RecordCount&"个相册,"&sm&"张图片。</p>" & VBCRLF
        GetPhotoIndex = GetPhotoIndex & WP_ALBUM_INTRO & VBCRLF
        GetPhotoIndex = GetPhotoIndex&"<table width='100%' border='0' cellspacing='0' cellpadding='5'>" & VBCRLF
        Do While irecordsshown<WP_INDEX_PAGERCOUNT And Not rs2.EOF
            GetPhotoIndex = GetPhotoIndex&"<tr align='center'>"

            For i = 1 To 3

                If Not rs2.EOF Then
                    sqlp = "select * from desktop where zhuanti="&rs2("id")&" and hot<>0 order by id asc"
                    Set rsp = Server.CreateObject("ADODB.RecordSet")
                    rsp.Open sqlp, Conn, 1, 1
                    If rsp.EOF Or rsp.bof Then
                        surl = WP_SUB_DOMAIN &"images/notop.gif"
                    Else
                        surl = rsp("surl")
                        If Left(surl, 4)<>"http" Then surl = WP_SUB_DOMAIN & surl End If
                        If InStr(surl, "photo.163.com") Or InStr(surl, "126.net") Or InStr(surl, "photo.sina.com") Or InStr(surl, "photos.baidu.com") Then surl = WP_SUB_DOMAIN &"stealink.asp?" & surl End If
                    End If
                    rsp.Close
                    Set rsp = Nothing
                    Set rso = Server.CreateObject("ADODB.RecordSet")
                    sql = "select count(*) as C from desktop where zhuanti="&rs2("id")&""
                    rso.Open sql, Conn, 3, 3
                    sm = rso("c")
                    rso.Close
                    Set rso = Nothing
                    Dim sqlp
                    Set rsp = Server.CreateObject("ADODB.RecordSet")
                    sqlp = "select pass from zhuanti where id="&rs2("id")&""
                    rsp.Open sqlp, Conn, 1, 1
                    p = rsp("pass")
                    If p<>"" Then
                        surl = WP_SUB_DOMAIN &"images/nopass.gif"
                    End If
                    rsp.Close
                    Set rsp = Nothing
                    GetPhotoIndex = GetPhotoIndex&"<td width='30%'><a href='"& WP_SUB_DOMAIN &"album.asp?typeid="&rs2("id")&"' title='"&rs2("name")&"'><img class='wp_top' src='"&surl&"' alt='"&rs2("name")&"' onload='WindsPhotoResizeImage(this,"&WP_SMALL_WIDTH&","&WP_SMALL_HEIGHT&")' /></a><br /><a href='"& WP_SUB_DOMAIN &"album.asp?typeid="&rs2("id")&"' title='"&rs2("name")&"'>"&rs2("name")&" | "&sm&"张</a></td>" & VBCRLF
                    irecordsshown = irecordsshown + 1
                    rs2.movenext
                End If

            Next

            GetPhotoIndex = GetPhotoIndex&"</tr>" & VBCRLF
        Loop
    End If
    GetPhotoIndex = GetPhotoIndex&"</table>" & VBCRLF

        '分页
    if ipagecount >1 then
        GetPhotoIndex=GetPhotoIndex&"<div class=""post pagebar""><span class=""other-page"">"&ipagecount&"页中的第"&ipagecurrent&"页</span>"
        if ipagecurrent=1 then
            GetPhotoIndex=GetPhotoIndex&"<span class=""other-page"">上一页</span>"
        else
            GetPhotoIndex=GetPhotoIndex&"<a href='default.asp?page="&ipagecurrent-1&"'>上一页</a>"
        end if

        for i = 1 to ipagecount
        ipagenow = ipagenow + 1
        if ipagecurrent=ipagenow then
            GetPhotoIndex=GetPhotoIndex&"<span class=""now-page"">"&ipagenow&"</span>"
        else
        GetPhotoIndex=GetPhotoIndex&"<a href='default.asp?page="&ipagenow&"'>"&ipagenow&"</a>"
            end if
        next

        if ipagecount>ipagecurrent then
            GetPhotoIndex=GetPhotoIndex&"<a href='default.asp?page="&ipagecurrent+1&"'>下一页</a>"
        else
            GetPhotoIndex=GetPhotoIndex&"<span class=""other-page"">下一页</span>"
        end if

        GetPhotoIndex=GetPhotoIndex&"</div>" & VBCRLF
    end if

    rs2.Close
    Set rs2 = Nothing

End Function

'*********************************************************
'函数名：SaveSortList
'功   能：输出相册分类目录列表
'参   数：
'返   回：
'*********************************************************
Function SaveSortList()
    Dim rssort, rssortcount, sqlsort1, sqlsort2, countsort
    Set rssort = Server.CreateObject("ADODB.RecordSet")
    sqlsort1= "select top 10 * from zhuanti where pass='' order by ordered,id asc"
    rssort.Open sqlsort1, Conn, 1, 1
    Do While Not rssort.EOF
        sqlsort2 = "select count(*) as C from desktop where zhuanti="&rssort("id")
        Set rssortcount = Server.CreateObject("ADODB.RecordSet")
        rssortcount.Open sqlsort2, Conn, 3, 3
        countsort = rssortcount("C")
        SortList = SortList&"<li><span class='feed-icon'><a href='"& WP_SUB_DOMAIN &"rss.asp?id="&rssort("id")&"' target='_blank'><img title='rss' width='20' height='12' src='http://www.wilf.cn/IMAGE/LOGO/rss.png' border='0' alt='rss' /></a>&nbsp;</span><a href='"& WP_SUB_DOMAIN &"album.asp?typeid="&rssort("id")&"'>"&rssort("name")&"<span class='article-nums'> ("&countsort&")</span></a></li>"&Chr(13)&Chr(10)
        rssortcount.Close
        Set rssortcount = Nothing
        rssort.movenext
    Loop
    rssort.Close
    Set rssort = Nothing

    Call SaveToFile(BlogPath & "zb_users/include/windsphoto_sort.asp", SortList, "utf-8", TRUE)

End Function
%>