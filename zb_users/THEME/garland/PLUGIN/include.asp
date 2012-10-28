<%

'注册插件
Call RegisterPlugin("garland","ActivePlugin_Garland")

'具体的接口挂接
Function ActivePlugin_Garland() 

	'Action_Plugin_MakeCalendar_Begin
	Call Add_Action_Plugin("Action_Plugin_MakeCalendar_Begin","MakeCalendar=WP_MakeCalendar2(dtmYearMonth):Exit Function")

End Function

'*********************************************************
' 目的：    WP 之 Make Calendar2
'*********************************************************
Function WP_MakeCalendar2(dtmYearMonth)

	Dim strCalendar

	Dim y
	Dim m
	Dim d
	Dim firw
	Dim lasw
	Dim ny
	Dim nm
	Dim py
	Dim pm

	Dim i
	Dim j
	Dim k
	Dim b

	Call CheckParameter(dtmYearMonth,"dtm",Date())

	y=year(dtmYearMonth)
	m=month(dtmYearMonth)
	ny=y
	nm=m+1
	If m=12 Then ny=ny+1:nm=1
	py=y
	pm=m-1
	if m=1 then py=py-1:pm=12

	firw=Weekday(Cdate(y&"-"&m&"-1"))

	For i=28 to 32
		If IsDate(y&"-"&m&"-"&i) Then
			lasw=Weekday(Cdate(y&"-"&m&"-"&i))
		Else
			Exit For
		End If
	Next

	d=i-1
	k=1

	If firw>5 Then b=42 Else b=35
	If (d=28) And (firw=1) Then b=28
	If (firw>5) And (d<31) And (d-firw<>23) Then b=35


'//////////////////////////////////////////////////////////
'	逻辑处理
	Dim aryDateLink(32)
	Dim aryDateID(32)
	Dim aryDateArticle(32)
	Dim objRS

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	objRS.Open("select [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] from [blog_Article] where ([log_Type]=0) And ([log_Level]>2) And ([log_PostTime] BETWEEN "& ZC_SQL_POUND_KEY &y&"-"&m&"-1"& ZC_SQL_POUND_KEY &" AND "& ZC_SQL_POUND_KEY &ny&"-"&nm&"-1"& ZC_SQL_POUND_KEY &")")

	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 To objRS.RecordCount
			j=CLng(Day(CDate(objRS("log_PostTime"))))
			aryDateLink(j)=True
			aryDateID(j)=objRS("log_ID")
			Set aryDateArticle(j)=New TArticle
			aryDateArticle(j).LoadInfobyArray Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	objRS.Close
	Set objRS=Nothing
'//////////////////////////////////////////////////////////

	strCalendar="<table summary=""日历"" id=""wp-calendar""><caption>"&y&"年"&m&"月</caption>"
	
'thead	
	strCalendar=strCalendar & "	<thead>	<tr> <th title=""星期日"" scope=""col"" abbr=""星期日"">日</th> <th title=""星期一"" scope=""col"" abbr=""星期一"">一</th> <th title=""星期二"" scope=""col"" abbr=""星期二"">二</th>	<th title=""星期三"" scope=""col"" abbr=""星期三"">三</th> <th title=""星期四"" scope=""col"" abbr=""星期四"">四</th>	<th title=""星期五"" scope=""col"" abbr=""星期五"">五</th> <th title=""星期六"" scope=""col"" abbr=""星期六"">六</th>	</tr>	</thead>"
	
'tfoot	
	dim strCalendarPrev
	dim strCalendarNext
	
	strCalendarPrev = "<td id=""prev"" colspan=""3"" abbr="""& pm &"月""><a title=""查看"& pm &"月的日志"" href="""& ZC_BLOG_HOST & "catalog.asp?date="& py &"-"& pm &""">« "&ZVA_Month_Abbr(pm)&"</a></td>"
	strCalendarNext = "<td id=""next"" colspan=""3"" abbr="""& nm &"月""><a title=""查看"& nm &"月的日志"" href="""& ZC_BLOG_HOST & "catalog.asp?date="& ny &"-"& nm &"""> "&ZVA_Month_Abbr(nm)&" »</a></td>"
		
	if dtmYearMonth=Date()  Then strCalendarNext = "<td class=""pad"" id=""next"" colspan=""3""> </td>"

	strCalendar=strCalendar & "	<tfoot>	<tr>" & strCalendarPrev & " <td class=""pad""> </td>" & strCalendarNext & "</tr></tfoot>"	
	
'tbody	
	strCalendar=strCalendar & "	<tbody>"
	
	j=0
	For i=1 to b

		If (j Mod 7)=0 Then strCalendar=strCalendar & "<tr>"
		If (j/7)<=0 and firw<>1 then strCalendar=strCalendar & "<td class=""pad"" colspan="""& (firw-1) &"""> </td>"

		If (j=>firw-1) and (k=<d) Then
		
			strCalendar=strCalendar & "<td "
			
			If 	Cdate(y&"-"&m&"-"&k) = Date() Then strCalendar=strCalendar & " id =""today"" "
			
			If aryDateLink(k) Then
				strCalendar=strCalendar & "><a  title=""点击查看当天文章"" href="""& ZC_BLOG_HOST &"catalog.asp?date="&Year(aryDateArticle(k).PostTime)&"-"&Month(aryDateArticle(k).PostTime)&"-"&Day(aryDateArticle(k).PostTime)& """>"&(k)&"</a></td>"
			Else
				strCalendar=strCalendar &">"&(k)&"</td>"
				
			End If

			k=k+1
		End If
			
		if j=b-1 then strCalendar=strCalendar & "<td class=""pad"" colspan="""& (7-lasw) &"""> </td>"		

		If (j Mod 7)=6 Then strCalendar=strCalendar & "</tr>"

		j=j+1
	Next

	strCalendar=strCalendar & "	</tbody></table>"
	WP_MakeCalendar2=strCalendar
	
End Function
'*********************************************************
%>