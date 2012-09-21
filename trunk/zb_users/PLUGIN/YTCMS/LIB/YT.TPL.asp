<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷
'// 技术支持:    33195@qq.com
'// 程序名称:    	模板类
'// 开始时间:    	2012.08.30
'// 最后修改:    2012.09.03
'///////////////////////////////////////////////////////////////////////////////
Class YT_TPL
	Private REG,DIC
	Private sIndex
	Private sTemplate
	Public Property Get template
		template = sTemplate
	End Property
	Public Property Let template(s)
		'替换标签兼容旧系统
		sTemplate = reg_replace("\<(YT\:([a-z]+)[^\>]+)\>([\s\S]+?)\<(\/YT)\>","{$1}$3{$4:$2}",s)
	End Property
	Public Property Get code
		Dim s
		s = "<"&"%"&vbCrlf
		s = s&"Set $1{%i%} = New T$1"&vbCrlf
		s = s&vbTab&"a{%i%} = New YT_$1.$2"&vbCrlf
		s = s&"If isArray(a{%i%}) Then"&vbCrlf
		s = s&"For x{%i%} = LBound(a{%i%},2) To UBound(a{%i%},2)"&vbCrlf
		s = s&vbTab&"If $1{%i%}.LoadInfoByID(a{%i%}(0,x{%i%})) Then"&vbCrlf
		s = s&vbTab&vbTab&"$3"&vbCrlf
		s = s&vbTab&vbTab&"$4"&vbCrlf
		s = s&"%"&">"&vbCrlf
		s = s&vbNullChar
		s = s&"<"&"%"&vbCrlf
		s = s&vbTab&"End If"&vbCrlf
		s = s&"Next"&vbCrlf
		s = s&"End If"&vbCrlf
		s = s&"%"&">"&vbCrlf
		s = s&vbNullChar
		s = s&"Model{%i%} = getModel($1{%i%}.CateID,$1{%i%}.ID)"&vbcrlf
		s = s&vbTab&vbTab&"If Not isEmpty(Model{%i%}) Then Set Model{%i%}=YT.eval(Model{%i%})"&vbcrlf
		s = s&vbTab&vbTab&"If isObject(Model{%i%}) Then"&vbcrlf
		s = s&vbTab&vbTab&"For Each k{%i%} In Model{%i%}.YTARRAY"&vbcrlf
		s = s&vbTab&vbTab&vbTab&"Execute(k{%i%}&""=YT.unescape(Model{%i%}.""&k{%i%}&"")"")"&vbcrlf
		s = s&vbTab&vbTab&"Next"&vbcrlf
		s = s&vbTab&vbTab&"End If"&vbcrlf
		code = s
	End Property
	Private Sub Class_Initialize()
			sIndex = 0
		Set REG = New Regexp
			REG.IgnoreCase = True
			REG.Global = True
		Set DIC = CreateObject("Scripting.Dictionary")	
			DIC.Add "<#article/trackback_url#>","<"&"%=$1{%i%}.TrackBack%"&">"
			DIC.Add "<#article/category/name#>","<"&"%=$1{%i%}.HtmlName%"&">"
			DIC.Add "<#article/category/url#>","<"&"%=$1{%i%}.HtmlUrl%"&">"
			DIC.Add "<#article/author/url#>","<"&"%=$1{%i%}.HtmlUrl%"&">"
			DIC.Add "<#article/author/level#>","<"&"%=ZVA_User_Level_Name(Users($1{%i%}.AuthorID).Level)%"&">"
			DIC.Add "<#article/posttime/longdate#>","<"&"%=FormatDateTime($1{%i%}.PostTime,vbLongDate)%"&">"
			DIC.Add "<#article/posttime/shortdate#>","<"&"%=FormatDateTime($1{%i%}.PostTime,vbShortDate)%"&">"
			DIC.Add "<#article/posttime/longtime#>","<"&"%=FormatDateTime($1{%i%}.PostTime,vbLongTime)%"&">"
			DIC.Add "<#article/posttime/shorttime#>","<"&"%=FormatDateTime($1{%i%}.PostTime,vbShortTime)%"&">"
			DIC.Add "<#article/posttime/year#>","<"&"%=Year($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/month#>","<"&"%=Month($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/monthname#>","<"&"%=ZVA_Month(Month($1{%i%}.PostTime))%"&">"
			DIC.Add "<#article/posttime/day#>","<"&"%=Day($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/weekday#>","<"&"%=Weekday($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/weekdayname#>","<"&"%=ZVA_Week(Weekday($1{%i%}.PostTime))%"&">"
			DIC.Add "<#article/posttime/hour#>","<"&"%=Hour($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/minute#>","<"&"%=Minute($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/second#>","<"&"%=Second($1{%i%}.PostTime)%"&">"
			DIC.Add "<#article/posttime/monthnameabbr#>","<"&"%=ZVA_Month_Abbr(Month($1{%i%}.PostTime))%"&">"
			DIC.Add "<#article/posttime/weekdaynameabbr#>","<"&"%=ZVA_Week_Abbr(Weekday($1{%i%}.PostTime))%"&">"
			DIC.Add "<#article/commentrss#>","<"&"%=Second($1{%i%}.WfwCommentRss)%"&">"
			DIC.Add "<#article/commentposturl#>","<"&"%=TransferHTML($1{%i%}.CommentPostUrl,""[html-format]"")%"&">"
			DIC.Add "<#article/pretrackback_url#>","<"&"%=TransferHTML($1{%i%}.PreTrackBack,""[html-format]"")%"&">"
			DIC.Add "<#article/comment/name#>","<"&"%=$1{%i%}.Author%"&">"
			DIC.Add "<#article/comment/url#>","<"&"%=$1{%i%}.HomePage%"&">"
			DIC.Add "<#article/comment/email#>","<"&"%=$1{%i%}.SafeEmail%"&">"
			DIC.Add "<#article/comment/urlencoder#>","<"&"%=$1{%i%}.HomePageForAntiSpam%"&">"
			DIC.Add "<#article/tag/name#>","<"&"%=$1{%i%}.HtmlName%"&">"
			DIC.Add "<#article/tag/intro#>","<"&"%=$1{%i%}.HtmlIntro%"&">"
	End Sub
	Private Sub Class_Terminate()
		Set DIC = Nothing
		Set REG = Nothing
	End Sub
	Private Function html_replace(str)
		Dim s
		s   = str
		REG.Global = True
		s = reg_replace("\<\!\-\-\{(.+?)\}\-\-\>","{$1}",s)
		s = reg_replace("\{\$(.+?)\}", "<"&"%=$1%"&">",s)
		s = reg_replace("\{foreach\s+(.+?)\s+(.+?)\}", "<"&"%For Each $1 In $2%"&">",s)
		s = reg_replace("\{for\s+(.+?)\s+(.+?)\}", "<"&"%For $1 To $2%"&">",s)
		s = Replace(s,"{/next}","<"&"%Next%"&">")
		s = reg_replace("\{(do|loop)\s+(while|until)\s+(.+?)\}","<"&"%$1 $2 $3%"&">",s)
		s = Replace(s,"{do}","<"&"%Do%"&">")
		s = Replace(s,"{loop}","<"&"%Loop%"&">")
		s = reg_replace("\{while\s+(.+?)\}"	,"<"&"%While $1%"&">",s)
		s = Replace(s,"{/wend}","<"&"%Wend%"&">")
		s = reg_replace("\{if\s+(.+?)\}","<"&"%if $1 Then%"&">",s)
		s = reg_replace("\{elseif\s+(.+?)\}","<"&"%ElseIf $1 Then%"&">",s)	
		s = Replace(s,"{/if}","<"&"%End If%"&">")
		s = Replace(s,"{else}","<"&"%Else%"&">")
		s = reg_replace("\{eval\s+(.+?)\}","<"&"% $1 %"&">",s)
		s = reg_replace("\{echo\s+(.+?)\}","<"&"%=$1%"&">",s)
		html_replace = s
	End Function
	Private Function ob_get_contents(str)
		Dim s, a, b, t, matches, m
		s = "dim htm : htm = """""&vbcrlf
		a = 1:REG.Global = True
		b = instr(a,str,"<%")+2
		While b > a+1
			t = mid(str,a,b-a-2)
			t = replace(t,vbcrlf,"{::vbcrlf}")
			t = replace(t,vblf,"{::vblf}")
			t = replace(t,vbcr,"{::vbcr}")
			t = replace(t,"""","""""")
			If crlf(t)=False Then s = s&vbTab&vbTab&"htm = htm&"""&t&""""&vbcrlf
			a = instr(b,str,"%\>")+2
			s = s&reg_replace("^\s*=",vbTab&vbTab&"htm = htm&",mid(str,b,a-b-2))&vbcrlf
			b = instr(a,str,"<%")+2
		Wend
		t = mid(str,a)
		t = replace(t,vbcrlf,"{::vbcrlf}")
		t = replace(t,vblf,"{::vblf}")
		t = replace(t,vbcr,"{::vbcr}")
		t = replace(t,"""","""""")
		If crlf(t)=False Then s = s&"htm = htm&"""&t&""""&vbcrlf
		s = replace(s,"response.write","htm = htm&",1,-1,1)
		ob_get_contents = s
	End Function
	Private Function crlf(t)
		crlf = False
		t = Trim(Replace(t,vbTab,Empty))
		If t = "{::vblf}" Or t = "{::vbcrlf}" Or t = "{::vbcr}" Then
			crlf = True
		End If
	End Function
	Public Function display()
		On Error Resume Next
		Dim d,e,h,j,k,l,m,n,p,r,t,u,v,w,x
		Dim s
		w = template
		template = html_replace(template)
		d = reg_match("\{YT\:([a-z]+)([^\>]+)\}",template)
		For Each e In d
			REG.Global = True
			m = DIC.Keys
			n = DIC.Items
			For l = 0 To DIC.Count - 1
				template=Replace(template,m(l),Replace(Replace(n(l),"{%i%}",sIndex),"$1",e(0)))
			Next
			s = code
			If LCase(e(0)) = "article" Then s = Replace(s,"$4",Split(s,vbNullChar)(2))
			s = Replace(s,"$4",Empty)
			s = Replace(s,"$1",e(0))
			h = reg_match("([a-z]+)\=""([^""]+)""",e(1))
			If isArray(h) Then
				j = getAttribute(h,"DataSource")
				If j <> False Then  s = Replace(s,"$2",Replace(j,"'",""""))
				j = getAttribute(h,"Name")
				If j <> False Then s = Replace(s,"$3","Set "&j&" = "&e(0)&sIndex)
			End If
			s = Replace(s,"{%i%}",sIndex)
			s = Replace(s,"$3",Empty)
			REG.Global = False
			template = reg_replace("\{YT\:"&e(0)&"([^\}]+)\}",Split(s,vbNullChar)(0),template)
			template = reg_replace("\{\/YT\:"&e(0)&"\}",Split(s,vbNullChar)(1),template)

			p = str_match(template,Split(s,vbNullChar)(0),Split(s,vbNullChar)(1))
			If Not isEmpty(p) Then
				r = p:REG.Global = True
				t = reg_match("\{YT\:([a-z]+)([^\}]+)\}",r)
				x = Replace(Replace(r,Split(s,vbNullChar)(0),""),Split(s,vbNullChar)(1),"")
				x = reg_replace(e(0)&"\.([a-z]+)",e(0)&sIndex&"."&"$1",x)
				r = Split(s,vbNullChar)(0)&x&Split(s,vbNullChar)(1)
				While UBound(t)>-1
					sIndex = sIndex + 1
					display()
					t = reg_match("\{YT\:([a-z]+)([^\}]+)\}",template)
				Wend
				r = reg_replace("\<#article\/([a-z]+)#\>","<"&"%="&e(0)&sIndex&".$1%"&">",r)
				r = reg_replace("\<#article\/comment\/([a-z]+)#\>","<"&"%="&e(0)&sIndex&".$1%"&">",r)
				r = reg_replace("\<#article\/category\/([a-z]+)#\>","<"&"%=Categorys("&e(0)&sIndex&".CateID).$1%"&">",r)
				r = reg_replace("\<#article\/author\/([a-z]+)#\>","<"&"%=Users("&e(0)&sIndex&".AuthorID).$1%"&">",r)
				r = reg_replace("\<#article\/model\/(\w+)#\>","<"&"%=$1%"&">"&"<"&"%$1=Empty%"&">",r)
				template = Replace(template,p,r)
			End If
			sIndex = sIndex + 1
		Next
		'Response.Write template
		k = ob_get_contents(template)
		Execute(k)
		If Err.Number<>0 then
			s = "<fieldset><legend>$1</legend>$2</fieldset>"
			s = Replace(s,"$1",Err.Source)
			s = Replace(s,"$2",Err.Number&vbTab&Err.Description)
			Err.Clear
			s = s&w
			display = s
		Else
			htm = replace(htm,"{::vbcrlf}",vbcrlf)
			htm = replace(htm,"{::vblf}",vblf)
			htm = replace(htm,"{::vbcr}",vbcr)
			display = htm
		End If
	End Function
	Private Function mfe(k,v)
		Dim b,s:s = Empty
		v = UCase(Trim(v))
		For Each b in k
			If UCase(Trim(b)) = v Then s = v:Exit For
		Next
		mfe = s
	End Function
	Private Function getModel(CateID,ID)
		Dim x,n,s:s = Empty
		If isNumeric(CateID) And isNumeric(ID) Then
			Set x = new YT_Model_XML
			Set n = x.GetModel(CateID)
				If Not n Is Nothing Then
					s = YT_Data_GetRow(n.selectSingleNode("Table/Name").Text,ID)
				End If
			Set n = Nothing
			Set x = Nothing
		End If
		getModel = s
	End Function
	Private Function str_match(s,b,e)
		Dim l,r,a,d:d = Empty
		l = inStr(s,b)
		If l > 0 Then
			a = Mid(s,l,Len(s))
			r = inStr(a,e)
			If r > 0 Then
				d = Mid(a,1,r+Len(e))
			End If
		End If
		str_match = d
	End Function
	Private Function reg_match(Pattern,s)
		Dim ms,m,i,j:j=0
			REG.Pattern = Pattern
		Set ms = REG.Execute(s)
		Dim a:a=Array()
		Dim b:b=Array()
		For Each m in ms
			For i=0 To m.SubMatches.Count - 1
				ReDim Preserve a(i)
				a(i)=m.SubMatches(i)
			Next
			i=UBound(a)+1
			ReDim Preserve a(i)
			a(i)=m.Value
			ReDim Preserve b(j)
			b(j)=a
			j=j+1
		Next
		Set ms = Nothing
		reg_match=b
	End Function
	Private Function reg_replace(Pattern,s,html)
		Dim t:t = False
		REG.Pattern = Pattern
		t = REG.Replace(html,s)
		reg_replace = t
	End Function
	Function getAttribute(k,v)
		Dim a,b:b=False
		For Each a in k
			If a(0) = v Then b=a(1):Exit For
		Next
		getAttribute=b
	End Function
End Class
%>