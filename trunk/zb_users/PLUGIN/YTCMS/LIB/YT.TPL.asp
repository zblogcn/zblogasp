<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷
'// 技术支持:    33195@qq.com
'// 程序名称:    	模板类
'// 开始时间:    	2012.08.30
'// 最后修改:    2012.10.24
'///////////////////////////////////////////////////////////////////////////////
Class YT_TPL
	Private REG
	public cache
	Private sIndex
	private tpl
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
		s = s&vbTab&"For Each h{%i%} In Split(YTARRAY,"","")"&vbcrlf
		s = s&vbTab&vbTab&"Execute(h{%i%}&""=empty"")"&vbcrlf
		s = s&vbTab&"Next"&vbcrlf
		s = s&vbTab&"YTARRAY=Empty"&vbCrlf
		s = s&vbTab&"End If"&vbCrlf
		s = s&"Next"&vbCrlf
		s = s&"End If"&vbCrlf
		s = s&"%"&">"&vbCrlf
		s = s&vbNullChar
		s = s&"Model{%i%} = getModel($1{%i%}.CateID,$1{%i%}.ID)"&vbcrlf
		s = s&vbTab&vbTab&"If len(Trim(Model{%i%}))>0 Then"&vbcrlf
		s = s&vbTab&vbTab&"Set Model{%i%}=YT.eval(Model{%i%})"&vbcrlf
		s = s&vbTab&vbTab&vbTab&"If isObject(Model{%i%}) Then"&vbcrlf
		s = s&vbTab&vbTab&vbTab&vbTab&"Execute(""YTARRAY=Model{%i%}.YTARRAY"")"&vbcrlf
		s = s&vbTab&vbTab&vbTab&vbTab&"For Each k{%i%} In Split(YTARRAY,"","")"&vbcrlf
		s = s&vbTab&vbTab&vbTab&vbTab&vbTab&"Execute(k{%i%}&""=YT.unescape(Model{%i%}.""&k{%i%}&"")"")"&vbcrlf
		s = s&vbTab&vbTab&vbTab&vbTab&"Next"&vbcrlf
		s = s&vbTab&vbTab&vbTab&"End If"&vbCrlf
		s = s&vbTab&vbTab&"End If"&vbCrlf
		code = s
	End Property
	Public Property Let template(s)
		if len(cache) > 0 then
			dim st,cPath
			cPath = "ZB_USERS/CACHE/"
			st = LoadFromFile(BlogPath&cPath&cache&".TPL","utf-8")
			if len(st) = 0 then
				tpl = parse(s)
				'//緩存文件
				Call SaveToFile(BlogPath&cPath&cache&".TPL", tpl,"utf-8",False)
			else
				'//加載緩存
				tpl = st
			end if
		else
			tpl = parse(s)
		end if
	End Property
	Private Sub Class_Initialize()
			sIndex = 0
		Set REG = New Regexp
			REG.IgnoreCase = True
			REG.Global = True
	End Sub
	Private Sub Class_Terminate()
		tpl = empty
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
		s = reg_replace("\{code\}","<"&"%",s)
		s = Replace(s,"{/code}","%"&">")
		s = reg_replace("\{eval\s+(.+?)\}","<"&"% $1 %"&">",s)
		s = reg_replace("\{echo\s+(.+?)\}","<"&"%=$1%"&">",s)
		s = reg_replace("\{template\:(.+?)\}","<"&"%YT_TPL_display(GetTemplate(UCase(""TEMPLATE_$1"")))%"&">",s)
		s = reg_replace("\<\#CACHE_INCLUDE_(.+?)\#\>","<"&"%=LoadFromFile(BlogPath&""zb_users/include/$1.asp"",""utf-8"")%"&">",s)
		s = reg_replace("\{cache\:(.+?)\}","<"&"%=LoadFromFile(BlogPath&""zb_users/include/$1.asp"",""utf-8"")%"&">",s)
		s = reg_replace("\<\#(ZC_[a-z,_]+)\#\>","<"&"%=$1%"&">",s)
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
			s = s&vbTab&vbTab&"htm = htm&"""&t&""""&vbcrlf
			a = instr(b,str,"%\>")+2
			s = s&reg_replace("^\s*=",vbTab&vbTab&"htm = htm&",mid(str,b,a-b-2))&vbcrlf
			b = instr(a,str,"<%")+2
		Wend
		t = mid(str,a)
		t = replace(t,vbcrlf,"{::vbcrlf}")
		t = replace(t,vblf,"{::vblf}")
		t = replace(t,vbcr,"{::vbcr}")
		t = replace(t,"""","""""")
		s = s&"htm = htm&"""&t&""""&vbcrlf
		s = replace(s,"response.write","htm = htm&",1,-1,1)
		s = replace(s,"YT_TPL_display(GetTemplate(UCase(","htm = htm&YT_TPL_display(GetTemplate(UCase(",1,-1,1)
		ob_get_contents = s
	End Function
	Private Function crlf(t)
		dim s
		s = t:crlf = False
		s = replace(s,vbcrlf,"")
		s = replace(s,vblf,"")
		s = replace(s,vbcr,"")
		s = trim(replace(s,vbTab,""))
		If len(s)=0 Then crlf = True
	End Function
	private function parse(str)
		Dim d,e,h,j,p,r,t,x,s
		Dim scripts,script,arrScript()
			str = TransferHTML(str,"[japan-html]")
			scripts = reg_match("(\<script(?![^<>]*?src[^<>]*?>).*?>[\s\S]*?\<\/script\>)",str)
			ReDim arrScript(-1)
			For Each script In scripts
				j = UBound(arrScript) + 1
				ReDim Preserve arrScript(j)
				arrScript(j) = script(0)
				str = replace(str,arrScript(j),"{::SCRIPT["&j&"]}")
			Next
			str = html_replace(str)
			d = reg_match("\{YT\:([a-z]+)([^\}]+)\}",str)
			For Each e In d
				REG.Global = True
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
				str = reg_replace("\{YT\:"&e(0)&"([^\}]+)\}",Split(s,vbNullChar)(0),str)
				str = reg_replace("\{\/YT\:"&e(0)&"\}",Split(s,vbNullChar)(1),str)

				p = str_match(str,Split(s,vbNullChar)(0),Split(s,vbNullChar)(1))
				If Not isEmpty(p) Then
					r = p:REG.Global = True
					t = reg_match("\{YT\:([a-z]+)([^\}]+)\}",r)
					x = Replace(Replace(r,Split(s,vbNullChar)(0),""),Split(s,vbNullChar)(1),"")
					x = reg_replace(e(0)&"\.([a-z]+)",e(0)&sIndex&"."&"$1",x)
					r = Split(s,vbNullChar)(0)&x&Split(s,vbNullChar)(1)
					do while UBound(t)>-1
						sIndex = sIndex + 1
						t = reg_match("\{YT\:([a-z]+)([^\}]+)\}",parse(str))
					loop
					str = Replace(str,p,r)
				End If
				sIndex = sIndex + 1
			Next
			j = 0
			For Each script In arrScript
				str = replace(str,"{::SCRIPT["&j&"]}",script)
				j = j + 1
			Next
			parse = ob_get_contents(str)
	end function
	Public Function display()
		On Error Resume Next
		Execute(tpl)
		If Err.Number<>0 then
			dim s
			s = "<fieldset><legend>$1</legend>$2</fieldset>"
			s = Replace(s,"$1","Content Manage System")
			s = Replace(s,"$2",Err.Source&vbTab&Err.Number&vbTab&Err.Description)
			Err.Clear
			display = s
		Else
			htm = replace(htm,"{::vbcrlf}",vbcrlf)
			htm = replace(htm,"{::vblf}",vblf)
			htm = replace(htm,"{::vbcr}",vbcr)
			htm = TransferHTML(htm,"[html-japan]")
			dim ui,i,htm2
			htm2 = empty
			ui = split(htm,vblf)
			for each i in ui
				if not crlf(i) then
					htm2 = htm2 & i & vblf
				end if
			next
			display = htm2
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
	Public Function reg_match(Pattern,s)
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
	Public Function reg_replace(Pattern,s,html)
		Dim t:t = False
		REG.Pattern = Pattern
		t = REG.Replace(html,s)
		reg_replace = t
	End Function
	Private Function getAttribute(k,v)
		Dim a,b:b=False
		For Each a in k
			If a(0) = v Then b=a(1):Exit For
		Next
		getAttribute=b
	End Function
End Class
%>