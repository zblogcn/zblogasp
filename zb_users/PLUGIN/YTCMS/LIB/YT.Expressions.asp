<%
Class YT_Expressions
	Sub GetTabCollection(Text,byref Block)
		Dim RE, Match, Matches, i 
		Set RE = New RegExp 
			RE.Pattern = "(<YT:.*?>)((.|\n)+?)</YT>" 
			RE.IgnoreCase = True 
			RE.Global = True 
		Set Matches = RE.Execute(Text)
		ReDim Block(Matches.Count)
		i = 0
		For Each Match In Matches
			Block(i) = Match.Value
			i = i + 1
		Next
		Set Matches = Nothing
		Set RE = Nothing
	End Sub
	Function GeTabContent(Text,Value)
		Dim RE 
		Set RE = New RegExp 
			RE.Pattern = "(<YT:.*?\"">)((.|\n)+?)</YT>"
			RE.IgnoreCase = True 
			RE.Global = True 
		GeTabContent = RE.Replace(Text,Value)
		Set RE = Nothing
	End Function
	Function GetAttributeCollection(Text)
		Dim RE, Match, Matches, d
		Set d = CreateObject("Scripting.Dictionary")
			Text = GeTabContent(Text,"$1")
		Set RE = New RegExp 
			RE.Pattern = "([a-z]+):([a-z]+)" 
			RE.IgnoreCase = True 
			RE.Global = True
		Set Matches = RE.Execute(Text)
		For Each Match in Matches 
			If Not d.Exists(Match.SubMatches(0)) Then
				Call d.Add(Match.SubMatches(0),Match.SubMatches(1))
			End If
		Next
		Set RE = New RegExp  
			RE.Pattern = "([a-z]+)=\""(.+?)\""" 
			RE.IgnoreCase = True
			RE.Global = True 
		Set Matches = RE.Execute(Text) 
		For Each Match in Matches
			If Not d.Exists(Match.SubMatches(0)) Then
				Call d.Add(Match.SubMatches(0),Match.SubMatches(1))
			End If
		Next
		Set Matches = Nothing
		Set GetAttributeCollection = d
		Set RE = Nothing
		Set d = Nothing
	End Function
	Function YT_Each_Tab(ByVal Text,d)
		Dim RE, Match, Matches
		Dim RE2, Match2, Matches2
		Dim RE3, Match3, Matches3
		Dim TabName
		'匹配<##>标签
		Set RE = New RegExp
			RE.Pattern = "\<\#.+?\#\>" 
			RE.IgnoreCase = True 
			RE.Global = True
		Set Matches = RE.Execute(Text)
		For Each Match in Matches
			'如果为默认标签直接取值
			If d.Exists(Match.Value) Then
				Text = Replace(Text,Match.Value,d(Match.Value))
			Else
				'获取方法及参数集合
				Set RE2 = New RegExp
					RE2.Pattern = "\{([a-z_]+\:([a-z\d\'\[\]]+)\,?)+\}" 
					RE2.IgnoreCase = True
					RE2.Global = True
				Set Matches2 = RE2.Execute(Match.Value)
					For Each Match2 in Matches2
						'应用函数
						Set RE3 = New RegExp
							RE3.Pattern = "([a-z_]+)\:([a-z\d\'\[\]]+)" 
							RE3.IgnoreCase = True
							RE3.Global = True
							Set Matches3 = RE3.Execute(Match2.Value)
							'提取标签
							TabName = Replace(Match.Value,Match2.Value,"")
							For Each Match3 in Matches3
								If d.Exists(TabName) Then
									'执行函数,返回处理后的字符
									 Call Execute("d(TabName) = "&Match3.SubMatches(0)&"(d(TabName),"&Replace(Match3.SubMatches(1),"'",Chr(34))&")")
								End If
							Next
							Text = Replace(Text,Match.Value,d(TabName))
						Set Matches3 = Nothing
						Set RE3 = Nothing
					Next
				Set Matches2 = Nothing
				Set RE2 = Nothing
			End If
		Next
		Set Matches = Nothing
		Set RE = Nothing
		YT_Each_Tab = Text
	End Function
End Class
%>