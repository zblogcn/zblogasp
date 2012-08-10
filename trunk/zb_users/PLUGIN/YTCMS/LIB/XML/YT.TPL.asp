<%
Class YT_TPL_XML
	Private XmlDom
	Private RootNode
	Private Sub Class_Initialize()
		Set XmlDom = CreateObject("Msxml2.DOMDocument")
		If Not XmlDom.load(BlogPath&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&Config.TPL) Then
			Call Create()
		Else
			Set RootNode = XmlDom.SelectSingleNode("//Root")
		End If
	End Sub
	Private Sub Class_Terminate()
		Set XmlDom = Nothing
	End Sub
	Sub Load()
		Dim Node
		For Each Node In RootNode.selectNodes("TPL")
			If Node.selectSingleNode("Type").Text="Single" Then
				Config.Single.push(jsonToObject("{t:'"&Node.selectSingleNode("File").Text&"',v:["&Node.selectSingleNode("Bind").Text&"]}"))
			Else
				Config.Multi.push(jsonToObject("{t:'"&Node.selectSingleNode("File").Text&"',v:["&Node.selectSingleNode("Bind").Text&"]}"))
			End If
		Next
	End Sub
	Function Add(Object,Index)
		If RootNode Is Nothing Then Exit Function
		If IsObject(Object) Then
			Dim TPLNode,Node
				Set TPLNode = XmlDom.createElement("TPL")
					Set Node = XmlDom.createElement("File")
						Node.Text = Object.File
						TPLNode.appendChild Node
					Set Node = Nothing
					Set Node = XmlDom.createElement("Bind")
						Node.Text = Object.Bind
						TPLNode.appendChild Node
					Set Node = Nothing
					Set Node = XmlDom.createElement("Type")
						Node.Text = Object.Type
						TPLNode.appendChild Node
					Set Node = Nothing
					If Index > -1 Then
						RootNode.replaceChild TPLNode,RootNode.childNodes.item(Index)
					Else
						RootNode.appendChild TPLNode
					End If
				Set TPLNode = Nothing
			Add = Save
			Exit Function
		End If
		Add = False
	End Function
	Sub Del(Index)
		RootNode.RemoveChild RootNode.childNodes.item(Index)
		Call Save()
	End Sub
	Sub Create()
		Dim Header
		Set Header = XmlDom.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8""")
			XmlDom.appendChild Header
		Set Header = Nothing
		Set RootNode = XmlDom.createElement("Root")
			XmlDom.appendChild  RootNode
		Call Save()
	End Sub
	Function Save()
		On Error Resume Next
		Save=false
		XmlDom.Save BlogPath&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&Config.TPL
		If Err.number<>0 then
			'//Response.Write Err.Description
			Err.clear
			Save=false
			Exit Function
		End If
		Save = true
	End Function
End Class
%>