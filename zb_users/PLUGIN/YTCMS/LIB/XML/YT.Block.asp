<%
Class YT_Block_XML
	Private XmlDom
	Private RootNode
	Private Sub Class_Initialize()
		Set XmlDom = CreateObject("Msxml2.DOMDocument")
		If Not XmlDom.load(BlogPath&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&YTConfig.Block) Then
			Call Create()
		Else
			Set RootNode = XmlDom.SelectSingleNode("//Root")
		End If
	End Sub
	Private Sub Class_Terminate()
		Set XmlDom = Nothing
	End Sub
	Function Add(Object,Index)
		If RootNode Is Nothing Then Exit Function
		if len(Object)>0 then set Object = YT.eval(Object)
		If IsObject(Object) Then
			Dim BlockNode,Node
				Set BlockNode = XmlDom.createElement("Block")
					Set Node = XmlDom.createElement("Name")
						Node.Text = Object.Name
						BlockNode.appendChild Node
					Set Node = Nothing
					Set Node = XmlDom.createElement("Content")
						Node.appendChild XmlDom.createCDATASection(Object.Content)
						BlockNode.appendChild Node
					Set Node = Nothing
					If Index > -1 Then
						RootNode.replaceChild BlockNode,RootNode.childNodes.item(Index)
					Else
						RootNode.appendChild BlockNode
					End If
				Set BlockNode = Nothing
			Add = Save
			Dim tpl
			Set tpl = New YT_TPL
				tpl.template = Object.Content
				Call SaveToFile(BlogPath & "ZB_USERS/INCLUDE/"&Object.Name&".asp",tpl.display(),"utf-8",True)
			Set tpl = Nothing
			Exit Function
		End If
		Add = False
	End Function
	Sub Build()
		Dim n,j,t
		For Each n In  RootNode.childNodes
			j=n.childNodes(0).Text
			t=n.childNodes(1).Text
			t=TransferHTML(t,"[no-asp]")
			Dim tpl
			Set tpl = New YT_TPL
				tpl.template = t
				Call SaveToFile(BlogPath & "ZB_USERS/INCLUDE/"&j&".asp",tpl.display(),"utf-8",True)
			Set tpl = Nothing
		Next
	End Sub
	Function Del(Index)
		RootNode.RemoveChild RootNode.childNodes.item(Index)
		Del = Save
	End Function
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
		XmlDom.Save BlogPath&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&YTConfig.Block
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