<%
Class YT_Model_XML
	Private XmlDom
	Private RootNode
	Private Sub Class_Initialize()
		Set XmlDom = CreateObject("Msxml2.DOMDocument")
		If Not XmlDom.load(BlogPath&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&YTConfig.Model) Then
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
			Dim ModelNode,Nodes,Node,Field
				Set ModelNode = XmlDom.createElement("Model")
					Set Nodes = XmlDom.createElement("Table")
						Set Node = XmlDom.createElement("Name")
							Node.Text = Object.Table.Name
							Nodes.appendChild Node
						Set Node = Nothing
						Set Node = XmlDom.createElement("Description")
							Node.Text = Object.Table.Description
							Nodes.appendChild Node
						Set Node = Nothing
						Set Node = XmlDom.createElement("Bind")
							Node.Text = Object.Table.Bind
							Nodes.appendChild Node
						Set Node = Nothing
						ModelNode.appendChild Nodes
					Set Nodes = Nothing
						For Each Field In Object.Fields
							Set Nodes = XmlDom.createElement("Field")
								Set Node = XmlDom.createElement("Name")
									Node.Text = Field.Name
									Nodes.appendChild Node
								Set Node = Nothing
								Set Node = XmlDom.createElement("Description")
									Node.Text = Field.Description
									Nodes.appendChild Node
								Set Node = Nothing
								Set Node = XmlDom.createElement("Value")
									Node.appendChild XmlDom.createCDATASection(Field.Value)
									Nodes.appendChild Node
								Set Node = Nothing
								Set Node = XmlDom.createElement("Property")
									Node.appendChild XmlDom.createCDATASection(Field.Property)
									Nodes.appendChild Node
								Set Node = Nothing
								Set Node = XmlDom.createElement("Type")
									Node.Text = Field.Type
									Nodes.appendChild Node
								Set Node = Nothing
								ModelNode.appendChild Nodes
							Set Nodes = Nothing
						Next
					If Index > -1 Then
						RootNode.replaceChild ModelNode,RootNode.childNodes.item(Index)
						'//修改已经安装的表字段,只(增/改)不删字段
						Dim t,old,s1,s2
						Set t=new YT_Table
							If t.Exist(Object.Table.Name) Then
								old=t.GetFields(Object.Table.Name)
								s2="ALTER TABLE ["&Object.Table.Name&"] ADD COLUMN "
								For Each Field In Object.Fields
									If t.FieldExist(old,Field.Name) Then
										s1="ALTER TABLE ["&Object.Table.Name&"] ALTER COLUMN "
										s1=s1&"["&Field.Name&"]"
										s1=s1&" "&Field.Property
										objConn.Execute(s1)
									Else
										s2=s2&"["&Field.Name&"]"
										s2=s2&" "&Field.Property
										s2=s2&","
									End If
								Next
								If Right(s2,1)="," Then objConn.Execute(Left(s2,Len(s2)-1))
							End If
						Set t=Nothing
					Else
						RootNode.appendChild ModelNode
					End If
				Set ModelNode = Nothing
			Add = Save
			Exit Function
		End If
		Add = False
	End Function
	Function Length()
		Length=RootNode.childNodes.Length
	End Function
	Function Model(Action,Index)
		If RootNode Is Nothing Then Exit Function
		Dim Nodes,Node
		Set Nodes = RootNode.selectNodes("Model")
		Dim YTTable,i:i=0
		Set YTTable = new YT_Table
			For Each Node In Nodes
				If Action = "Install" Then
					If Not YTTable.Exist(Node.selectSingleNode("Table/Name").Text) Then
						If Not IsNumeric(Index) Then
							Model = YTTable.Create(Node)
							Exit Function
						Else
							If Int(Index) = i Then
								Model = YTTable.Create(Node)
								Exit Function
							End If
						End If
					End If
				Else
					If YTTable.Exist(Node.selectSingleNode("Table/Name").Text) Then
						If Not IsNumeric(Index) Then
							Model = YTTable.Delete(Node)
							Exit Function
						Else
							If Int(Index) = i Then
								Model = YTTable.Delete(Node)
								Exit Function
							End If
						End If
					End If
				End If
				i=i+1
			Next
		Set YTTable = Nothing
		Set Nodes = Nothing
	End Function
	Function Del(Index)
		RootNode.RemoveChild RootNode.childNodes.item(Index)
		Del = Save
	End Function
	Function GetModel(Cate)
		If RootNode Is Nothing Then Exit Function
		Dim Nodes,Node,Bind,isBind,i
		Set Nodes = RootNode.selectNodes("Model")
			For Each Node In Nodes
				Bind = Split(Node.selectSingleNode("Table/Bind").Text,",")
				For i = LBound(Bind) To UBound(Bind)
					If Int(Bind(i)) = Int(Cate) Then
						isBind = True
						Exit For
					End If
				Next
				If isBind Then
					Set GetModel = Node
					Exit Function
				End If
			Next
		Set Nodes = Nothing
		Set GetModel = Nothing
	End Function
	Function Create()
		Dim Header
		Set Header = XmlDom.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8""")
			XmlDom.appendChild Header
		Set Header = Nothing
		Set RootNode = XmlDom.createElement("Root")
			XmlDom.appendChild  RootNode
		Create = Save
	End Function
	Function Save()
		On Error Resume Next
		Save=false
		XmlDom.Save BlogPath&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&YTConfig.Model
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