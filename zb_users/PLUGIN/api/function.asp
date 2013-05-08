<%
'****************************************
' api 子菜单
'****************************************
Function api_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("首页")
	aryPath=Array("main.asp")
	aryFloat=Array("m-left")
	aryInNewWindow=Array(False)
	For i=0 To Ubound(aryName)
		api_SubMenu=api_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function

'调用测试
' Base64Encode(sText1)
' Base64Decode(Base64Encode(sText1)))
Function Base64Encode(psText)
     dim oXml, oStream, oNode
     Set oXml =Server.CreateObject("MSXML2.DOMDocument")
         Set oStream =Server.CreateObject("ADODB.Stream")
             Set oNode =oXml.CreateElement("tmpNode")
                 oNode.dataType ="bin.base64"
                 oStream.Charset ="gb2312"
                 oStream.Type =2
                 If oStream.state =0 Then oStream.Open()
                 oStream.WriteText(psText)
                 oStream.Position =0
                 oStream.Type =1
                 oNode.nodeTypedValue =oStream.Read(-1)
                 oStream.Close()
                 Base64Encode =oNode.Text
             Set oNode =Nothing
         Set oStream =Nothing
     Set oXml =Nothing
End Function

Function Base64Decode(psText)
     dim oXml, oStream, oNode
     Set oXml =Server.CreateObject("MSXML2.DOMDocument")
         Set oStream =Server.CreateObject("ADODB.Stream")
             Set oNode =oXml.CreateElement("tmpNode")
                 oNode.dataType ="bin.base64"
                 oNode.Text =psText
                 oStream.Charset ="gb2312"
                 oStream.Type =1
                 oStream.Open()
                 oStream.Write(oNode.nodeTypedValue)
                 oStream.Position =0
                 oStream.Type =2
                 Base64Decode =oStream.ReadText(-1)
                 oStream.Close
             Set oNode =Nothing
         Set oStream =Nothing
     Set oXml =Nothing
End Function


Function allRand(n) '生成n位随机数字字母子组合
	Dim i,temp
	For i=1 to n
		Randomize
		temp = cint(25*Rnd)
		If temp mod 2 = 0 then
		  temp = temp + 97
		ElseIf temp < 9 then 
		  temp = temp + 48
		Else
		  temp = temp + 65
		End If
		allRand = allRand & chr(temp)
	Next
End Function
%>