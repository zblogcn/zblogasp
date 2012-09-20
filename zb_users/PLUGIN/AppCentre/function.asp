<%


Dim App_ID,App_Name,App_URL,App_Note,App_PubDate
Dim App_Adapted,App_Version,App_Modified
Dim App_Type,App_Path,App_Include,App_Level
Dim App_Author_Name,App_Author_Url,App_Author_Email

Dim Action,SelectedPlugin,SelectedPluginName
Dim objXmlVerChk,NewVersionExists

Const DownLoad_URL="http://download.rainbowsoft.org/Plugins/ps.asp"
Const Resource_URL="http://download.rainbowsoft.org/Plugins/"    '注意. Include 文件里还有一同名变量要修改
Const Update_URL="http://download.rainbowsoft.org/Plugin/dlcs/download2.asp?plugin="

Const XML_Pack_Ver="1.0"
Const XML_Pack_Type="Plugin"
Const XML_Pack_Version="Z-Blog_2——0"




Sub SubMenu(id)
	Dim aryName,aryValue,aryPos
	aryName=Array("插件下载","主题下载")
	aryValue=Array("plugin_list.asp","theme_list.asp")
	aryPos=Array("m-left","m-left")
	Dim i 
	For i=0 To Ubound(aryName)
		Response.Write MakeSubMenu(aryName(i),aryValue(i),aryPos(i) & IIf(id=i," m-now",""),False)
	Next
End Sub


Sub ExportPluginList()
	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s, t
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "/ZB_USERS/PLUGIN/")
	Set fc = f.SubFolders

	Dim aryPL
	aryPL=Split(ZC_USING_PLUGIN_LIST,"|")

	For Each s in aryPL

		App_Author_Name=Empty
		App_Author_Url=Empty
		App_Author_Email=Empty

		App_ID=Empty
		App_Name=Empty
		App_URL=Empty
		App_Modified=Empty
		App_Version=Empty
		App_Note=Empty

		strXmlFile =BlogPath & "/ZB_USERS/PLUGIN/" & s & "/" & "Plugin.xml"
		If fso.FileExists(strXmlFile) Then

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					App_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					App_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					App_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					App_ID=objXmlFile.documentElement.selectSingleNode("id").text
					App_Name=objXmlFile.documentElement.selectSingleNode("name").text
					App_URL=objXmlFile.documentElement.selectSingleNode("url").text
					App_Modified=objXmlFile.documentElement.selectSingleNode("modified").text
					App_Version=objXmlFile.documentElement.selectSingleNode("version").text
					App_Note=objXmlFile.documentElement.selectSingleNode("note").text

					App_Name=TransferHTML(App_Name,"[html-format]")
					App_Note=TransferHTML(App_Note,"[nohtml]")

				End If
			End If
			Set objXmlFile=Nothing




			Response.Write "<div class=""pluginPanel"">"
			Response.Write "<div class=""listTitle"" onclick='showhidediv("""& App_ID &""");'>"

			Response.Write "<div class=""delicon""><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(App_ID) & "&amp;pluginname=" & Server.URLEncode(App_Name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(App_ID) & """ title=""导出此插件""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(App_ID) & """ title=""编辑插件信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""PluginDetail.asp?plugin=" & Server.URLEncode(App_ID) & "&amp;pluginname=" & Server.URLEncode(App_Name) & """ title=""查看插件信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & App_ID) & """ title=""升级修复插件""><font color=""Gray""><b>↓</b></font></a></div>"

			Response.Write "<div class=""updinfo"">"& App_Name &" Ver "& App_Version &" <font color=""Green"">(启用中)</font> "& t &"</div>"
			Response.Write "</div>"


			Response.Write "<p><table width=""98%"" id="""& App_ID &""" style=""display:none;"">"

			Response.Write "<tr>"

			Response.Write "<td width=""32"" align=""right"">ID:</td><td>"
			Response.Write "" & App_ID & ""
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">名称:</td><td>"
			If App_URL=Empty Then
				Response.Write "" & App_Name & ""
			Else
				Response.Write "<a href=""" & App_URL & """ target=""_blank"" title=""插件发布地址"">" & App_Name & "</a>"
			End If
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">作者:</td><td>"
			If App_Author_Url=Empty Then
				Response.Write "" & App_Author_Name & ""
			Else
				Response.Write "<a href=""" & App_Author_Url & """ target=""_blank"" title=""作者主页"">" & App_Author_Name & "</a>"
			End If
			If Not App_Author_Email=Empty Then Response.Write " (" & App_Author_Email & ")"
			Response.Write "</td>"

			Response.Write "<td width=""64"" align=""right"">最后发布:</td><td width=""80"">" & App_Modified & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right"">简介:</td><td colspan=7>" & App_Note & "</td>"
			Response.Write "</tr>"
			Response.Write "</table></p>"

			Response.Write "</div>"

		End If
	Next

	For Each f1 in fc

		App_Author_Name=Empty
		App_Author_Url=Empty
		App_Author_Email=Empty

		App_ID=Empty
		App_Name=Empty
		App_URL=Empty
		App_Modified=Empty
		App_Version=Empty
		App_Note=Empty


		If fso.FileExists(BlogPath & "/zb_users/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then

			strXmlFile =BlogPath & "/zb_users/PLUGIN/" & f1.name & "/" & "Plugin.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					App_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					App_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					App_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					'PluginID=f1.name
					App_ID=objXmlFile.documentElement.selectSingleNode("id").text
					App_Name=objXmlFile.documentElement.selectSingleNode("name").text
					App_URL=objXmlFile.documentElement.selectSingleNode("url").text
					App_Modified=objXmlFile.documentElement.selectSingleNode("modified").text
					App_Version=objXmlFile.documentElement.selectSingleNode("version").text
					App_Note=objXmlFile.documentElement.selectSingleNode("note").text

					App_Name=TransferHTML(App_Name,"[html-format]")
					App_Note=TransferHTML(App_Note,"[nohtml]")

				End If
			End If
			Set objXmlFile=Nothing

			If CheckPluginState(App_ID) Then
			Else

			If fso.FileExists(BlogPath & "/zb_users/PLUGIN/" & App_ID & "/" & "verchk.xml") Then
				t="<a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & App_ID) & """ title=""升级插件""><b class=""notice"">发现新版本!</b></a>"
				NewVersionExists=True
			ElseIf fso.FileExists(BlogPath & "/zb_users/PLUGIN/" & App_ID & "/" & "error.log") Then
				t="<b class=""somehow"">不支持在线更新.</b>"
			Else
				t=""
			End If

			Response.Write "<div class=""pluginPanel pluginPanelAlt"">"
			Response.Write "<div class=""listTitle listTitleAlt"" onclick='showhidediv("""& App_ID &""");'>"

			Response.Write "<div class=""delicon""><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(App_Name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""导出此插件""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""编辑插件信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""PluginDetail.asp?plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(App_Name) & """ title=""查看插件信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & App_ID) & """ title=""升级修复插件""><font color=""Gray""><b>↓</b></font></a></div>"

			If UCase(App_ID)<>UCase(f1.name) Then
				Response.Write "<div>该插件ID错误, 请 <a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""编辑插件信息""><font color=""red""><b>[重新编辑插件信息]</b></font></a>.</div>"
			Else
				Response.Write "<div class=""updinfo"">"& App_Name &" Ver "& App_Version &" <font color=""Orange"">(停用中)</font> "& t &"</div>"
			End If

			Response.Write "</div>"


			Response.Write "<p><table width=""98%"" id="""& App_ID &""" style=""display:none;"">"

			Response.Write "<tr>"

			Response.Write "<td width=""32"" align=""right"">ID:</td><td>"
			Response.Write "" & App_ID & ""
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">名称:</td><td>"
			If App_URL=Empty Then
				Response.Write "" & App_Name & ""
			Else
				Response.Write "<a href=""" & App_URL & """ target=""_blank"" title=""插件发布地址"">" & App_Name & "</a>"
			End If
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">作者:</td><td>"
			If App_Author_Url=Empty Then
				Response.Write "" & App_Author_Name & ""
			Else
				Response.Write "<a href=""" & App_Author_Url & """ target=""_blank"" title=""作者主页"">" & App_Author_Name & "</a>"
			End If
			If Not App_Author_Email=Empty Then Response.Write " (" & App_Author_Email & ")"
			Response.Write "</td>"

			Response.Write "<td width=""64"" align=""right"">最后发布:</td><td>" & App_Modified & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right"">简介:</td><td colspan=7>" & App_Note & "</td>"
			Response.Write "</tr>"
			Response.Write "</table></p>"

			Response.Write "</div>"

			End If

		End If

	Next

	For Each f1 in fc
		If fso.FileExists(BlogPath & "/zb_users/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then
		Else

			App_ID=f1.name

			Response.Write "<div class=""pluginPanel"" style=""background-color:#FFFFFF;"">"
			Response.Write "<div class=""listTitle"" style=""border-bottom:1px dotted #BDD3EF;background:#EDEDED;"">"

			Response.Write "<div class=""delicon""><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(f1.name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""导出此插件""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""编辑插件信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""PluginDetail.asp?plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(f1.name) & """ title=""查看插件信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & App_ID) & """ title=""升级修复插件""><font color=""Gray""><b>↓</b></font></a></div>"

			Response.Write "<div>ID: "& App_ID &"</div>"

			Response.Write "</div>"

			Response.Write "<p>该插件信息不完全, 并不是完整的 Z-Blog 插件.</p>"
			Response.Write "</div>"

		End If

	Next
	Set fso = nothing
	Err.Clear
%>
<!-- 		<div class="PluginPanel" style="background-color:#FFFFFF;">
		<p><a href="Xml_Upload.asp" title="导入本地的 ZPI 文件"><img src="Images/import.png" alt="ScreenShot" width="200" height="160" /></a></p>
			<p><b>从本地导入 ZPI 文件:</b><br />	<form border="1" name="edit" method="post" enctype="multipart/form-data" action="XML_Upload.asp?act=FileUpload"><p>选择插件安装包文件,TS 将从该文件导入插件并安装到 Plugin 目录下: </p><p><input type="file" id="edtFileLoad" name="edtFileLoad" size="15"></p><p><input type="submit" class="button" value="提交" name="B1" onclick="return window.confirm('确定导入该插件数据包??');" /> <input class="button" type="reset" value="重置" name="B2" /></p></form></p>
		</div> -->

		<hr style="clear:both;"/><p><form name="edit" method="get" action="#"  class="status-box">
			<p><input onclick="window.scrollTo(0,0);" type="button" class="button" value="TOP" title="返回页面顶部" /> <input onclick="self.location.href='Xml_ChkVer.asp?act=check&n=0';" type="button" class="button" value="查找更新" title="开始查找插件的可用更新" /></p>
		</form></p>
<%
	If NewVersionExists Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('newVersion').style.display = 'block';</script>"
	End If
	Response.Flush

	Dim FileList,l,c
	FileList=LoadIncludeFiles("ZB_USERS/PLUGIN/AppCentre/Export/")

	For Each l In FileList
		c=c & l
	Next

	If (InStr(LCase(c),".xml")>0) Or (InStr(LCase(c),".zpi")>0) Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('edit').style.display = 'block';</script>"
	End If

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

'End If

End Sub

%>