Option Strict Off
Option Explicit On
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Private Declare Function LoadImage Lib "user32.dll"  Alias "LoadImageA"(ByVal hInst As Integer, ByVal lpsz As String, ByVal un1 As Integer, ByVal n1 As Integer, ByVal n2 As Integer, ByVal un2 As Integer) As Integer
	'UPGRADE_ISSUE: 不支持将参数声明为“As Any”。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"”
	Private Declare Function SendMessage Lib "user32.dll"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Const WM_SETICON As Integer = &H80s
	Private Const ICON_SMALL As Integer = 0
	Private Const IMAGE_ICON As Integer = 1
	Private Const LR_DEFAULTSIZE As Integer = &H40s
	Private Const LR_LOADFROMFILE As Integer = &H10s
	
	
	
	
	Dim strSource, strTemplateFolder, strXMLPath As String
	Dim aryTemplateFile() As String
	Dim aryPluginFile() As String
	Dim objAero As clsAero
	
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
		Dim strTemp As String
		'UPGRADE_WARNING: 未能解析对象 GetFolderPath() 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		strTemp = GetFolderPath("请选择模板文件夹", Me.Handle.ToInt32)
		If Not strTemp = "False" Then
			strTemplateFolder = strTemp
			txtPath.Text = strTemp
			Log("选择模板文件夹：" & strTemp)
		End If
	End Sub
	
	Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
		Dim i As Short
		Log_Clear()
		strTemplateFolder = txtPath.Text
		If GetSubFolder(strTemplateFolder) Then
			Log("开始升级模板文件")
			For i = 0 To UBound(aryTemplateFile)
				If Trim(aryTemplateFile(i)) <> "" Then Update_Renamed(aryTemplateFile(i), 1) : Update_Renamed(aryTemplateFile(i), 4)
			Next 
			Log("模板文件升级完毕")
			Log("开始升级source下asp")
			Update_Renamed(strSource, 2)
			Log("source下asp升级完毕")
			Log("开始升级主题自带插件")
			For i = 0 To UBound(aryPluginFile)
				If Trim(aryPluginFile(i)) <> "" Then Update_Renamed(aryPluginFile(i), 3)
			Next 
			Log("主题自带插件升级完毕")
			'还差侧栏管理的升级
			'Log "升级XML信息"
			'升级XML是不是让APP升级好一点
			
			MsgBox(Replace("升级完毕！\n\n剩余以下部分没有升级，请自行修改：\n\n侧栏部分（须符合2.0侧栏规范）\n主题插件\nXML信息\n\n升级完成后，请在APP中心里编辑主题信息并保存，即可在2.0里激活主题。", "\n", vbCrLf), MsgBoxStyle.Information)
		End If
		
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Call GetSystemVersion()
		
		If bolAero Then
			objAero = New clsAero
			'UPGRADE_ISSUE: Form 属性 frmMain.hDc 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"”
			objAero.hDc = Me.hDc
			objAero.hWnd = Me.Handle.ToInt32
			objAero.Init()
		End If
		
		'UPGRADE_NOTE: 在对对象 Me.Icon 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		Me.Icon = Nothing
		Dim hIcon As Integer
		hIcon = LoadImage(0, My.Application.Info.DirectoryPath & "\zblog.ico", IMAGE_ICON, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE)
		If hIcon Then
			SendMessage(Me.Handle.ToInt32, WM_SETICON, ICON_SMALL, hIcon)
		End If
		
		
		objRegExp = New VBScript_RegExp_55.RegExp
		objFSO = New Scripting.FileSystemObject
		objADO = CreateObject("ADODB.Stream")
		strTemplateFolder = ""
		objRegExp.Global = True
		objRegExp.IgnoreCase = True
		ReDim aryTemplateFile(0)
		ReDim aryPluginFile(0)
		strSource = ""
		lblNote.Text = "说明：" & vbCrLf & "升级前必须备份。" & vbCrLf & "您要升级的1.8模板必须符合以下要求：" & vbCrLf & "      1.模板在TEMPLATE文件夹下，扩展名为html" & vbCrLf & "      2.HTML标签全部闭合" & vbCrLf & "      3.未重写系统自带的common.js" & vbCrLf & "      4.未使用主题插件" & vbCrLf & "以上条件有任意一点不符合，则本程序无法升级你的主题。"
	End Sub
	
	Private Sub frmMain_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		If bolAero Then
			objAero.Paint()
		End If
	End Sub
	
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: 在对对象 objRegExp 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		objRegExp = Nothing
		'UPGRADE_NOTE: 在对对象 objFSO 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		objFSO = Nothing
		'UPGRADE_NOTE: 在对对象 objADO 进行垃圾回收前，不可以将其销毁。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"”
		objADO = Nothing
	End Sub
	
	
	
	
	'Usage:日志
	'Param:str--日志内容
	'UPGRADE_NOTE: str 已升级到 str_Renamed。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"”
	Sub Log(ByVal str_Renamed As String)
		lstLog.Items.Add("【" & Now & "】" & str_Renamed)
	End Sub
	
	'Usage:清除日志
	Sub Log_Clear()
		lstLog.Items.Clear()
	End Sub
	
	'Usage:扫描文件夹
	'Param:Folder--文件夹
	Function GetSubFolder(ByVal Folder As String) As Boolean
		objRegExp.Pattern = "b_article-guestbook|b_article_trackback|guestbook|search"
		GetSubFolder = False
		Dim objSub As Object
		Dim objFor As Object
		If objFSO.FolderExists(Folder) Then
			If objFSO.FileExists(Folder & "/theme.xml") Then
				strXMLPath = objFSO.GetFile(Folder & "/theme.xml").Path
				Log("找到主题XML信息")
			Else
				Log("主题XML不存在")
				Exit Function
			End If
			If objFSO.FolderExists(Folder & "/template") Then
				For	Each objFor In objFSO.GetFolder(Folder & "/template").Files
					
					
					'顺便做个删除吧
					
					'UPGRADE_WARNING: 未能解析对象 objFor.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					If objRegExp.Test(objFor.Name) Then
						'UPGRADE_WARNING: 未能解析对象 objFor.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log("删除无用文件：" & objFor.Name)
						'UPGRADE_WARNING: 未能解析对象 objFor.Path 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						objFSO.DeleteFile(objFor.Path)
					Else
						ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
						'复制page模板
						'UPGRADE_WARNING: 未能解析对象 objFor.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						If objFor.Name Like "single*" Then
							If Not objFSO.FileExists(Folder & "/template/page.html") Then
								'UPGRADE_WARNING: 未能解析对象 objFor.Path 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
								objFSO.CopyFile(objFor.Path, Folder & "/template/page.html") : Log("复制PAGE模板")
							End If
						End If
						'UPGRADE_WARNING: 未能解析对象 objFor.Path 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						aryTemplateFile(UBound(aryTemplateFile)) = objFor.Path
						'UPGRADE_WARNING: 未能解析对象 objFor.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log("找到主题文件：" & objFor.Name)
					End If
				Next objFor
			End If
			If objFSO.FolderExists(Folder & "/include") Then
				For	Each objFor In objFSO.GetFolder(Folder & "/include").Files
					ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
					'UPGRADE_WARNING: 未能解析对象 objFor.Path 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					aryTemplateFile(UBound(aryTemplateFile)) = objFor.Path
					'UPGRADE_WARNING: 未能解析对象 objFor.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					Log("找到主题文件：" & objFor.Name)
				Next objFor
			End If
			If objFSO.FolderExists(Folder & "/plugin") Then
				For	Each objFor In objFSO.GetFolder(Folder & "/plugin").Files
					ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
					'UPGRADE_WARNING: 未能解析对象 objFor.Path 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					aryPluginFile(UBound(aryPluginFile)) = objFor.Path
					'UPGRADE_WARNING: 未能解析对象 objFor.Name 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					Log("找到主题插件：" & objFor.Name)
				Next objFor
			End If
			If objFSO.FileExists(Folder & "/source/style.css.asp") Then
				strSource = objFSO.GetFile(Folder & "/source/style.css.asp").Path
				Log("找到STYLE.CSS.ASP")
			End If
			GetSubFolder = True
		Else
			Log("文件夹不存在！")
		End If
	End Function
	
	
	'Usage:得到XML信息以判断是否Z-Blog
	'Param:XMLPath--XML地址
	Function LoadXMLInfo(ByVal XMLPath As String) As Boolean
		
	End Function
	
	
	
	
	'Usage:升级
	'Param:strFilePath--文件名,intType--升级类型
	'UPGRADE_NOTE: Update 已升级到 Update_Renamed。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"”
	Function Update_Renamed(ByVal strFilePath As String, Optional ByRef intType As Short = 1) As Boolean
		Dim vbSpace As Object
		Dim strFile As String
		Dim objExec As Object
		If objFSO.FileExists(strFilePath) Then
			Log("Update: " & strFilePath & "  type:" & intType)
			'UPGRADE_WARNING: 未能解析对象 LoadFromFile() 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			strFile = LoadFromFile(strFilePath)
			Select Case intType
				Case 1
					'模板主体和INCLUDE文件夹升级
					
					
					'替换zb_system下文件
					objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)"
					
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.SubMatches(0), 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.SubMatches(0) & "-->" & "zb_system/" & objExec.SubMatches(0))
					Next objExec
					
					'替换zb_users下文件
					objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(plugin|language|cache|upload)"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_users/" & objExec.SubMatches(0), 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.SubMatches(0) & "-->" & "zb_users/" & objExec.SubMatches(0))
					Next objExec
					
					'替换theme
					objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>themes)"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>zb_users/theme", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>zb_users/theme")
					Next objExec
					
					'替换rss
					objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>rss\.xml)"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>feed.asp", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.SubMatches 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>feed.asp")
					Next objExec
					
					
					'替换那些玩意
					objRegExp.Pattern = "var (str0[0-9]|intMaxLen|strBatchView|strBatchInculde|strBatchCount|strFaceName|strFaceSize)=.+?;"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					'强插c_html_js_add.asp
					If InStr(LCase(strFile), "c_html_js_add.asp") = 0 And InStr(LCase(strFile), "</head>") > 0 Then
						strFile = Replace(strFile, "</head>", "<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js_add.asp"" type=""text/javascript""></script>" & vbCrLf & "</head>")
						Log("强制插入c_html_js_add.asp")
					End If
					
					'删除无用UBB部分
					objRegExp.Pattern = "InsertQuote.+?\;|ExportUbbFrame\(\)\;?"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					
					'替换计数
					strFile = Replace(strFile, "strBatchCount+=""spn<#article/id#>=<#article/id#>,""", "AddViewCount(<#article/id#>)")
					strFile = Replace(strFile, "strBatchView+=""spn<#article/id#>=<#article/id#>,""", "LoadViewCount(<#article/id#>)")
					Log("计数部分修改")
					
					'替换无用标签
					objRegExp.Pattern = "<#template:article_trackback#>|<#article/pretrackback_url#>|<#ZC_MSG014#>|<#article/trackbacknums#>"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					'替换Try--elScript
					objRegExp.Pattern = "try{" & vbCrLf & ".+?elScript[\d\D]+?catch\(e\){};?"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					
					'替换验证码
					objRegExp.Pattern = "if.+?inpVerify[\d\D]+?Math.random\(\)[\d\D]+?}[\d\D]+?}"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					'替换空行
					'UPGRADE_WARNING: 未能解析对象 vbSpace 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					objRegExp.Pattern = "[" & vbTab & vbSpace & "]+" & vbCrLf
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: 未能解析对象 objExec.Value 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					
					
					'保存
					SaveToFile(strFilePath, strFile)
					Log("保存完毕")
				Case 2
					'SOURCE\STYLE.CSS.ASP升级
					
					'替换<%
					strFile = Replace(strFile, "<%", "<!-- #include file=""../../../../zb_system/function/c_function.asp"" -->" & vbCrLf & "<%")
					Log("引用c_function.asp")
					
					'替换路径
					strFile = Replace(strFile, """themes""", """zb_users/theme""")
					Log("""themes"" --> ""zb_users/theme""")
					
					'替换HOST
					strFile = Replace(strFile, "ZC_BLOG_HOST", "GetCurrentHost()")
					Log("ZC_BLOG_HOST --> GetCurrentHost()")
					
					
					SaveToFile(strFilePath, strFile)
					Log("保存完毕")
					
				Case 3
					'插件\主题插件升级
				Case 4
					'侧栏管理升级
					'侧栏管理只按照默认主题的结构弄，非默认主题的结构不管他
					'抽样调查20个主题，默认主题侧栏结构约占50%上下
					
					objRegExp.Pattern = "<div id=""divSidebar"">[\d\D]+?<div class=""function"""
					'判断是否存在结构与默认主题相同的侧栏
					If objRegExp.Test(strFile) Then
						
						'objRegExp.Pattern = "<div id=""divSidebar"">[\d\D]+?</div>"
						
					End If
					
				Case 5
					'XML升级
					
			End Select
		Else
			Log(strFile & "找不到！")
		End If
	End Function
End Class