Option Strict Off
Option Explicit On
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Private Declare Function LoadImage Lib "user32.dll"  Alias "LoadImageA"(ByVal hInst As Integer, ByVal lpsz As String, ByVal un1 As Integer, ByVal n1 As Integer, ByVal n2 As Integer, ByVal un2 As Integer) As Integer
	'UPGRADE_ISSUE: ��֧�ֽ���������Ϊ��As Any���� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"��
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
		'UPGRADE_WARNING: δ�ܽ������� GetFolderPath() ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		strTemp = GetFolderPath("��ѡ��ģ���ļ���", Me.Handle.ToInt32)
		If Not strTemp = "False" Then
			strTemplateFolder = strTemp
			txtPath.Text = strTemp
			Log("ѡ��ģ���ļ��У�" & strTemp)
		End If
	End Sub
	
	Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
		Dim i As Short
		Log_Clear()
		strTemplateFolder = txtPath.Text
		If GetSubFolder(strTemplateFolder) Then
			Log("��ʼ����ģ���ļ�")
			For i = 0 To UBound(aryTemplateFile)
				If Trim(aryTemplateFile(i)) <> "" Then Update_Renamed(aryTemplateFile(i), 1) : Update_Renamed(aryTemplateFile(i), 4)
			Next 
			Log("ģ���ļ��������")
			Log("��ʼ����source��asp")
			Update_Renamed(strSource, 2)
			Log("source��asp�������")
			Log("��ʼ���������Դ����")
			For i = 0 To UBound(aryPluginFile)
				If Trim(aryPluginFile(i)) <> "" Then Update_Renamed(aryPluginFile(i), 3)
			Next 
			Log("�����Դ�����������")
			'����������������
			'Log "����XML��Ϣ"
			'����XML�ǲ�����APP������һ��
			
			MsgBox(Replace("������ϣ�\n\nʣ�����²���û���������������޸ģ�\n\n�������֣������2.0�����淶��\n������\nXML��Ϣ\n\n������ɺ�����APP������༭������Ϣ�����棬������2.0�Ｄ�����⡣", "\n", vbCrLf), MsgBoxStyle.Information)
		End If
		
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Call GetSystemVersion()
		
		If bolAero Then
			objAero = New clsAero
			'UPGRADE_ISSUE: Form ���� frmMain.hDc δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"��
			objAero.hDc = Me.hDc
			objAero.hWnd = Me.Handle.ToInt32
			objAero.Init()
		End If
		
		'UPGRADE_NOTE: �ڶԶ��� Me.Icon ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
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
		lblNote.Text = "˵����" & vbCrLf & "����ǰ���뱸�ݡ�" & vbCrLf & "��Ҫ������1.8ģ������������Ҫ��" & vbCrLf & "      1.ģ����TEMPLATE�ļ����£���չ��Ϊhtml" & vbCrLf & "      2.HTML��ǩȫ���պ�" & vbCrLf & "      3.δ��дϵͳ�Դ���common.js" & vbCrLf & "      4.δʹ��������" & vbCrLf & "��������������һ�㲻���ϣ��򱾳����޷�����������⡣"
	End Sub
	
	Private Sub frmMain_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		If bolAero Then
			objAero.Paint()
		End If
	End Sub
	
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: �ڶԶ��� objRegExp ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
		objRegExp = Nothing
		'UPGRADE_NOTE: �ڶԶ��� objFSO ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
		objFSO = Nothing
		'UPGRADE_NOTE: �ڶԶ��� objADO ������������ǰ�������Խ������١� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"��
		objADO = Nothing
	End Sub
	
	
	
	
	'Usage:��־
	'Param:str--��־����
	'UPGRADE_NOTE: str �������� str_Renamed�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"��
	Sub Log(ByVal str_Renamed As String)
		lstLog.Items.Add("��" & Now & "��" & str_Renamed)
	End Sub
	
	'Usage:�����־
	Sub Log_Clear()
		lstLog.Items.Clear()
	End Sub
	
	'Usage:ɨ���ļ���
	'Param:Folder--�ļ���
	Function GetSubFolder(ByVal Folder As String) As Boolean
		objRegExp.Pattern = "b_article-guestbook|b_article_trackback|guestbook|search"
		GetSubFolder = False
		Dim objSub As Object
		Dim objFor As Object
		If objFSO.FolderExists(Folder) Then
			If objFSO.FileExists(Folder & "/theme.xml") Then
				strXMLPath = objFSO.GetFile(Folder & "/theme.xml").Path
				Log("�ҵ�����XML��Ϣ")
			Else
				Log("����XML������")
				Exit Function
			End If
			If objFSO.FolderExists(Folder & "/template") Then
				For	Each objFor In objFSO.GetFolder(Folder & "/template").Files
					
					
					'˳������ɾ����
					
					'UPGRADE_WARNING: δ�ܽ������� objFor.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					If objRegExp.Test(objFor.Name) Then
						'UPGRADE_WARNING: δ�ܽ������� objFor.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log("ɾ�������ļ���" & objFor.Name)
						'UPGRADE_WARNING: δ�ܽ������� objFor.Path ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						objFSO.DeleteFile(objFor.Path)
					Else
						ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
						'����pageģ��
						'UPGRADE_WARNING: δ�ܽ������� objFor.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						If objFor.Name Like "single*" Then
							If Not objFSO.FileExists(Folder & "/template/page.html") Then
								'UPGRADE_WARNING: δ�ܽ������� objFor.Path ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
								objFSO.CopyFile(objFor.Path, Folder & "/template/page.html") : Log("����PAGEģ��")
							End If
						End If
						'UPGRADE_WARNING: δ�ܽ������� objFor.Path ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						aryTemplateFile(UBound(aryTemplateFile)) = objFor.Path
						'UPGRADE_WARNING: δ�ܽ������� objFor.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log("�ҵ������ļ���" & objFor.Name)
					End If
				Next objFor
			End If
			If objFSO.FolderExists(Folder & "/include") Then
				For	Each objFor In objFSO.GetFolder(Folder & "/include").Files
					ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
					'UPGRADE_WARNING: δ�ܽ������� objFor.Path ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					aryTemplateFile(UBound(aryTemplateFile)) = objFor.Path
					'UPGRADE_WARNING: δ�ܽ������� objFor.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					Log("�ҵ������ļ���" & objFor.Name)
				Next objFor
			End If
			If objFSO.FolderExists(Folder & "/plugin") Then
				For	Each objFor In objFSO.GetFolder(Folder & "/plugin").Files
					ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
					'UPGRADE_WARNING: δ�ܽ������� objFor.Path ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					aryPluginFile(UBound(aryPluginFile)) = objFor.Path
					'UPGRADE_WARNING: δ�ܽ������� objFor.Name ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					Log("�ҵ���������" & objFor.Name)
				Next objFor
			End If
			If objFSO.FileExists(Folder & "/source/style.css.asp") Then
				strSource = objFSO.GetFile(Folder & "/source/style.css.asp").Path
				Log("�ҵ�STYLE.CSS.ASP")
			End If
			GetSubFolder = True
		Else
			Log("�ļ��в����ڣ�")
		End If
	End Function
	
	
	'Usage:�õ�XML��Ϣ���ж��Ƿ�Z-Blog
	'Param:XMLPath--XML��ַ
	Function LoadXMLInfo(ByVal XMLPath As String) As Boolean
		
	End Function
	
	
	
	
	'Usage:����
	'Param:strFilePath--�ļ���,intType--��������
	'UPGRADE_NOTE: Update �������� Update_Renamed�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"��
	Function Update_Renamed(ByVal strFilePath As String, Optional ByRef intType As Short = 1) As Boolean
		Dim vbSpace As Object
		Dim strFile As String
		Dim objExec As Object
		If objFSO.FileExists(strFilePath) Then
			Log("Update: " & strFilePath & "  type:" & intType)
			'UPGRADE_WARNING: δ�ܽ������� LoadFromFile() ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			strFile = LoadFromFile(strFilePath)
			Select Case intType
				Case 1
					'ģ�������INCLUDE�ļ�������
					
					
					'�滻zb_system���ļ�
					objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)"
					
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.SubMatches(0), 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.SubMatches(0) & "-->" & "zb_system/" & objExec.SubMatches(0))
					Next objExec
					
					'�滻zb_users���ļ�
					objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(plugin|language|cache|upload)"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_users/" & objExec.SubMatches(0), 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.SubMatches(0) & "-->" & "zb_users/" & objExec.SubMatches(0))
					Next objExec
					
					'�滻theme
					objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>themes)"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>zb_users/theme", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>zb_users/theme")
					Next objExec
					
					'�滻rss
					objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>rss\.xml)"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>feed.asp", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.SubMatches ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>feed.asp")
					Next objExec
					
					
					'�滻��Щ����
					objRegExp.Pattern = "var (str0[0-9]|intMaxLen|strBatchView|strBatchInculde|strBatchCount|strFaceName|strFaceSize)=.+?;"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					'ǿ��c_html_js_add.asp
					If InStr(LCase(strFile), "c_html_js_add.asp") = 0 And InStr(LCase(strFile), "</head>") > 0 Then
						strFile = Replace(strFile, "</head>", "<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js_add.asp"" type=""text/javascript""></script>" & vbCrLf & "</head>")
						Log("ǿ�Ʋ���c_html_js_add.asp")
					End If
					
					'ɾ������UBB����
					objRegExp.Pattern = "InsertQuote.+?\;|ExportUbbFrame\(\)\;?"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					
					'�滻����
					strFile = Replace(strFile, "strBatchCount+=""spn<#article/id#>=<#article/id#>,""", "AddViewCount(<#article/id#>)")
					strFile = Replace(strFile, "strBatchView+=""spn<#article/id#>=<#article/id#>,""", "LoadViewCount(<#article/id#>)")
					Log("���������޸�")
					
					'�滻���ñ�ǩ
					objRegExp.Pattern = "<#template:article_trackback#>|<#article/pretrackback_url#>|<#ZC_MSG014#>|<#article/trackbacknums#>"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					'�滻Try--elScript
					objRegExp.Pattern = "try{" & vbCrLf & ".+?elScript[\d\D]+?catch\(e\){};?"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					
					'�滻��֤��
					objRegExp.Pattern = "if.+?inpVerify[\d\D]+?Math.random\(\)[\d\D]+?}[\d\D]+?}"
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					'�滻����
					'UPGRADE_WARNING: δ�ܽ������� vbSpace ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					objRegExp.Pattern = "[" & vbTab & vbSpace & "]+" & vbCrLf
					For	Each objExec In objRegExp.Execute(strFile)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						strFile = Replace(strFile, objExec.Value, "", 1, 1)
						'UPGRADE_WARNING: δ�ܽ������� objExec.Value ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
						Log(objExec.Value & "-->" & """""")
					Next objExec
					
					
					
					'����
					SaveToFile(strFilePath, strFile)
					Log("�������")
				Case 2
					'SOURCE\STYLE.CSS.ASP����
					
					'�滻<%
					strFile = Replace(strFile, "<%", "<!-- #include file=""../../../../zb_system/function/c_function.asp"" -->" & vbCrLf & "<%")
					Log("����c_function.asp")
					
					'�滻·��
					strFile = Replace(strFile, """themes""", """zb_users/theme""")
					Log("""themes"" --> ""zb_users/theme""")
					
					'�滻HOST
					strFile = Replace(strFile, "ZC_BLOG_HOST", "GetCurrentHost()")
					Log("ZC_BLOG_HOST --> GetCurrentHost()")
					
					
					SaveToFile(strFilePath, strFile)
					Log("�������")
					
				Case 3
					'���\����������
				Case 4
					'������������
					'��������ֻ����Ĭ������ĽṹŪ����Ĭ������Ľṹ������
					'��������20�����⣬Ĭ����������ṹԼռ50%����
					
					objRegExp.Pattern = "<div id=""divSidebar"">[\d\D]+?<div class=""function"""
					'�ж��Ƿ���ڽṹ��Ĭ��������ͬ�Ĳ���
					If objRegExp.Test(strFile) Then
						
						'objRegExp.Pattern = "<div id=""divSidebar"">[\d\D]+?</div>"
						
					End If
					
				Case 5
					'XML����
					
			End Select
		Else
			Log(strFile & "�Ҳ�����")
		End If
	End Function
End Class