Option Strict Off
Option Explicit On
Imports System.IO
Friend Class frmMain
    Inherits System.Windows.Forms.Form

	Dim strSource, strTemplateFolder, strXMLPath As String
	Dim aryTemplateFile() As String
	Dim aryPluginFile() As String
	Dim objAero As clsAero
	
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
        Dim strTemp As String
        fbdDialog.Description = "��ѡ��ģ���ļ���"
        fbdDialog.ShowDialog()
        strTemp = fbdDialog.SelectedPath
        If Not strTemp = "" Then
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
            For i = 0 To CShort(UBound(aryTemplateFile))
                If Trim(aryTemplateFile(i)) <> "" Then Update_Renamed(aryTemplateFile(i), 1) : Update_Renamed(aryTemplateFile(i), 4)
            Next
			Log("ģ���ļ��������")
			Log("��ʼ����source��asp")
			Update_Renamed(strSource, 2)
			Log("source��asp�������")
			Log("��ʼ���������Դ����")
            For i = 0 To CShort(UBound(aryPluginFile))
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
        If Environment.OSVersion.Version.Major >= 6 Then bolAero = True

        If bolAero Then
            objAero = New clsAero
            objAero.Form = Me
            objAero.Go()
        End If




        objRegExp = New VBScript_RegExp_55.RegExp
        strTemplateFolder = ""
        objRegExp.Global = True
        objRegExp.IgnoreCase = True
        ReDim aryTemplateFile(0)
        ReDim aryPluginFile(0)
        strSource = ""
        lblNote.Text = "˵����" & vbCrLf & "����ǰ���뱸�ݡ�" & vbCrLf & "��Ҫ������1.8ģ������������Ҫ��" & vbCrLf & "      1.ģ����TEMPLATE�ļ����£���չ��Ϊhtml" & vbCrLf & "      2.HTML��ǩȫ���պ�" & vbCrLf & "      3.δ��дϵͳ�Դ���common.js" & vbCrLf & "      4.δʹ��������" & vbCrLf & "��������������һ�㲻���ϣ��򱾳����޷�����������⡣"
    End Sub
	
	
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objRegExp = Nothing
    End Sub
	
	
	
	
	'Usage:��־
	'Param:str--��־����
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
        Dim objFor As Object
        If Directory.Exists(Folder) Then
            If File.Exists(Folder & "\theme.xml") Then
                strXMLPath = Folder & "\theme.xml"
                Log("�ҵ�����XML��Ϣ")
            Else
                Log("����XML������")
                Exit Function
            End If
            If Directory.Exists(Folder & "\template") Then
                For Each objFor In Directory.GetFiles(Folder & "\template")

                    '˳������ɾ����

                    If objRegExp.Test(objFor) Then
                        Log("ɾ�������ļ���" & objFor)
                        File.Delete(objFor)
                    Else
                        ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
                        '����pageģ��
                        If objFor Like "single*" Then
                            If Not File.Exists(Folder & "\template\page.html") Then
                                File.Copy(objFor, Folder & "\template\page.html") : Log("����PAGEģ��")
                            End If
                        End If
                        aryTemplateFile(UBound(aryTemplateFile)) = objFor
                        Log("�ҵ������ļ���" & objFor)
                    End If
                Next objFor
            End If
            If Directory.Exists(Folder & "\include") Then
                For Each objFor In Directory.GetFiles(Folder & "\include")
                    ReDim Preserve aryTemplateFile(UBound(aryTemplateFile) + 1)
                    aryTemplateFile(UBound(aryTemplateFile)) = objFor
                    Log("�ҵ������ļ���" & objFor)
                Next objFor
            End If
            If Directory.Exists(Folder & "\plugin") Then
                For Each objFor In Directory.GetFiles(Folder & "\plugin")
                    ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
                    aryPluginFile(UBound(aryPluginFile)) = objFor
                    Log("�ҵ���������" & objFor)
                Next objFor
            End If
            If Directory.Exists(Folder & "\source\style.css.asp") Then
                strSource = Folder & "\source\style.css.asp"
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
    Function Update_Renamed(ByVal strFilePath As String, Optional ByRef intType As Short = 1) As Boolean
        Dim strFile As String = ""
        Dim objExec As Object = Nothing
        If File.Exists(strFilePath) Then

            Log("Update: " & strFilePath & "  type:" & intType)
            strFile = File.ReadAllText(strFilePath)
            Select Case intType
                Case 1
                    'ģ�������INCLUDE�ļ�������


                    '�滻zb_system���ļ�
                    objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)"

                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.SubMatches(0), 1, 1)
                        Log(objExec.SubMatches(0) & "-->" & "zb_system/" & objExec.SubMatches(0))
                    Next objExec

                    '�滻zb_users���ļ�
                    objRegExp.Pattern = "\<\#ZC_BLOG_HOST\#\>(plugin|language|cache|upload)"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_users/" & objExec.SubMatches(0), 1, 1)
                        Log(objExec.SubMatches(0) & "-->" & "zb_users/" & objExec.SubMatches(0))
                    Next objExec

                    '�滻theme
                    objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>themes)"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>zb_users/theme", 1, 1)
                        Log(objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>zb_users/theme")
                    Next objExec

                    '�滻rss
                    objRegExp.Pattern = "(\<\#ZC_BLOG_HOST\#\>rss\.xml)"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.SubMatches(0), "<#ZC_BLOG_HOST#>feed.asp", 1, 1)
                        Log(objExec.SubMatches(0) & "-->" & "<#ZC_BLOG_HOST#>feed.asp")
                    Next objExec


                    '�滻��Щ����
                    objRegExp.Pattern = "var (str0[0-9]|intMaxLen|strBatchView|strBatchInculde|strBatchCount|strFaceName|strFaceSize)=.+?;"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec

                    'ǿ��c_html_js_add.asp
                    If InStr(LCase(strFile), "c_html_js_add.asp") = 0 And InStr(LCase(strFile), "</head>") > 0 Then
                        strFile = Replace(strFile, "</head>", "<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js_add.asp"" type=""text/javascript""></script>" & vbCrLf & "</head>")
                        Log("ǿ�Ʋ���c_html_js_add.asp")
                    End If

                    'ɾ������UBB����
                    objRegExp.Pattern = "InsertQuote.+?\;|ExportUbbFrame\(\)\;?"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec


                    '�滻����
                    strFile = Replace(strFile, "strBatchCount+=""spn<#article/id#>=<#article/id#>,""", "AddViewCount(<#article/id#>)")
                    strFile = Replace(strFile, "strBatchView+=""spn<#article/id#>=<#article/id#>,""", "LoadViewCount(<#article/id#>)")
                    Log("���������޸�")

                    '�滻���ñ�ǩ
                    objRegExp.Pattern = "<#template:article_trackback#>|<#article/pretrackback_url#>|<#ZC_MSG014#>|<#article/trackbacknums#>"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec

                    '�滻Try--elScript
                    objRegExp.Pattern = "try{" & vbCrLf & ".+?elScript[\d\D]+?catch\(e\){};?"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec


                    '�滻��֤��
                    objRegExp.Pattern = "if.+?inpVerify[\d\D]+?Math.random\(\)[\d\D]+?}[\d\D]+?}"
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec

                    '�滻����
                    objRegExp.Pattern = "[" & vbTab & " ]+" & vbCrLf
                    For Each objExec In objRegExp.Execute(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec



                    '����
                    File.WriteAllText(strFilePath, strFile)
                    Log("�������")
                Case 2
                    'SOURCE\STYLE.CSS.ASP����

                    '�滻<%
                    strFile = Replace(strFile, "<%", "<!-- #include file=""..\..\..\..\zb_system\function\c_function.asp"" -->" & vbCrLf & "<%")
                    Log("����c_function.asp")

                    '�滻·��
                    strFile = Replace(strFile, """themes""", """zb_users/theme""")
                    Log("""themes"" --> ""zb_users/theme""")

                    '�滻HOST
                    strFile = Replace(strFile, "ZC_BLOG_HOST", "GetCurrentHost()")
                    Log("ZC_BLOG_HOST --> GetCurrentHost()")


                    File.WriteAllText(strFilePath, strFile)
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