Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text.RegularExpressions

Friend Class frmUpdateTheme
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
                If Trim(aryTemplateFile(i)) <> "" Then Update_Theme(aryTemplateFile(i), 1) : Update_Theme(aryTemplateFile(i), 4)
            Next
            Log("ģ���ļ��������")
            Log("��ʼ����source��asp")
            Update_Theme(strSource, 2)
            Log("source��asp�������")
                Log("��ʼ���������Դ����")

            If Directory.Exists(strTemplateFolder & "\plugin") Then

                frmUpdatePlugin.Show()
                frmUpdatePlugin.Enabled = True
                frmUpdatePlugin.cmdOpen.Enabled = False
                frmUpdatePlugin.cmdBrowse.Enabled = False
                frmUpdatePlugin.txtPath.Enabled = False
                frmUpdatePlugin.strTemplateFolder = strTemplateFolder & "\plugin"

                frmUpdatePlugin.GetSubFolder(strTemplateFolder & "\plugin", "..\..\..\", "..\..\..\..\", True)
                frmUpdatePlugin.Update_P(True)
                'frmUpdatePlugin.Hide()
            End If

                Log("�����Դ�����������")

            '����������������
            'Log "����XML��Ϣ"
            '����XML�ǲ�����APP������һ��

            MsgBox(Replace("������ϣ�\n\nʣ�����²���û���������������޸ģ�\n\n�������֣������2.0�����淶��\n������\nXML��Ϣ\n\n������ɺ�����APP������༭������Ϣ�����棬������2.0�Ｄ�����⡣", "\n", vbCrLf), MsgBoxStyle.Information)
            Process.Start("http://www.zsxsoft.com/updatesuccess.html")
        End If

    End Sub

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If Environment.OSVersion.Version.Major >= 6 Then bolAero = True

        If bolAero Then
            objAero = New clsAero
            objAero.Form = Me
            objAero.Go()
        End If




        strTemplateFolder = ""
        ReDim aryTemplateFile(0)
        ReDim aryPluginFile(0)
        strSource = ""
        lblNote.Text = "˵����" & vbCrLf & "����ǰ���뱸�ݡ�" & vbCrLf & "��Ҫ������1.8ģ������������Ҫ��" & vbCrLf & "      1.ģ����TEMPLATE�ļ����£���չ��Ϊhtml" & vbCrLf & "      2.HTML��ǩȫ���պ�" & vbCrLf & "      3.δ��дϵͳ�Դ���common.js" & vbCrLf & "��������������һ�㲻���ϣ��򱾳����޷�����������⡣"
    End Sub


    Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
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
        Dim objRegExp As New Regex("b_article-guestbook|b_article_trackback|guestbook|search", RegexOptions.IgnoreCase)
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

                    If objRegExp.Match(objFor).Success Then
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
    Function Update_Theme(ByVal strFilePath As String, Optional ByRef intType As Short = 1) As Boolean
        Dim strFile As String = ""
        Dim objExec As Object = Nothing
        If File.Exists(strFilePath) Then

            Log("Update: " & strFilePath & "  type:" & intType)
            strFile = File.ReadAllText(strFilePath)
            If Trim(strFile) = "" Then Return False
            Select Case intType
                Case 1
                    'ģ�������INCLUDE�ļ�������


                    '�滻zb_system���ļ�

                    For Each objExec In New Regex("\<\#ZC_BLOG_HOST\#\>(admin|script|function|image|cmd.asp|login.asp)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_system/" & objExec.Groups(1).Value, 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & "zb_system/" & objExec.Groups(1).Value)
                    Next objExec

                    '�滻zb_users���ļ�
                    For Each objExec In New Regex("\<\#ZC_BLOG_HOST\#\>(plugin|language|cache|upload)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "<#ZC_BLOG_HOST#>zb_users/" & objExec.Groups(1).Value, 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & "zb_users/" & objExec.Groups(1).Value)
                    Next objExec

                    '�滻theme
                    For Each objExec In New Regex("(\<\#ZC_BLOG_HOST\#\>themes)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Groups(1).Value, "<#ZC_BLOG_HOST#>zb_users/theme", 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & "<#ZC_BLOG_HOST#>zb_users/theme")
                    Next objExec

                    '�滻rss
                    For Each objExec In New Regex("(\<\#ZC_BLOG_HOST\#\>rss\.xml)", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Groups(1).Value, "<#ZC_BLOG_HOST#>feed.asp", 1, 1)
                        Log(objExec.Groups(1).Value & "-->" & "<#ZC_BLOG_HOST#>feed.asp")
                    Next objExec


                    '�滻��Щ����
                    For Each objExec In New Regex("var (str0[0-9]|intMaxLen|strBatchView|strBatchInculde|strBatchCount|strFaceName|strFaceSize)=.+?;", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec

                    'ǿ��c_html_js_add.asp
                    If InStr(LCase(strFile), "c_html_js_add.asp") = 0 And InStr(LCase(strFile), "</head>") > 0 Then
                        strFile = Replace(strFile, "</head>", "<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js_add.asp"" type=""text/javascript""></script>" & vbCrLf & "</head>")
                        Log("ǿ�Ʋ���c_html_js_add.asp")
                    End If

                    'ɾ������UBB����
                    For Each objExec In New Regex("InsertQuote.+?\;|ExportUbbFrame\(\)\;?", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec


                    '�滻����
                    strFile = Replace(strFile, "strBatchCount+=""spn<#article/id#>=<#article/id#>,""", "AddViewCount(<#article/id#>)")
                    strFile = Replace(strFile, "strBatchView+=""spn<#article/id#>=<#article/id#>,""", "LoadViewCount(<#article/id#>)")
                    Log("���������޸�")

                    '�滻���ñ�ǩ
                    For Each objExec In New Regex("<#template:article_trackback#>|<#article/pretrackback_url#>|<#ZC_MSG014#>|<#article/trackbacknums#>", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec

                    '�滻Try--elScript
                    For Each objExec In New Regex("try{" & vbCrLf & ".+?elScript[\d\D]+?catch\(e\){};?", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec


                    '�滻��֤��
                    For Each objExec In New Regex("if.+?inpVerify[\d\D]+?Math.random\(\)[\d\D]+?}[\d\D]+?}", RegexOptions.IgnoreCase).Matches(strFile)
                        strFile = Replace(strFile, objExec.Value, "", 1, 1)
                        Log(objExec.Value & "-->" & """""")
                    Next objExec

                    '�滻����
                    For Each objExec In New Regex("[" & vbTab & " ]+" & vbCrLf).Matches(strFile)
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

                    Dim objRegExp As New Regex("<div id=""divSidebar"">[\d\D]+?<div class=""function""", RegexOptions.IgnoreCase)
                    '�ж��Ƿ���ڽṹ��Ĭ��������ͬ�Ĳ���
                    If objRegExp.Match(strFile).Success Then

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