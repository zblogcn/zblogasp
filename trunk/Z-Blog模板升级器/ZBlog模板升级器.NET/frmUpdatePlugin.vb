Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text.RegularExpressions

Friend Class frmUpdatePlugin
    Inherits System.Windows.Forms.Form

    Public strSource, strTemplateFolder, strXMLPath As String
    Public aryPluginFile() As String
    Public aryOldPath() As String
    Public aryNewPath() As String
    Dim objAero As clsAero


    Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
        Dim strTemp As String
        fbdDialog.Description = "请选择插件文件夹"
        fbdDialog.ShowDialog()
        strTemp = fbdDialog.SelectedPath
        If Not strTemp = "" Then
            strTemplateFolder = strTemp
            txtPath.Text = strTemp
            Log("选择插件文件夹：" & strTemp)
        End If
    End Sub

    Public Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
        Update_P(False)
        MsgBox(Replace("升级完毕！\n\n剩余以下部分没有升级，请自行修改：\n\n升级完成后，请在APP中心里编辑插件信息并保存，即可在2.0里激活插件。", "\n", vbCrLf), MsgBoxStyle.Information)
        Process.Start("http://www.zsxsoft.com/updatesuccess.html")

    End Sub

    Sub Update_P(ByVal otherForm As Boolean)
        Dim i As Short
        Log_Clear()
        If Not otherForm Then
            strTemplateFolder = txtPath.Text
            If Not GetSubFolder(strTemplateFolder) Then
                Return
            End If
        End If

        Log("开始升级插件文件")
        For i = 0 To CShort(UBound(aryPluginFile))
            If Trim(aryPluginFile(i)) <> "" Then
                Update_Plugin(aryPluginFile(i), 1, aryOldPath(i), aryNewPath(i))
                Update_Plugin(aryPluginFile(i), 2, aryOldPath(i), aryNewPath(i))
            End If

        Next

    End Sub

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If Environment.OSVersion.Version.Major >= 6 Then bolAero = True

        If bolAero Then
            objAero = New clsAero
            objAero.Form = Me
            objAero.Go()
        End If


        strTemplateFolder = ""
        ReDim aryPluginFile(0)
        ReDim aryOldPath(0)
        ReDim aryNewPath(0)
        strSource = ""
        lblNote.Text = "说明：" & vbCrLf & "升级前必须备份。" & vbCrLf & _
            "您要升级的1.8插件必须符合以下要求：" & vbCrLf & _
            "以上条件有任意一点不符合，则本程序无法升级你的插件。"
    End Sub


    Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub


    'Usage:清除日志
    Sub Log_Clear()
        lstLog.Items.Clear()
    End Sub

    'Usage:扫描文件夹
    'Param:Folder--文件夹
    Function GetSubFolder(ByVal Folder As String, Optional ByVal OldPath As String = "..\..\", Optional ByVal NewPath As String = "..\..\..\", Optional ByVal SubFolder As Boolean = False) As Boolean
        GetSubFolder = False
        Dim objFor As Object
        If Directory.Exists(Folder) Then
            If Not SubFolder Then
                If File.Exists(Folder & "\plugin.xml") Then
                    strXMLPath = Folder & "\plugin.xml"
                    Log("找到插件XML信息")
                Else
                    Log("插件XML不存在")
                    Exit Function
                End If
            End If
            For Each objFor In Directory.GetFiles(Folder)
                If objFor Like "*.asp" Then
                    ReDim Preserve aryPluginFile(UBound(aryPluginFile) + 1)
                    ReDim Preserve aryOldPath(UBound(aryOldPath) + 1)
                    ReDim Preserve aryNewPath(UBound(aryNewPath) + 1)
                    aryPluginFile(UBound(aryPluginFile)) = objFor
                    aryOldPath(UBound(aryOldPath)) = OldPath
                    aryNewPath(UBound(aryNewPath)) = NewPath
                    Log("找到插件文件：" & objFor)
                End If
            Next objFor

            For Each objFor In Directory.GetDirectories(Folder)
                GetSubFolder(objFor, OldPath & "..\", NewPath & "..\", True)
            Next objFor

            GetSubFolder = True
        Else
            Log("文件夹不存在！")
        End If
    End Function


    'Usage:得到XML信息以判断是否Z-Blog
    'Param:XMLPath--XML地址
    Function LoadXMLInfo(ByVal XMLPath As String) As Boolean

    End Function





End Class