<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdatePlugin
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdatePlugin))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lstLog = New System.Windows.Forms.ListBox()
        Me.cmdOpen = New System.Windows.Forms.Button()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.lblFolder = New System.Windows.Forms.Label()
        Me.fbdDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.SuspendLayout()
        '
        'lstLog
        '
        Me.lstLog.BackColor = System.Drawing.SystemColors.Window
        Me.lstLog.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstLog.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstLog.ItemHeight = 12
        Me.lstLog.Location = New System.Drawing.Point(14, 52)
        Me.lstLog.Name = "lstLog"
        Me.lstLog.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstLog.Size = New System.Drawing.Size(681, 280)
        Me.lstLog.TabIndex = 10
        '
        'cmdOpen
        '
        Me.cmdOpen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpen.Location = New System.Drawing.Point(622, 12)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpen.Size = New System.Drawing.Size(65, 25)
        Me.cmdOpen.TabIndex = 9
        Me.cmdOpen.Text = "升级(&U)"
        Me.cmdOpen.UseVisualStyleBackColor = False
        '
        'cmdBrowse
        '
        Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.Location = New System.Drawing.Point(550, 12)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowse.Size = New System.Drawing.Size(65, 25)
        Me.cmdBrowse.TabIndex = 8
        Me.cmdBrowse.Text = "浏览(&B)"
        Me.cmdBrowse.UseVisualStyleBackColor = False
        '
        'txtPath
        '
        Me.txtPath.AcceptsReturn = True
        Me.txtPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtPath.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPath.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPath.Location = New System.Drawing.Point(70, 15)
        Me.txtPath.MaxLength = 0
        Me.txtPath.Name = "txtPath"
        Me.txtPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPath.Size = New System.Drawing.Size(473, 21)
        Me.txtPath.TabIndex = 7
        '
        'lblNote
        '
        Me.lblNote.BackColor = System.Drawing.Color.Transparent
        Me.lblNote.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNote.Location = New System.Drawing.Point(14, 340)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNote.Size = New System.Drawing.Size(681, 105)
        Me.lblNote.TabIndex = 11
        '
        'lblFolder
        '
        Me.lblFolder.BackColor = System.Drawing.Color.Transparent
        Me.lblFolder.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFolder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFolder.Location = New System.Drawing.Point(14, 18)
        Me.lblFolder.Name = "lblFolder"
        Me.lblFolder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFolder.Size = New System.Drawing.Size(65, 17)
        Me.lblFolder.TabIndex = 6
        Me.lblFolder.Text = "模板路径"
        '
        'frmUpdatePlugin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(703, 453)
        Me.Controls.Add(Me.lstLog)
        Me.Controls.Add(Me.cmdOpen)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.lblNote)
        Me.Controls.Add(Me.lblFolder)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmUpdatePlugin"
        Me.Text = "1.8插件升级器"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lstLog As System.Windows.Forms.ListBox
    Public WithEvents cmdOpen As System.Windows.Forms.Button
    Public WithEvents cmdBrowse As System.Windows.Forms.Button
    Public WithEvents txtPath As System.Windows.Forms.TextBox
    Public WithEvents lblNote As System.Windows.Forms.Label
    Public WithEvents lblFolder As System.Windows.Forms.Label
    Friend WithEvents fbdDialog As System.Windows.Forms.FolderBrowserDialog
End Class
