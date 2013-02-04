<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows 窗体设计器生成的代码 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()
	End Sub
	'窗体重写释放，以清理组件列表。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows 窗体设计器所必需的
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents lstLog As System.Windows.Forms.ListBox
	Public WithEvents cmdOpen As System.Windows.Forms.Button
	Public WithEvents cmdBrowse As System.Windows.Forms.Button
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents lblNote As System.Windows.Forms.Label
	Public WithEvents lblFolder As System.Windows.Forms.Label
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.lstLog = New System.Windows.Forms.ListBox
		Me.cmdOpen = New System.Windows.Forms.Button
		Me.cmdBrowse = New System.Windows.Forms.Button
		Me.txtPath = New System.Windows.Forms.TextBox
		Me.lblNote = New System.Windows.Forms.Label
		Me.lblFolder = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.White
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "1.8 模板升级器"
		Me.ClientSize = New System.Drawing.Size(715, 457)
		Me.Location = New System.Drawing.Point(514, 330)
		Me.MaximizeBox = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmMain"
		Me.lstLog.Size = New System.Drawing.Size(681, 283)
		Me.lstLog.Location = New System.Drawing.Point(16, 56)
		Me.lstLog.TabIndex = 4
		Me.lstLog.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstLog.BackColor = System.Drawing.SystemColors.Window
		Me.lstLog.CausesValidation = True
		Me.lstLog.Enabled = True
		Me.lstLog.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstLog.IntegralHeight = True
		Me.lstLog.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstLog.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstLog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstLog.Sorted = False
		Me.lstLog.TabStop = True
		Me.lstLog.Visible = True
		Me.lstLog.MultiColumn = False
		Me.lstLog.Name = "lstLog"
		Me.cmdOpen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOpen.Text = "升级(&U)"
		Me.cmdOpen.Size = New System.Drawing.Size(65, 25)
		Me.cmdOpen.Location = New System.Drawing.Point(624, 16)
		Me.cmdOpen.TabIndex = 3
		Me.cmdOpen.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOpen.CausesValidation = True
		Me.cmdOpen.Enabled = True
		Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOpen.TabStop = True
		Me.cmdOpen.Name = "cmdOpen"
		Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowse.Text = "浏览(&B)"
		Me.cmdBrowse.Size = New System.Drawing.Size(65, 25)
		Me.cmdBrowse.Location = New System.Drawing.Point(552, 16)
		Me.cmdBrowse.TabIndex = 2
		Me.cmdBrowse.CausesValidation = True
		Me.cmdBrowse.Enabled = True
		Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowse.TabStop = True
		Me.cmdBrowse.Name = "cmdBrowse"
		Me.txtPath.AutoSize = False
		Me.txtPath.Size = New System.Drawing.Size(473, 18)
		Me.txtPath.Location = New System.Drawing.Point(72, 19)
		Me.txtPath.TabIndex = 1
		Me.txtPath.AcceptsReturn = True
		Me.txtPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtPath.CausesValidation = True
		Me.txtPath.Enabled = True
		Me.txtPath.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPath.HideSelection = True
		Me.txtPath.ReadOnly = False
		Me.txtPath.Maxlength = 0
		Me.txtPath.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPath.MultiLine = False
		Me.txtPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPath.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPath.TabStop = True
		Me.txtPath.Visible = True
		Me.txtPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPath.Name = "txtPath"
		Me.lblNote.Size = New System.Drawing.Size(681, 105)
		Me.lblNote.Location = New System.Drawing.Point(16, 344)
		Me.lblNote.TabIndex = 5
		Me.lblNote.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblNote.BackColor = System.Drawing.Color.Transparent
		Me.lblNote.Enabled = True
		Me.lblNote.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblNote.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblNote.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblNote.UseMnemonic = True
		Me.lblNote.Visible = True
		Me.lblNote.AutoSize = False
		Me.lblNote.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblNote.Name = "lblNote"
		Me.lblFolder.Text = "模板路径"
		Me.lblFolder.Size = New System.Drawing.Size(65, 17)
		Me.lblFolder.Location = New System.Drawing.Point(16, 22)
		Me.lblFolder.TabIndex = 0
		Me.lblFolder.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFolder.BackColor = System.Drawing.Color.Transparent
		Me.lblFolder.Enabled = True
		Me.lblFolder.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFolder.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFolder.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFolder.UseMnemonic = True
		Me.lblFolder.Visible = True
		Me.lblFolder.AutoSize = False
		Me.lblFolder.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFolder.Name = "lblFolder"
		Me.Controls.Add(lstLog)
		Me.Controls.Add(cmdOpen)
		Me.Controls.Add(cmdBrowse)
		Me.Controls.Add(txtPath)
		Me.Controls.Add(lblNote)
		Me.Controls.Add(lblFolder)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class