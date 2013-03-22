<%
Dim WATERMARK_WIDTH_POSITION,WATERMARK_HEIGHT_POSITION,WATERMARK_QUALITY,WATERMARK_FONTCOLOR,WATERMARK_FONTBOLD,WATERMARK_FONTSIZE,WATERMARK_FONTQUALITY,WATERMARK_TYPE,WATERMARK_TEXT,WATERMARK_LOGO,WATERMARK_ALPHA

'注册插件
Call RegisterPlugin("Watermark","ActivePlugin_Watermark")

Dim Watermark_Config
Function Watermark_Initialize()
	Set Watermark_Config=New TConfig
	Watermark_Config.Load "Watermark"
	If Watermark_Config.Exists("VERSION")=False Then
		Watermark_Config.Write "VERSION","1.0"
		Watermark_Config.Write "WIDTH_POSITION","right"
		Watermark_Config.Write "HEIGHT_POSITION","bottom"
		Watermark_Config.Write "QUALITY",80
		Watermark_Config.Write "FONTCOLOR","#000"
		Watermark_Config.Write "FONTBOLD","True"
		Watermark_Config.Write "FONTSIZE","14"
		Watermark_Config.Write "FONTQUALITY","4"
		Watermark_Config.Write "TYPE","1"
		Watermark_Config.Write "TEXT","我是文字水印"
		Watermark_Config.Write "LOGO","test.jpg"
		Watermark_Config.Write "ALPHA","0.7"
		Watermark_Config.Save
	End If
	WATERMARK_WIDTH_POSITION = Watermark_Config.Read("WIDTH_POSITION")
	WATERMARK_HEIGHT_POSITION = Watermark_Config.Read("HEIGHT_POSITION")
	WATERMARK_QUALITY = CInt(Watermark_Config.Read("QUALITY"))
	WATERMARK_FONTCOLOR = Watermark_Config.Read("FONTCOLOR")
	WATERMARK_FONTBOLD = CBool(Watermark_Config.Read("FONTBOLD"))
	WATERMARK_FONTSIZE = Watermark_Config.Read("FONTSIZE")
	WATERMARK_FONTQUALITY = Watermark_Config.Read("FONTQUALITY")
	WATERMARK_TYPE = CInt(Watermark_Config.Read("TYPE"))
	WATERMARK_TEXT = Watermark_Config.Read("TEXT")
	WATERMARK_LOGO = Watermark_Config.Read("LOGO")
	WATERMARK_ALPHA = Watermark_Config.Read("ALPHA")
End Function

'挂口部分
Function ActivePlugin_Watermark()

	Call Add_Action_Plugin("Action_Plugin_uEditor_FileUpload_End","Watermark_uEditorUpload("")")

End Function


Function InstallPlugin_Watermark()


End Function


Function UnInstallPlugin_Watermark()


End Function

'重写这个函数，Call水印处理
Function UploadFile(bolAutoName)

	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile

	objUpLoadFile.AuthorID=BlogUser.ID
	objUpLoadFile.AutoName=bolAutoName

	If objUpLoadFile.UpLoad() Then

		Call Watermark_uEditorUpload(objUpLoadFile.FullPath)

		UploadFile=True

	End If

	Set objUpLoadFile=Nothing

End Function

Function Watermark_uEditorUpload(url)
	'On Error Resume Next
	Call Watermark_Initialize

	If url = "" Then url = BlogPath & strUPLOADDIR &"\" & objUpload.form(uEditor_ASPFormName)

	If Instr(LCase(url),"jpg") = 0 Then Exit Function

	Dim Jpeg,Logo,LogoPath,TextWidth,PositionWidth,PositionHeight
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	If Jpeg.Version >= "1.9" Then Jpeg.PreserveMetadata = True
	Jpeg.Open url
	Jpeg.Quality = WATERMARK_QUALITY
	If WATERMARK_TYPE = 1 Then
		Jpeg.Canvas.Font.Color = Replace(WATERMARK_FONTCOLOR, "#", "&h") '字体颜色
		Jpeg.Canvas.Font.Family = "Tahoma" 'family设置字体
		Jpeg.Canvas.Font.Bold = WATERMARK_FONTBOLD '是否设置成粗体
		Jpeg.Canvas.Font.Size = WATERMARK_FONTSIZE '字体大小
		Jpeg.Canvas.Font.Quality = WATERMARK_FONTQUALITY ' 输出文字质量
		TextWidth = Jpeg.Canvas.GetTextExtent(WATERMARK_TEXT)
		Select Case WATERMARK_WIDTH_POSITION
			Case "left"
				PositionWidth = 10
			Case "center"
				PositionWidth = (Jpeg.Width - TextWidth) / 2
			Case "right"
				PositionWidth = Jpeg.Width - TextWidth - 10
		End Select
		Select Case WATERMARK_HEIGHT_POSITION
			Case "top"
				PositionHeight = 10
			Case "center"
				PositionHeight = (Jpeg.Height - 12) / 2
			Case "bottom"
				PositionHeight = Jpeg.Height - 12 - 10
		End Select
		Jpeg.Canvas.Print PositionWidth, PositionHeight, WATERMARK_TEXT
		Jpeg.Save url
	Else
		Set Logo = Server.CreateObject("Persits.Jpeg")
		'LogoPath = BlogPath & "zb_users\PLGUIN\Watermark\" & WATERMARK_LOGO
		LogoPath = "G:\z2test\zb_users\PLUGIN\Watermark\test.jpg"
		Logo.Open LogoPath
		Select Case WATERMARK_WIDTH_POSITION
			Case "left"
				PositionWidth = 10
			Case "center"
				PositionWidth = (Jpeg.Width - Logo.Width) / 2
			Case "right"
				PositionWidth = Jpeg.Width - Logo.Width - 10
		End Select
		Select Case WATERMARK_HEIGHT_POSITION
			Case "top"
				PositionHeight = 10
			Case "center"
				PositionHeight = (Jpeg.Height - Logo.Height) / 2
			Case "bottom"
				PositionHeight = Jpeg.Height - Logo.Height - 10
		End Select
		If Instr(WATERMARK_LOGO,"png") Or Instr(WATERMARK_LOGO,"gif") And Jpeg.Version >= "1.8" Then
			Jpeg.Canvas.DrawPNG PositionWidth, PositionHeight, LogoPath
		Else
			Jpeg.Canvas.DrawImage PositionWidth, PositionHeight, Logo, WATERMARK_ALPHA
		End If
		Jpeg.Save url
		Set Logo = Nothing
	End If
	Set Jpeg = Nothing

End Function
%>