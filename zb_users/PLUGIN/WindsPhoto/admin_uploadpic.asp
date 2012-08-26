<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备    注:    WindsPhoto
'// 最后修改：   2011.8.22
'// 最后版本:    2.7.3
'///////////////////////////////////////////////////////////////////////////////
%>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="data/conn.asp" -->
<!-- #include file="function.asp"-->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)
%>
<%
Dim upload, File, formName, formPath, iCount, FileName, FileExt, i
Set upload = New upload_5xSoft '建立上传对象

name = upload.Form("name")
url = upload.Form("url")
surl = upload.Form("surl")
zhuanti = upload.Form("zhuanti")
mark = upload.Form("mark")
autoname = upload.Form("autoname")
photointro = upload.Form("photointro")
category = upload.Form("category")
quick = upload.Form("quick")
hot = 0
itime=now()

If url<>"" Then

    If InStr(url, "http") = 0 Or InStr(url, "http") = Null Then

        Call SetBlogHint_Custom("? 远程图片的话,你只能添加http开头的图片地址.</a>")
        Response.Redirect "admin_addphoto.asp?typeid=" & zhuanti

    End If

    photourlb = url

    If surl<>"" Then
        photourls = surl
    Else
        If instr(url,"ggpht.com") then   'Picasa转换缩略图
            photourls = Replace(url, "s800", "s144")
        elseif instr(url,"flickr.com") and instr(url,"_o") = false then     'flickr转换缩略图
            photourls = Replace(url, ".jpg", "") & "_m.jpg"
        elseif instr(url,"bababian.com") and instr(url,"_") = false then  '巴巴变转换缩略图
            photourls = Replace(url, ".jpg", "") & "_240.jpg"
        else
            photourls = url
        end if
    End If

    strSQL = "insert into desktop ([name],[itime],zhuanti,jj,url,surl,hot) values ('"&Name&"','"&itime&"',"&zhuanti&",'"&photointro&"','"&photourlb&"','"&photourls&"','"&hot&"')"
    conn.Execute strSQL
    Call SetBlogHint_Custom("√ 添加远程图片成功.</a>")
    Response.Redirect "admin_addphoto.asp?typeid=" & zhuanti

Else

    FilePath = Server.MapPath("./") '设置上传目录位置
    If WP_UPLOAD_DIRBY = 1 Then
        CreatDirectoryByCustomDirectory("plugin/windsphoto/" & WP_UPLOAD_DIR & "/" &Year(GetTime(Now()))&Month(GetTime(Now())))
        FilePath = FilePath & "/" & WP_UPLOAD_DIR & "/" &Year(GetTime(Now()))&Month(GetTime(Now())) & "/"
    ElseIf WP_UPLOAD_DIRBY = 2 Then
        CreatDirectoryByCustomDirectory("plugin/windsphoto/" & WP_UPLOAD_DIR & "/" & zhuanti)
        FilePath = FilePath & "/" & WP_UPLOAD_DIR & "/" & zhuanti & "/"
    Else
        CreatDirectoryByCustomDirectory("plugin/windsphoto/" & WP_UPLOAD_DIR)
        FilePath = FilePath & "/" & WP_UPLOAD_DIR & "/"
    End If

	For Each formName in upload.objFile '列出所有上传了的文件
		Set File = upload.File(formName) '生成一个文件对象

		FileExt = LCase(Right(File.FileName, 4))
		If FileExt<>".gif" And FileExt<>".jpg" And FileExt<>".jpeg" And FileExt<>".png" And FileExt<>".bmp" Then
			Response.Redirect "admin_addphoto.asp?typeid=" & zhuanti
			Response.End
		End If

		If File.filesize>WP_UPLOAD_FILESIZE Then
			Call SetBlogHint_Custom("!! 图片文件太大,无法上传.你可以在修改<a href='admin_setting.asp'>上传限制的最大字节数</a>.")
			Response.Redirect "admin_addphoto.asp?typeid=" & zhuanti
			Response.End
		End If

		'变更jpeg和bmp格式为jpg格式
		If FileExt="jpeg" or FileExt=".bmp" then FileExt = ".jpg"
		Randomize
		RanNum = Int((99 -10 + 1) * Rnd + 99)
		If autoname = "1" Then
			FileNamet = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&RanNum&FileExt
		Else
			FileNamet = File.FileName
		End If
		FileName = FilePath&FileNamet

		If File.FileSize>0 Then '如果 FileSize > 0 说明有文件数据
			File.SaveAs FileName '保存文件
		Else
			'Call SetBlogHint_Custom("!! 文件上传错误.")
			Response.Redirect "admin_addphoto.asp?typeid=" & zhuanti
			Response.End
		End If

		'ASPJPEG处理
		If WP_IF_ASPJPEG="1" Then

			Dim Jpeg
            Set Jpeg = Server.CreateObject("Persits.Jpeg")

            '如果aspjpeg版本大于1.9，启用保护Metadata
            If Jpeg.Version>= "1.9" then Jpeg.PreserveMetadata = True

            Jpeg.Open(FileName)

            '变更缩略图文件扩展名为jpg
            FileNamelen = Len(FileNamet) - 4
            FileNamet1 = FileNamet
            FileNamet = Left(FileNamet, FileNamelen) &".jpg"

            '缩略图处理，判断哪边为长边，以长边进行缩放
            imgWidth = Jpeg.OriginalWidth
            imgHeight = Jpeg.OriginalHeight

            If imgWidth>= imgHeight And imgWidth>WP_SMALL_WIDTH Then
                Jpeg.Width = WP_SMALL_WIDTH
                Jpeg.Height = Jpeg.OriginalHeight / (Jpeg.OriginalWidth / WP_SMALL_WIDTH)
            End If
            If imgHeight>imgWidth And imgHeight>WP_SMALL_HEIGHT Then
                Jpeg.Height = WP_SMALL_HEIGHT
                Jpeg.Width = Jpeg.OriginalWidth / (Jpeg.OriginalHeight / WP_SMALL_HEIGHT)
            End If

            '保存缩略图，并进行微度锐化
            Jpeg.Sharpen 1, 110
            Jpeg.Save (FilePath & "small_" & FileNamet)

            '水印处理
			If mark<>"" Then

                If WP_WATERMARK_TYPE = "1" Then '图片水印
                    If Jpeg.Version>= "1.9" then Jpeg.PreserveMetadata = True
                    Jpeg.Open FileName
                    Jpeg.Canvas.Font.Color = Replace(WP_JPEG_FONTCOLOR, "#", "&h") '字体颜色
                    Jpeg.Canvas.Font.Family = "Tahoma" 'family设置字体
                    Jpeg.Canvas.Font.Bold = WP_JPEG_FONTBOLD '是否设置成粗体
                    Jpeg.Canvas.Font.Size = WP_JPEG_FONTSIZE '字体大小
                    Jpeg.Canvas.Font.Quality = WP_JPEG_FONTQUALITY ' 输出文字质量
                    Title = WP_WATERMARK_TEXT
                    TitleWidth = Jpeg.Canvas.GetTextExtent(Title)
                    Select Case WP_WATERMARK_WIDTH_POSITION
                        Case "left"
                            PositionWidth = 10
                        Case "center"
                            PositionWidth = (Jpeg.Width - TitleWidth) / 2
                        Case "right"
                            PositionWidth = Jpeg.Width - TitleWidth - 10
                    End Select
                    Select Case WP_WATERMARK_HEIGHT_POSITION
                        Case "top"
                            PositionHeight = 10
                        Case "center"
                            PositionHeight = (Jpeg.Height - 12) / 2
                        Case "bottom"
                            PositionHeight = Jpeg.Height - 12 - 10
                    End Select
                    Jpeg.Canvas.Print PositionWidth, PositionHeight, WP_WATERMARK_TEXT
                    Jpeg.Save FileName

                ElseIf WP_WATERMARK_TYPE = "2" Then

                    Dim Jpeg1
                    Set Jpeg1 = Server.CreateObject("Persits.Jpeg")
                    Jpeg.PreserveMetadata = True
                    Jpeg.Open FileName
                    Jpeg1.Open Server.MapPath(""& WP_WATERMARK_LOGO &"")
                    Select Case WP_WATERMARK_WIDTH_POSITION
                        Case "left"
                            PositionWidth = 10
                        Case "center"
                            PositionWidth = (Jpeg.Width - Jpeg1.Width) / 2
                        Case "right"
                            PositionWidth = Jpeg.Width - Jpeg1.Width - 10
                    End Select
                    Select Case WP_WATERMARK_HEIGHT_POSITION
                        Case "top"
                            PositionHeight = 10
                        Case "center"
                            PositionHeight = (Jpeg.Height - Jpeg1.Height) / 2
                        Case "bottom"
                            PositionHeight = Jpeg.Height - Jpeg1.Height - 10
                    End Select
                    Jpeg.DrawImage PositionWidth, PositionHeight, Jpeg1, WP_WATERMARK_ALPHA, &HFFFFFF
                    Jpeg.Save FileName
                    Set Jpeg1 = Nothing
                End If

            End If
            Set Jpeg = Nothing

			'带缩略图的URL路径生成
			If WP_UPLOAD_DIRBY = 1 Then
				photourlb = WP_UPLOAD_DIR & "/" & Year(GetTime(Now()))&Month(GetTime(Now())) & "/" & FileNamet1
				photourls = WP_UPLOAD_DIR & "/" & Year(GetTime(Now()))&Month(GetTime(Now())) & "/small_" & FileNamet
			ElseIf WP_UPLOAD_DIRBY = 2 Then
				photourlb = WP_UPLOAD_DIR & "/" & zhuanti & "/" & FileNamet1
				photourls = WP_UPLOAD_DIR & "/" & zhuanti & "/small_" & FileNamet
			Else
				photourlb = WP_UPLOAD_DIR & "/" & FileNamet1
				photourls = WP_UPLOAD_DIR & "/small_" & FileNamet
			End If

		Else

			'不带缩略图的URL路径生成
			If WP_UPLOAD_DIRBY = 1 Then
				photourlb = WP_UPLOAD_DIR & "/" & Year(GetTime(Now()))&Month(GetTime(Now())) & "/" & FileNamet
				photourls = WP_UPLOAD_DIR & "/" & Year(GetTime(Now()))&Month(GetTime(Now())) & "/" & FileNamet
			ElseIf WP_UPLOAD_DIRBY = 2 Then
				photourlb = WP_UPLOAD_DIR & "/" & zhuanti & "/" & FileNamet
				photourls = WP_UPLOAD_DIR & "/" & zhuanti & "/" & FileNamet
			Else
				photourlb = WP_UPLOAD_DIR & "/" & FileNamet
				photourls = WP_UPLOAD_DIR & "/" & FileNamet
			End If

		End If

		'获取文件名作为标题
		If upload.Form("name")<>"" Then
			name = upload.Form("name")
		Else
			name = Replace(File.FileName, FileExt, "")
		End If

		'写入数据库
		strSQL = "insert into desktop ([name],[itime],zhuanti,jj,url,surl,hot) values ('"&name&"','"&itime&"',"&zhuanti&",'"&photointro&"','"&photourlb&"','"&photourls&"','"&hot&"')"
		conn.Execute strSQL
		iCount = iCount + 1
		Set File = Nothing
	Next
	Set upload = Nothing

	'处理快速上传等
	If quick = 1 then

		strFileName = "[IMG]"& ZC_BLOG_HOST &"plugin/windsphoto/" & photourlb & "[/IMG]"
		strFileName1 = ZC_BLOG_HOST &"plugin/windsphoto/" & photourlb

		Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/><meta http-equiv=""Content-Language"" content=""zh-cn"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""../../CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"
		Response.Write "<form border=""1"" name=""edit"" id=""edit"" method=""post"" enctype=""multipart/form-data""><p>已上传文件:" & strFileName1 & " <a href='admin_addphoto2.asp'>继续上传</a></p></form>"
		Response.Write "<script language=""Javascript"">try{parent.document.getElementById('MyEditor___Frame').contentWindow.frames[0].document.getElementsByTagName('body')[0].innerHTML+='"&Replace(TransferHTML(UBBCode(strFileName,"[link][image][media][flash]"),"[upload]"),"'","\'")&"'}catch(e){}</script>"
		Response.Write "</body></html>"

	Else

		Call SetBlogHint_Custom("√ 上传照片成功,如果是批量上传,照片信息都是一样的,请单个编辑.</a>")
		Response.Redirect "admin_addphoto.asp?typeid=" & zhuanti

	End If

End If
%>
<%
Call System_Terminate()

'If Err.Number<>0 Then
   'Call ShowError(0)
'End If
%>