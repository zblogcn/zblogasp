<%
'注册插件
Dim WP_ALBUM_NAME,WP_ALBUM_INTRO,WP_SUB_DOMAIN,WP_SCRIPT_TYPE,WP_ORDER_BY,WP_SMALL_WIDTH,WP_SMALL_HEIGHT,WP_LIST_WIDTH,WP_LIST_HEIGHT,WP_UPLOAD_FILESIZE,WP_UPLOAD_DIR,WP_UPLOAD_DIRBY,WP_UPLOAD_RENAME,WP_WATERMARK_WIDTH_POSITION,WP_WATERMARK_HEIGHT_POSITION,WP_JPEG_FONTCOLOR,WP_JPEG_FONTBOLD,WP_JPEG_FONTSIZE,WP_JPEG_FONTQUALITY,WP_WATERMARK_AUTO,WP_WATERMARK_TYPE,WP_WATERMARK_TEXT,WP_WATERMARK_LOGO,WP_WATERMARK_ALPHA,WP_INDEX_PAGERCOUNT,WP_SMALL_PAGERCOUNT,WP_LIST_PAGERCOUNT,WP_BLOGPHOTO_ID,WP_IF_ASPJPEG,WP_HIDE_DIVFILESND

Call Registerplugin("WindsPhoto", "Activeplugin_WindsPhoto")
Dim WP_Config
Function WindsPhoto_Initialize()
	Set WP_Config=New TConfig
	WP_Config.Load "WindsPhoto"
	If WP_Config.Exists("WP_VER")=False Then
		WP_Config.Write "WP_VER","2.7.4":WP_Config.Write "WP_ALBUM_NAME","我的WindsPhoto":WP_Config.Write "WP_ALBUM_INTRO","<p>WindsPhoto是基于asp+access的Z-Blog图片相册管理插件，功能简洁实用。</p>":WP_Config.Write "WP_SUB_DOMAIN","":WP_Config.Write "WP_SCRIPT_TYPE","1":WP_Config.Write "WP_ORDER_BY","1":WP_Config.Write "WP_SMALL_WIDTH",144:WP_Config.Write "WP_SMALL_HEIGHT",144:WP_Config.Write "WP_LIST_WIDTH",600:WP_Config.Write "WP_LIST_HEIGHT",600:WP_Config.Write "WP_UPLOAD_FILESIZE",2048000:WP_Config.Write "WP_UPLOAD_DIR","photofile":WP_Config.Write "WP_UPLOAD_DIRBY","1":WP_Config.Write "WP_UPLOAD_RENAME","1":WP_Config.Write "WP_WATERMARK_WIDTH_POSITION","right":WP_Config.Write "WP_WATERMARK_HEIGHT_POSITION","bottom":WP_Config.Write "WP_JPEG_FONTCOLOR","#000":WP_Config.Write "WP_JPEG_FONTBOLD","true":WP_Config.Write "WP_JPEG_FONTSIZE","14":WP_Config.Write "WP_JPEG_FONTQUALITY","4":WP_Config.Write "WP_WATERMARK_AUTO","0":WP_Config.Write "WP_WATERMARK_TYPE","1":WP_Config.Write "WP_WATERMARK_TEXT","WindsPhoto":WP_Config.Write "WP_WATERMARK_LOGO","images/nopic.jpg":WP_Config.Write "WP_WATERMARK_ALPHA","0.7":WP_Config.Write "WP_INDEX_PAGERCOUNT",12:WP_Config.Write "WP_SMALL_PAGERCOUNT",18:WP_Config.Write "WP_LIST_PAGERCOUNT",8:WP_Config.Write "WP_BLOGPHOTO_ID",1:WP_Config.Write "WP_IF_ASPJPEG","1":WP_Config.Write "WP_HIDE_DIVFILESND","1":WP_Config.Save
	End If
	WP_ALBUM_NAME = WP_Config.Read("WP_ALBUM_NAME")
	WP_ALBUM_INTRO = WP_Config.Read("WP_ALBUM_INTRO")
	WP_SUB_DOMAIN = WP_Config.Read("WP_SUB_DOMAIN")
	WP_SCRIPT_TYPE = WP_Config.Read("WP_SCRIPT_TYPE")
	WP_ORDER_BY = WP_Config.Read("WP_ORDER_BY")
	WP_SMALL_WIDTH = CInt(WP_Config.Read("WP_SMALL_WIDTH"))
	WP_SMALL_HEIGHT = CInt(WP_Config.Read("WP_SMALL_HEIGHT"))
	WP_LIST_WIDTH = CInt(WP_Config.Read("WP_LIST_WIDTH"))
	WP_LIST_HEIGHT = CInt(WP_Config.Read("WP_LIST_HEIGHT"))
	WP_UPLOAD_FILESIZE = CLng(WP_Config.Read("WP_UPLOAD_FILESIZE"))
	WP_UPLOAD_DIR = WP_Config.Read("WP_UPLOAD_DIR")
	WP_UPLOAD_DIRBY = WP_Config.Read("WP_UPLOAD_DIRBY")
	WP_UPLOAD_RENAME = WP_Config.Read("WP_UPLOAD_RENAME")
	WP_WATERMARK_WIDTH_POSITION = WP_Config.Read("WP_WATERMARK_WIDTH_POSITION")
	WP_WATERMARK_HEIGHT_POSITION = WP_Config.Read("WP_WATERMARK_HEIGHT_POSITION")
	WP_JPEG_FONTCOLOR = WP_Config.Read("WP_JPEG_FONTCOLOR")
	WP_JPEG_FONTBOLD = WP_Config.Read("WP_JPEG_FONTBOLD")
	WP_JPEG_FONTSIZE = WP_Config.Read("WP_JPEG_FONTSIZE")
	WP_JPEG_FONTQUALITY = WP_Config.Read("WP_JPEG_FONTQUALITY")
	WP_WATERMARK_AUTO = WP_Config.Read("WP_WATERMARK_AUTO")
	WP_WATERMARK_TYPE = WP_Config.Read("WP_WATERMARK_TYPE")
	WP_WATERMARK_TEXT = WP_Config.Read("WP_WATERMARK_TEXT")
	WP_WATERMARK_LOGO = WP_Config.Read("WP_WATERMARK_LOGO")
	WP_WATERMARK_ALPHA = WP_Config.Read("WP_WATERMARK_ALPHA")
	WP_INDEX_PAGERCOUNT = CInt(WP_Config.Read("WP_INDEX_PAGERCOUNT"))
	WP_SMALL_PAGERCOUNT = CInt(WP_Config.Read("WP_SMALL_PAGERCOUNT"))
	WP_LIST_PAGERCOUNT = CInt(WP_Config.Read("WP_LIST_PAGERCOUNT"))
	WP_BLOGPHOTO_ID = CInt(WP_Config.Read("WP_BLOGPHOTO_ID"))
	WP_IF_ASPJPEG = WP_Config.Read("WP_IF_ASPJPEG")
	WP_HIDE_DIVFILESND = WP_Config.Read("WP_HIDE_DIVFILESND")
End Function

'安装插件
Function Installplugin_WindsPhoto()
    'On Error Resume Next
    Call WindsPhoto_Addto_Navbar()
    Call WindsPhoto_Copy_Template()
    Call WindsPhoto_Database_Rename()
    Call SetBlogHint_Custom("? 提示:[WindsPhoto]已启用,现在进入初始化系统设置.升级请手动修改include.asp数据库参数.2.7以下版本请<a href=""update.asp"">升级数据库结构</a>")
    Response.Redirect BlogHost &"zb_users/plugin/WindsPhoto/admin_setting.asp"
	
    Err.Clear
End Function

'卸载插件
Function UnInstallplugin_WindsPhoto()
    Call DelSiteFile("/zb_users/include/windsphoto_sort.asp")
    Call DelSiteFile("/photo.html")
    Call WindsPhoto_Delfrom_Navbar
    Call SetBlogHint_Custom("? 提示:[WindsPhoto]已停用,由本插件生成的文件已删除,你上传的图片文件以及数据库仍然保留,卸载请手动删除插件目录.")
End Function

Function Activeplugin_WindsPhoto()
    '网站管理加上二级菜单项
    Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu", MakeSubMenu("WindsPhoto设置", BlogHost & "zb_users/plugin/windsphoto/admin_setting.asp", "m-left", FALSE))
    Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu",MakeSubMenu("[WindsPhoto管理]",BlogHost & "zb_users/plugin/windsphoto/admin_main.asp","m-left",False))
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(2,"相册管理",GetCurrentHost&"zb_users/plugin/windsphoto/admin_main.asp","nav_windsphoto","aWindsPhoto",GetCurrentHost&"zb_users/plugin/windsphoto/images/MD-camera-photo.png"))
	Call Add_Action_Plugin("Action_Plugin_uEditor_FileUpload_Begin","WindsPhoto_uEditorUpload")
	Call Add_Filter_Plugin("Filter_Plugin_UEditor_Config","WindsPhoto_ExportUEConfig")	
	Call Add_Action_Plugin("Action_Plugin_uEditor_imageManager_Begin","WindsPhoto_uEditorAlbumList")
End Function

Function WindsPhoto_ExportUEConfig(m)
	
	Dim j,k
	j=Split(m,"imagePath:""")(0)
	k=Split(m,""",imageFieldName:")(1)
	OpenConnect
	WindsPhoto_Initialize
	Dim tmp
	tmp=BlogHost & WindsPhoto_GetImgFolder
'	response.write tmp

	m=j&"imagePath:"""&tmp&""",imageFieldName:"&k
	
End Function

'首次安装数据库改名....
Function WindsPhoto_Database_Rename()
	On Error Resume Next
	objConn.Execute "SELECT TOP 1 id FROM WINDSPHOTO_ZHUANTI"
	If Err.Number<>0 Then
		If ZC_MSSQL_ENABLE Then
			objConn.BeginTrans
			objConn.Execute "CREATE TABLE [WindsPhoto_zhuanti] ("&_
							"[id] int identity(1,1) not null primary key,"&_
							"[data] datetime,"&_
							"[time1] datetime,"&_
							"[js] nvarchar(max),"&_
							"[hot] int,"&_
							"[pass] nvarchar(30),"&_
							"[view] smallint,"&_
							"[name] nvarchar(50),"&_
							"[ordered] int)"
			objConn.Execute "CREATE TABLE [WindsPhoto_desktop] ("&_
							"[ID] int identity(1,1) not null primary key,"&_
							"[name] nvarchar(50),"&_
							"[zhuanti] smallint,"&_
							"[jj] nvarchar(max),"&_
							"[url] nvarchar(max),"&_
							"[surl] nvarchar(max),"&_
							"[hot] smallint,"&_
							"[itime] datetime,"&_
							"[viewnums] int"&_
							")"
			objConn.CommitTrans
		Else
			objConn.BeginTrans
			objConn.Execute "CREATE TABLE 'WindsPhoto_desktop' ("&_
							"'ID' AutoIncrement primary key,"&_
							"'name' VarChar(50),"&_
							"'zhuanti' Short,"&_
							"'jj' LongText,"&_
							"'url' LongText,"&_
							"'surl' LongText,"&_
							"'hot' Short,"&_
							"'itime' DateTime,"&_
							"'viewnums' Long"&_
							")"
			objConn.Execute "CREATE TABLE 'WindsPhoto_zhuanti' ("&_
							"'id' AutoIncrement primary key,"&_
							"'data' DateTime,"&_
							"'time1' DateTime,"&_
							"'js' LongText,"&_
							"'hot' Long,"&_
							"'pass' VarChar(30),"&_
							"'view' Short,"&_
							"'name' VarChar(50),"&_
							"'ordered' Long"&_
							")"
			objConn.CommitTrans
		End If
	End If
End Function

'添加到导航栏和后台菜单
Function WindsPhoto_Addto_Navbar()
	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=Functions(FunctionMetas.GetValue("navbar")).Content & "<li><a href=""<#ZC_BLOG_HOST#>zb_users/plugin/windsphoto/"">相册</a></li>"
	Functions(FunctionMetas.GetValue("navbar")).Save
	Call ClearGlobeCache
	Call LoadGlobeCache
End Function

'从后台菜单和导航栏删除
Function WindsPhoto_Delfrom_Navbar()
	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=RemoveLibyUrl(Functions(FunctionMetas.GetValue("navbar")).Content,"<#ZC_BLOG_HOST#>zb_users/plugin/windsphoto/")
	Functions(FunctionMetas.GetValue("navbar")).Save
	Call ClearGlobeCache
	Call LoadGlobeCache
End Function

'添加相册默认模板
Function WindsPhoto_Copy_Template()
    Dim strContent
    strContent = LoadFromFile(BlogPath & "/zb_users/Themes/" & ZC_BLOG_THEME & "/Template/page.html", "utf-8")
    strContent = Replace(strContent, "<#ZC_MSG050#>", "相册分类")
    strContent = Replace(strContent, "<#CACHE_INCLUDE_CALENDAR#>", "<ul><#CACHE_INCLUDE_WINDSPHOTO_SORT#></ul>")
    strContent = Replace(strContent, "id=""divCalendar""", "id=""divCatalog""")
    Call SaveToFile(BlogPath & "/zb_users/Theme/" & ZC_BLOG_THEME & "/Template/wp_index.html", strContent, "utf-8", TRUE)
    Call SaveToFile(BlogPath & "/zb_users/Theme/" & ZC_BLOG_THEME & "/Template/wp_album.html", strContent, "utf-8", TRUE)
End Function
	
Function WindsPhoto_GetImgFolder()
			Dim FilePath
			FilePath = "zb_users/plugin/windsphoto/" '设置上传目录位置
			If WP_UPLOAD_DIRBY = 1 Then
				CreatDirectoryByCustomDirectory("zb_users/plugin/windsphoto/" & WP_UPLOAD_DIR & "/" &Year(GetTime(Now()))&Month(GetTime(Now())))
				FilePath = FilePath & "/" & WP_UPLOAD_DIR & "/" &Year(GetTime(Now()))&Month(GetTime(Now())) & "/"
			ElseIf WP_UPLOAD_DIRBY = 2 Then
				CreatDirectoryByCustomDirectory("zb_users/plugin/windsphoto/" & WP_UPLOAD_DIR & "/" & zhuanti)
				FilePath = FilePath & "/" & WP_UPLOAD_DIR & "/" & zhuanti & "/"
			Else
				CreatDirectoryByCustomDirectory("zb_users/plugin/windsphoto/" & WP_UPLOAD_DIR)
				FilePath = FilePath & "/" & WP_UPLOAD_DIR & "/"
			End If
			WindsPhoto_GetImgFolder=filepath
End Function

Function WindsPhoto_uEditorUpload()
	On Error Resume Next
	Call WindsPhoto_Initialize
	If WP_BLOGPHOTO_ID <> 0 then 
		zhuanti=WP_BLOGPHOTO_ID
		If Instr(Request.ServerVariables("URL"),"imageUp.asp") Then 
			dim upload,file,state,uploadPath,PostTime
			Randomize
			
			PostTime=GetTime(Now())
			filepath=WindsPhoto_GetImgFolder
			FilePath=Replace(FilePath,"/","\")
			Dim formname
			formname="edtFileLoad"
			Set upload=New UpLoadClass
			upload.AutoSave=2
			upload.Charset="UTF-8"
			upload.FileType=Replace(ZC_UPLOAD_FILETYPE,"|","/")
			upload.savepath=BlogPath & FilePath
			upload.maxsize=WP_UPLOAD_FILESIZE
			upload.open
			Dim Path,FileNamet,FileNamelen,FileNamet1,imgWidth,imgHeight
			Path=Replace(BlogPath & FilePath & upload.form("edtFileLoad_Name"),"\","/")
			Dim s
			FileName=BlogHost & strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")
			Err.Clear
			Dim Jpeg
			Set Jpeg = Server.CreateObject("Persits.Jpeg")
			Dim haveJpeg
			If Err.Number=0 Then haveJpeg=True
			If upload.Save("edtFileLoad",0)=True Then
				Filename=FilePath & upload.form(formname)
				FileNamet=upload.form(formname)
				If WP_IF_ASPJPEG="1" And haveJpeg Then
					
					
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
					Dim Title,TitleWidth,PositionWidth,PositionHeight
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
				If upload.form("pictitle")<>"" Then
					name = TransferHTML(upload.form("pictitle"),"[&][<][>][""][space][enter][nohtml]")
				Else
					name = Replace(File.FileName, FileExt, "")
				End If
				
				'写入数据库
				strSQL = "insert into WindsPhoto_desktop ([name],[itime],zhuanti,jj,url,surl,hot) values ('"&name&"','"&now&"',"&zhuanti&",'"&photointro&"','"&photourlb&"','"&photourls&"','0')"
				objConn.Execute strSQL
				iCount = iCount + 1
		
			End If
			
			Dim strJSON
			strJSON="{'state':'"& upload.Error2Info("edtFileLoad") & "','url':'"& upload.form("edtFileLoad") &"','fileType':'"&upload.form("edtFileLoad_Ext")&"','title':'"&TransferHTML(upload.form("pictitle"),"[&][<][>][""][space][enter][nohtml]")&"','original':'"&upload.Form("edtFileLoad_Name")&"'}"
			
				
			For Each sAction_Plugin_uEditor_FileUpload_End in Action_Plugin_uEditor_FileUpload_End
				If Not IsEmpty(sAction_Plugin_uEditor_FileUpload_End) Then Call Execute(sAction_Plugin_uEditor_FileUpload_End)
			Next
			response.AddHeader "json",strjson
			response.write strJSON
			
			set upload=nothing

			Response.End
		end if
	end if

End Function

Function WindsPhoto_uEditorAlbumList()
	Call WindsPhoto_Initialize
	
	If WP_BLOGPHOTO_ID <> 0 then 
			dim db,conn,objRs,i
			Dim strResponse
			db=BlogPath & "\zb_users\plugin\windsphoto\" & WP_DATA_PATH
				
			Set objRs=objConn.Execute("SELECT TOP 100 url FROM WindsPhoto_desktop WHERE zhuanti="&WP_BLOGPHOTO_ID)

			If Not(objRs.Eof) Then
				For i=1 to objRS.PageSize
					Response.Write Replace(WP_SUB_DOMAIN & photofile & "/"&objRS("url")&uEditor_Split,"%","%25")
					objRS.MoveNext
					If objRS.Eof Then Exit For
					
				Next
			End If
			Response.End
	End If
	
End Function
%>