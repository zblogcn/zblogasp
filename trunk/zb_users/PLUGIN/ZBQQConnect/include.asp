<!-- #include file="function\ZBConnectQQ_QQConnect.asp"--><%'QQ连接主程序%>
<!-- #include file="function\ZBConnectQQ_Wb.asp"--><%'微博程序%>
<!-- #include file="function\ZBConnectQQ_JSON.asp"--><%'JSON处理类%>
<!-- #include file="function\ZBConnectQQ_DB.asp"--><%'数据库类%>
<!-- #include file="function\ZBConnectQQ_HMACSHA1.asp"--><%'微博oAuth1 HMAC-SHA1%>
<!-- #include file="function\ZBConnectQQ_NetWork.asp"--><%'网络操作类%>
<!-- #include file="function\ZBConnectQQ_Public.asp"--><%'公共函数%>

<%
'ZBQQConnect插件配置全局变量，必须先ZBQQConnect_Initialize
Dim ZBQQConnect_notfoundpic,ZBQQConnect_PicSendToWb, ZBQQConnect_strLong, ZBQQConnect_CommentToOwner, ZBQQConnect_OpenComment, ZBQQConnect_DefaultToZone, ZBQQConnect_DefaultToT, ZBQQConnect_CommentToZone, ZBQQConnect_CommentToT, ZBQQConnect_allowQQLogin, ZBQQConnect_allowQQReg, ZBQQConnect_HeadMode, ZBQQConnect_Head, ZBQQConnect_Content, ZBQQConnect_WBKey, ZBQQConnect_WBSecret, ZBQQConnect_CommentTemplate

'同步时使用临时变量
Dim ZBQQConnect_tmpObj,ZBQQConnect_Eml(1)
Dim ZBQQConnect_SToWb,ZBQQConnect_SToZone

'定义各种Class，包括QQ连接、数据库、配置和网络
Dim ZBQQConnect_class,ZBQQConnect_DB,ZBQQConnect_Config,ZBQQConnect_Net

'注册插件
Call RegisterPlugin("ZBQQConnect","ActivePlugin_ZBQQConnect")


Function ActivePlugin_ZBQQConnect() 
	Dim strQQ,objQQ
	If CheckPluginState("RegPage")=True Then
			If IsEmpty(Request.QueryString("QQOPENID"))=False Then
				strQQ=Replace(TransferHTML(FilterSQL(Request.QueryString("QQOPENID")),"[no-html]"),"""","""""")
				If Len(strQQ)=32 Then
					ZBQQConnect_Initialize
					ZBQQConnect_DB.OpenID=strQQ
					If ZBQQConnect_DB.LoadInfo(4)=True Then
						ZBQQConnect_Class.OpenID=strQQ
						ZBQQConnect_Class.AccessToken=ZBQQConnect_DB.AccessToken
						Call Add_Response_Plugin("Response_Plugin_RegPage_End","<input type=""hidden"" value="""&strQQ&""" name=""QQOpenID""/>")
					End If
				End If

			End If
		Call Add_Action_Plugin("Action_Plugin_RegSave_End","Call ZBQQConnect_RegSave(RegUser.ID)")
	
	End If
	Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","ZBQQConnect_CommentPst")
	Call Add_Action_Plugin("Action_Plugin_CommentPost_Succeed","Call ZBQQConnect_SendComment()")
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","ZBQQConnect_SToZone=Request.Form(""syn_qq""):ZBQQConnect_SToWb=Request.Form(""syn_tqq""):Call ZBQQConnect_Main()")
	Call Add_Action_Plugin("Action_Plugin_Edit_ueditor_getArticleInfo","Set ZBQQConnect_tmpObj=EditArticle:Call ZBQQConnect_addForm()")
	Call Add_Filter_Plugin("Filter_Plugin_TComment_LoadInfoByArray","ZBQQConnect_getcmt")
	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_Template","ZBQQConnect_AddCommentCode")
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(5,"QQ互联",GetCurrentHost&"zb_users/plugin/zbqqconnect/main.asp","nav_QQConnect","aQQConnect",GetCurrentHost&"zb_users/plugin/zbqqconnect/Connect_logo_1.png"))

	
End Function

'得到评论者的E-Mail和登录帐号
Function ZBQQConnect_getcmt(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,IsCheck,MetaString)
	ZBQQConnect_Eml(0)=Email
	ZBQQConnect_Eml(1)=AuthorID
End Function

'添加评论代码
Function ZBQQConnect_AddCommentCode(ByRef a)
	Dim c
	If Instr(a,"article/comment/avatar") And ZBQQConnect_HeadMode=2 Then
		ZBQQConnect_Initialize
		ZBQQConnect_DB.Email=ZBQQConnect_Eml(0)
		If ZBQQConnect_Eml(0)="" Then
			If ZBQQConnect_Eml(1)>0 Then
				Set c=New TUser
				c.LoadInfoById ZBQQConnect_Eml(1)
				ZBQQConnect_DB.Email=c.Email
				Set c=Nothing
			End If
		End If
		If ZBQQConnect_DB.LoadInfo(3)=True And ZBQQConnect_DB.EMail<>"" Then 
					If ZBQQConnect_HeadMode=0 Then
						a=Replace(a,"<#article/comment/avatar#>",ZBQQConnect_DB.tHead&"/100")
					ElseIf ZBQQConnect_HeadMode=1 Then
						a=Replace(a,"<#article/comment/avatar#>",ZBQQConnect_DB.QzoneHead)
					End If
		End If
	End If
End Function

'添加文章发布页的右侧图标
Function ZBQQConnect_addForm()
	ZBQQConnect_Initialize
	Dim CSS,JS,HTML,ResponseText
	CSS="<style type=""text/css"">.syn_qq, .syn_tqq, .syn_qq_check, .syn_tqq_check{display: inline-block;margin-top: 3px;width: 19px;height: 19px;background: transparent url(../../zb_users/plugin/zbqqconnect/connect_post_syn.png) no-repeat 0 0;line-height: 64px;overflow: hidden;vertical-align: top;cursor: pointer;}.syn_tqq{background-position: 0 -22px;margin-left: 5px;}.syn_qq_check{background-position: -22px 0;}.syn_tqq_check{background-position: -22px -22px;margin-left: 5px;}</style>"
	JS="<script type='text/javascript'>var a="&IIf(ZBQQConnect_DefaultToZone=True,"true","false")&",b="&IIf(ZBQQConnect_DefaultToT=True,"true","false")&";var d=$('#connectPost_synQQ');var f=$('#connectPost_synT');function c(){if(a){d.removeClass('syn_qq_check');d.addClass('syn_qq');d.attr('title','未设置同步至QQ空间');$('#syn_qq').attr('value','0');a=false}else{d.removeClass('syn_qq');d.addClass('syn_qq_check');d.attr('title','已设置同步至QQ空间');$('#syn_qq').attr('value','1');a=true}};function e(){if(b){f.removeClass('syn_tqq_check');f.addClass('syn_tqq');f.attr('title','未设置同步至腾讯微博');$('#syn_tqq').attr('value','0');b=false}else{f.removeClass('syn_tqq');f.addClass('syn_tqq_check');f.attr('title','已设置同步至腾讯微博');$('#syn_tqq').attr('value','1');b=true}};$(document).ready(function(){c();e();d.bind('click',function(){c()});f.bind('click',function(){e()})});</script>"
	HTML=IIF(ZBQQConnect_DefaultToZone=True,"<a title='已设置同步至QQ空间' class='syn_qq_check' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a><input type='hidden' name='syn_qq' id='syn_qq' value='1'/>","<a title='未设置同步至QQ空间' class='syn_qq' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a><input type='hidden' name='syn_qq' id='syn_qq' value='0'/>")
	Html=html&iif(ZBQQConnect_DefaultTot=True,"<a title='已设置同步至腾讯微博' class='syn_tqq_check' href='javascript:void(0);' id='connectPost_synT'>腾讯微博</a><input type='hidden' name='syn_tqq' id='syn_tqq' value='1'/>","<a title='未设置同步至腾讯微博' class='syn_tqq' href='javascript:void(0);' id='connectPost_synT'>腾讯微博</a><input type='hidden' name='syn_tqq' id='syn_tqq' value='0'/>")
	ResponseText=CSS&HTML&JS
	If ZBQQConnect_tmpObj.ID=0 Then
		Call Add_Response_Plugin("Response_Plugin_Edit_Form3",ResponseText)
	End If
	Set ZBQQConnect_tmpObj=nothing
End Function

'判断是否同步
Function ZBQQConnect_Main()
	If ZBQQConnect_SToWb="0" Then
		ZBQQConnect_SToWb=False
	Else
		ZBQQConnect_SToWb=True
	End If
	If ZBQQConnect_SToZone="0" Then
		ZBQQConnect_SToZone=False
	Else
		ZBQQConnect_SToZone=True
	End If
	Application(ZC_BLOG_CLSID&"ZBQQConnect_a")=ZBQQConnect_SToWb
	Application(ZC_BLOG_CLSID&"ZBQQConnect_b")=ZBQQConnect_SToZone
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Core","ZBQQConnect_ArticleToWb")
End Function

'得到评论object并赋值给临时变量
Function ZBQQConnect_CommentPst(objA)
	on error resume next
	Set ZBQQConnect_tmpObj=objA
End Function


'提交评论
Function ZBQQConnect_SendComment()
	on error resume next
	Call ZBQQConnect_Initialize
	If (ZBQQConnect_OpenComment=False) Then Exit Function
	Dim tupian
	Dim strT,tea,strTemp
	If ZBQQConnect_tmpObj.ID = 0 then Exit Function
    Dim objArticle
	set objArticle = new tarticle
	objArticle.loadinfobyid(ZBQQConnect_tmpObj.Log_Id)
	If Len(ZBQQConnect_tmpObj.Content) <= ZBQQConnect_strLong Then
	    tea=ZBQQConnect_ReplaceXO(replace(replace(replace(ZBQQConnect_tmpObj.Content,vbcrlf," "),vbcr," "),vblf," "))
	Else
	    tea=left(ZBQQConnect_ReplaceXO(replace(replace(replace(ZBQQConnect_tmpObj.Content,vbcrlf," "),vbcr," "),vblf," ")),ZBQQConnect_strLong) & "..."
	End If
	tea=TransferHTML(tea,"[nohtml]")
	strTemp=TransferHTML(UBBCode(objArticle.Intro,"[link][email][font][code][face][image][flash][typeset][media][autolink][link-antispam]"),"[nohtml]")
	strTemp=Replace(ZBQQConnect_ReplaceXO(strTemp),"'","\'")
	Dim t_add
	if ZBQQConnect_PicSendToWb=true then
		tupian=UBBCode(objArticle.Content,"[image]")
		tupian=ZBQQConnect_LoadPicture(tupian)
		if tupian="" then tupian=ZBQQConnect_notfoundpic
		tupian=replace(replace(tupian,"\","/"),"'","\'")
	else
		tupian="~"
	end if

	If BlogUser.Level<5 And ZBQQConnect_CommentToZone Then
		
		If ZBQQConnect_CommentToOwner=True Then
			Dim o
			Set o=objConn.Execute("SELECT TOP 1 [mem_ID] FROM [blog_Member] WHERE [mem_Level]=1")
			ZBQQConnect_DB.objUser.ID=o("mem_id")
			ZBQQConnect_DB.LoadInfo 2
			Set o=Nothing
		Else
			Set ZBQQConnect_DB.objUser=BlogUser
			If BlogUser.Meta.Exists("ZBQQConnect_a")=False Or (BlogUser.Meta.Exists("ZBQQConnect_a")=True And CBool(BlogUser.Meta.GetValue("ZBQQConnect_a"))=True) Then	ZBQQConnect_DB.LoadInfo 2
		End If
		
		ZBQQConnect_class.OpenID=ZBQQConnect_DB.OpenID
		ZBQQConnect_class.AccessToken=ZBQQConnect_DB.AccessToken
		If ZBQQConnect_DB.openID="" Then Exit Function
		t_add = ZBQQConnect_class.Share(objArticle.Title,objArticle.Url,tea,strTemp,tupian,1)
	End If
	If ZBQQConnect_CommentToT Then
		t_Add=objArticle.Meta.GetValue("ZBQQConnect_WBID")
		t_Add=ZBQQConnect_class.fakeQQConnect.r(Replace(Replace(ZBQQConnect_CommentTemplate,"%a",ZBQQConnect_tmpObj.Author),"%c",tea),t_Add)
	End If
	set ZBQQConnect_tmpObj = nothing
End Function


'确认是否是新建文章，如果是修改则不同步
Function ZBQQConnect_ArticleToWb(ByRef objArticle)
	Application(ZC_BLOG_CLSID&"ZBQQConnect_c")=objArticle.ID
	If objArticle.ID=0 Then Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","ZBQQConnect_GetArticleID")
End Function

'得到最新文章ID并发布批处理事件
Function ZBQQConnect_GetArticleID(ByRef objArticle)
	If CInt(Application(ZC_BLOG_CLSID&"ZBQQConnect_c"))=0 Then Call AddBatch("ZBQQConnect正在提交数据<br/>","ZBQQConnect_Batch "&objArticle.ID)	
End Function

'批处理
Function ZBQQConnect_Batch(id)
	On Error Resume Next
	Dim strT ,bolN,objTemp,strTemp
	Dim objArticle
	Set objArticle=New TArticle
	SetBlogHint_custom ID
	objArticle.LoadInfoById id
	If objArticle.CateID=0 Then Exit Function
	
	ZBQQConnect_SToWb=Application(ZC_BLOG_CLSID&"ZBQQConnect_a")
	ZBQQConnect_SToZone=Application(ZC_BLOG_CLSID&"ZBQQConnect_b")
	
	Call ZBQQConnect_Initialize
	Set ZBQQConnect_DB.objUser=BlogUser
	ZBQQConnect_DB.LoadInfo 2
	ZBQQConnect_Class.OpenID=ZBQQConnect_DB.OpenID
	ZBQQConnect_Class.AccessToken=ZBQQConnect_DB.AccessToken
	If IsObject(objArticle)=False Then Exit Function
	If int(objArticle.level)>2 Then
		strTemp=ZBQQConnect_r(objArticle.Intro)
		dim t_add,tupian
		if ZBQQConnect_PicSendToWb=true then
			tupian=UBBCode(objArticle.Content,"[image]")
			tupian=ZBQQConnect_LoadPicture(tupian)
			if tupian="" then tupian=ZBQQConnect_notfoundpic
			tupian=replace(replace(tupian,"\","/"),"'","\'")
		else
			tupian="~"
		end if
		If ZBQQConnect_SToZone=True Then
			t_add = ZBQQConnect_class.Share(objArticle.Title,Replace(objArticle.FullUrl,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST),"",strTemp,tupian,1)
			Set t_add=ZBQQConnect_Toobject(t_add)
			If t_add.ret=0 Then
				Response.Write "恭喜，同步到QQ空间成功"
			else
				Response.Write "同步到QQ空间出现问题" & t_add.ret
				Response.Write "调试信息：<br/>"&ZBQQConnect_class.debugMsg
			End If
		End If
		ZBQQConnect_class.debugMsg=""
		If ZBQQConnect_SToWb=True Then 
			t_add = ZBQQConnect_class.fakeQQConnect.t(Replace(Replace(Replace(Replace(Replace(ZBQQConnect_Content,"%t",ZBQQConnect_r(objArticle.Title)),"%u",objArticle.FullUrl),"%b",ZBQQConnect_r(BlogTitle)),"%i",strTemp),"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST),tupian)
			Set t_add=ZBQQConnect_Toobject(t_add)
			If t_add.ret=0 Then
				Response.Write  "恭喜，同步到腾讯微博成功"
				objArticle.Meta.SetValue "ZBQQConnect_WBID",t_add.data.id
			else
				Response.Write "同步到腾讯微博出现问题" & t_add.ret
			End If
		End If
	End If
		'If Instr(lcase(ZC_USING_PLUGIN_LIST),"autoposturl") Then
		'	If InStr(Split(lcase(ZC_USING_PLUGIN_LIST),"autoposturl")(0),"ZBQQConnect") Then
		'		Call SetBlogHint_Custom("<span style='color:red'>发现自动文章别名插件！请<a href='" & zc_blog_host & "/ZB_SYSTEM/cmd.asp?act=PlugInDisable&name=ZBQQConnect' target='_blank'>点击这里停用</a>ZBQQConnect然后<a href='" & zc_blog_host & "/ZB_SYSTEM/cmd.asp?act=PlugInActive&name=ZBQQConnect' target='_blank'>重新启用以使ZBQQConnect兼容该插件！</a></span>")
		'	End If
		'End If
end function

'处理西欧字符和HTML代码
Function ZBQQConnect_r(c)
	dim a
	a=c
	a=TransferHTML(UBBCode(a,"[link][email][font][code][face][image][flash][typeset][media][autolink][link-antispam]"),"[nohtml]")
	a=ZBQQConnect_ReplaceXO(a)
	ZBQQConnect_r=a
end function

%>