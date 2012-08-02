<!-- #include file="function\ZBConnectQQ_Public.asp"-->
<!-- #include file="function\ZBConnectQQ_Wb.asp"-->
<!-- #include file="function\ZBConnectQQ_JSON.asp"-->
<!-- #include file="function\ZBConnectQQ_DB.asp"-->
<!-- #include file="function\ZBConnectQQ_HMACSHA1.asp"-->
<!-- #include file="function\ZBConnectQQ_NetWork.asp"-->

<%
Session.CodePage=65001
'Temp 
Dim ZBQQConnect_notfoundpic
Dim ZBQQConnect_PicSendToWb
Dim ZBQQConnect_strLong
Dim ZBQQConnect_CommentToOwner
Dim ZBQQConnect_OpenComment
Dim ZBQQConnect_DefaultToZone
Dim ZBQQConnect_DefaultToT
Dim ZBQQConnect_CommentToZone
Dim ZBQQConnect_CommentToT
Dim ZBQQConnect_allowQQLogin
Dim ZBQQConnect_allowQQReg
Dim ZBQQConnect_HeadMode
Dim ZBQQConnect_Head
Dim ZBQQConnect_Content
Dim ZBQQConnect_WBKey
Dim ZBQQConnect_WBSecret
Dim ZBQQConnect_CommentTemplate
'Temp
Dim ZBQQConnect_tmpObj,ZBQQConnect_Eml(1)

Dim ZBQQConnect_SToWb,ZBQQConnect_SToZone


dim ZBQQConnect_class,ZBQQConnect_DB,ZBQQConnect_Config,ZBQQConnect_Net

Function ZBQQConnect_Initialize
	dim i
	Set ZBQQConnect_Config=New TConfig
	ZBQQConnect_Config.Load "ZBQQConnect"
	If ZBQQConnect_Config.Exists("-。-")=False Then
		ZBQQConnect_Config.Write "-。-","1.0"
		For i=97 To 105
			ZBQQConnect_Config.Write Chr(i),iIf(chr(i)<>"g",True,False)
		Next
		ZBQQConnect_Config.Write "a1","0"
		ZBQQConnect_Config.Write "Gravatar","http://www.gravatar.com/avatar/<#EmailMD5#>?s=40&d=<#ZC_BLOG_HOST#>%2FZB%5FSYSTEM%2Fimage%2Fadmin%2Favatar%2Epng"
		ZBQQConnect_Config.Write "content","更新了文章：《%t》，%u"
		ZBQQConnect_Config.Write "WBKEY","2e21c7b056f341b080d4d3691f3d50fb"
		ZBQQConnect_Config.Write "WBAPPSecret","1b84a3016c132a6839d082605b854bbe"
		ZBQQConnect_Config.Write "pl","@%a 评论 %c"
		ZBQQConnect_Config.Save
	End If
	
	ZBQQConnect_notfoundpic="~"
	ZBQQConnect_strLong=30
	ZBQQConnect_DefaultToZone=CBool(ZBQQConnect_Config.Read("a"))
	ZBQQConnect_DefaultTot=CBool(ZBQQConnect_Config.Read("b"))
	ZBQQConnect_PicSendToWb=CBool(ZBQQConnect_Config.Read("c"))
	ZBQQConnect_OpenComment=CBool(ZBQQConnect_Config.Read("d"))
	ZBQQConnect_CommentToZone=CBool(ZBQQConnect_Config.Read("e"))
	ZBQQConnect_CommentToT=CBool(ZBQQConnect_Config.Read("f"))
	ZBQQConnect_CommentToOwner=CBool(ZBQQConnect_Config.Read("g"))
	ZBQQConnect_allowQQLogin=CBool(ZBQQConnect_Config.Read("h"))
	ZBQQConnect_allowQQReg=CBool(ZBQQConnect_Config.Read("i"))
	ZBQQConnect_HeadMode=CInt(ZBQQConnect_Config.Read("a1"))
	ZBQQConnect_Head=ZBQQConnect_Config.Read("Gravatar")
	ZBQQConnect_Content=ZBQQConnect_Config.Read("content")
	ZBQQConnect_WBKey=ZBQQConnect_Config.Read("WBKEY")
	ZBQQConnect_WBSecret=ZBQQConnect_Config.Read("WBAPPSecret")
	ZBQQConnect_CommentTemplate=ZBQQConnect_Config.Read("pl")
	Set ZBQQConnect_Net=New ZBQQConnect_NetWork
	set ZBQQConnect_class=new ZBQQConnect
	Set ZBQQConnect_DB=New ZBConnectQQ_DB
	ZBQQConnect_class.app_key=ZBQQConnect_Config.Read("AppID")    '设置appkey
	ZBQQConnect_class.app_secret=ZBQQConnect_Config.Read("KEY")  '设置app_secret
	ZBQQConnect_class.callbackurl="http://www.zsxsoft.com/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp"  '设置回调地址
	ZBQQConnect_class.debug=false 'Debug模式设置
	ZBQQConnect_class.fakeQQConnect.app_key=ZBQQConnect_WBKey
	ZBQQConnect_class.fakeQQConnect.app_secret=ZBQQConnect_WBSecret
	ZBQQConnect_class.fakeQQConnect.Token=ZBQQConnect_Config.Read("WBToken")
	ZBQQConnect_class.fakeQQConnect.Secret=ZBQQConnect_Config.Read("WBSecret")
	ZBQQConnect_class.fakeQQConnect.UserID=ZBQQConnect_Config.Read("WBName")
End Function

 


Call RegisterPlugin("ZBQQConnect","ActivePlugin_ZBQQConnect")

Sub ZBQQConnect_RegSave(UID)
	If Not IsEmpty(Request.Form("QQOpenID")) Then
call setbloghint_custom(uid)
		ZBQQConnect_Initialize
		ZBQQConnect_DB.OpenID=Request.Form("QQOpenID")
		If ZBQQConnect_DB.LoadInfo(4)=True Then
			ZBQQConnect_DB.objUser.LoadInfoById UID
			ZBQQConnect_DB.Email=ZBQQConnect_DB.objUser.Email
			ZBQQConnect_DB.Bind
		End If
	End If
End Sub


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
						'objQQ=ZBQQConnect_class.API("https://graph.qq.com/user/get_user_info","{'format':'json'}","GET&")
						'Set objQQ=ZBQQConnect_ToObject(objqq)
						'Call Add_Action_Plugin("Action_Plugin_RegPage_Begin","dUsername="""&objQQ.nickname&"""")
					End If
				End If

			End If
		Call Add_Action_Plugin("Action_Plugin_RegSave_End","Call ZBQQConnect_RegSave(RegUser.ID)")
	
	End If
	'挂上接口
	'Filter_Plugin_PostArticle_Core
	Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","ZBQQConnect_CommentPst")
	Call Add_Action_Plugin("Action_Plugin_CommentPost_Succeed","Call ZBQQConnect_SendComment()")
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","ZBQQConnect_SToZone=Request.Form(""syn_qq""):ZBQQConnect_SToWb=Request.Form(""syn_tqq""):Call ZBQQConnect_Main()")
	Call Add_Action_Plugin("Action_Plugin_Edit_ueditor_getArticleInfo","Set ZBQQConnect_tmpObj=EditArticle:Call ZBQQConnect_addForm()")
	
	Call Add_Filter_Plugin("Filter_Plugin_TComment_LoadInfoByArray","ZBQQConnect_getcmt")
	
	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_Template","ZBQQConnect_AddCommentCode")
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(5,"QQ互联",GetCurrentHost&"zb_users/plugin/zbqqconnect/main.asp","","aQQConnect",GetCurrentHost&"zb_users/plugin/zbqqconnect/Connect_logo_1.png"))

	
End Function

Function ZBQQConnect_getcmt(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,IsCheck,MetaString)
	ZBQQConnect_Eml(0)=Email
	ZBQQConnect_Eml(1)=AuthorID
End Function
Function ZBQQConnect_AddCommentCode(ByRef a)
	Dim c
	If Instr(a,"<#ZBQQConnect_") Then
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
					If ZBQQConnect_DB.tHead<>"" Then
						a=Replace(a,"<#ZBQQConnect_tHead#>",ZBQQConnect_DB.tHead&"/100")
					Else
						a=Replace(a,"<#ZBQQConnect_tHead#>",Replace(Replace(ZBQQConnect_Head,"<#EmailMD5#>",MD5(ZBQQConnect_Eml(0))),"<#ZC_BLOG_HOST#>",GetCurrentHost))
					End If
					If ZBQQConnect_DB.QzoneHead<>"" Then
						a=Replace(a,"<#ZBQQConnect_zHead#>",ZBQQConnect_DB.QzoneHead)
					Else
						a=Replace(a,"<#ZBQQConnect_zHead#>",Replace(Replace(ZBQQConnect_Head,"<#EmailMD5#>",MD5(ZBQQConnect_Eml(0))),"<#ZC_BLOG_HOST#>",GetCurrentHost))
					End If
		End If
		a=Replace(a,"<#ZBQQConnect_Head#>",Replace(Replace(ZBQQConnect_Head,"<#EmailMD5#>",MD5(ZBQQConnect_Eml(0))),"<#ZC_BLOG_HOST#>",GetCurrentHost))
		a=Replace(a,"<#ZBQQConnect_zHead#>",Replace(Replace(ZBQQConnect_Head,"<#EmailMD5#>",MD5(ZBQQConnect_Eml(0))),"<#ZC_BLOG_HOST#>",GetCurrentHost))
		a=Replace(a,"<#ZBQQConnect_tHead#>",Replace(Replace(ZBQQConnect_Head,"<#EmailMD5#>",MD5(ZBQQConnect_Eml(0))),"<#ZC_BLOG_HOST#>",GetCurrentHost))

	End If
End Function

Function ZBQQConnect_LoadPicture(ByVal str)
	Dim objRegExp,Match,Matches,tmp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True
	objRegExp.Pattern="<img.*src\s*=\s*[\""|\']?\s*([^>\""\'\s]*)" 
	Set Matches=objRegExp.Execute(str)
	For Each Match in Matches 
		tmp=objRegExp.Replace(Match.Value,"$1") 
		Exit For
	Next
	set objregexp=nothing
	If Instr(tmp,"http")<0 And tmp<>"" Then tmp=ZC_BLOG_HOST & "/" & tmp
	ZBQQConnect_LoadPicture=tmp
	'tmp=BlogPath & replace(tmp,ZC_BLOG_HOST,"")
End Function

Function ZBQQConnect_addForm()
	ZBQQConnect_Initialize
	
	Dim CSS,JS,HTML,ResponseText
	CSS="<style type=""text/css"">.syn_qq, .syn_tqq, .syn_qq_check, .syn_tqq_check{display: inline-block;margin-top: 3px;width: 19px;height: 19px;background: transparent url(../../zb_users/plugin/zbqqconnect/connect_post_syn.png) no-repeat 0 0;line-height: 64px;overflow: hidden;vertical-align: top;cursor: pointer;}.syn_tqq{background-position: 0 -22px;margin-left: 5px;}.syn_qq_check{background-position: -22px 0;}.syn_tqq_check{background-position: -22px -22px;margin-left: 5px;}</style>"
	JS="<script type='text/javascript'>var a="&IIf(ZBQQConnect_DefaultToZone=True,"true","false")&",b="&IIf(ZBQQConnect_DefaultToT=True,"true","false")&";var d=$('#connectPost_synQQ');var f=$('#connectPost_synT');function c(){if(a){d.removeClass('syn_qq_check');d.addClass('syn_qq');d.attr('title','未设置同步至QQ空间');$('#syn_qq').attr('value','0');a=false}else{d.removeClass('syn_qq');d.addClass('syn_qq_check');d.attr('title','已设置同步至QQ空间');$('#syn_qq').attr('value','1');a=true}};function e(){if(b){f.removeClass('syn_tqq_check');f.addClass('syn_tqq');f.attr('title','未设置同步至腾讯微博');$('#syn_tqq').attr('value','0');b=false}else{f.removeClass('syn_tqq');f.addClass('syn_tqq_check');f.attr('title','已设置同步至腾讯微博');$('#syn_tqq').attr('value','1');b=true}};$(document).ready(c();e();function(){d.bind('click',function(){c()});f.bind('click',function(){e()})});</script>"
	HTML=IIF(ZBQQConnect_DefaultToZone=True,"<a title='已设置同步至QQ空间' class='syn_qq_check' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a><input type='hidden' name='syn_qq' id='syn_qq' value='1'/>","<a title='未设置同步至QQ空间' class='syn_qq' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a><input type='hidden' name='syn_qq' id='syn_qq' value='0'/>")
	Html=html&iif(ZBQQConnect_DefaultTot=True,"<a title='已设置同步至腾讯微博' class='syn_tqq_check' href='javascript:void(0);' id='connectPost_synT'>腾讯微博</a><input type='hidden' name='syn_tqq' id='syn_tqq' value='1'/>","<a title='未设置同步至腾讯微博' class='syn_tqq' href='javascript:void(0);' id='connectPost_synT'>腾讯微博</a><input type='hidden' name='syn_tqq' id='syn_tqq' value='0'/>")
	ResponseText=CSS&HTML&JS
	If ZBQQConnect_tmpObj.ID=0 Then
		Call Add_Response_Plugin("Response_Plugin_Edit_Form3",ResponseText)
	End If
	Set ZBQQConnect_tmpObj=nothing
End Function
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
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Core","ZBQQConnect_ArticleToWb")	
End Function


Function ZBQQConnect_CommentPst(objA)
	on error resume next
	Set ZBQQConnect_tmpObj=objA
End Function



Function ZBQQConnect_SendComment()
	'on error resume next
	
	Call ZBQQConnect_Initialize
	If (ZBQQConnect_OpenComment=False) Then Exit Function
	If ZBQQConnect_CommentToOwner=True Then
		Dim o
		Set o=objConn.Execute("SELECT TOP 1 [mem_ID] FROM [blog_Member] WHERE [mem_Level]=1")
		ZBQQConnect_DB.objUser.ID=o("mem_id")
		ZBQQConnect_DB.LoadInfo 2
		Set o=Nothing
	Else
		If (Not IsEmpty(Request.Cookies("QQOPENID"))) And (Not isNull(Request.Cookies("QQOPENID"))) And ( Request.Cookies("QQOPENID")<>"")  Then
			ZBQQConnect_DB.openID=Request.Cookies("QQOPENID")
			ZBQQConnect_DB.LoadInfo 4
		ElseIf BlogUser.Level<5 Then
			Set ZBQQConnect_DB.objUser=BlogUser
			ZBQQConnect_DB.LoadInfo 2
		Else
			ZBQQConnect_DB.Email=ZBQQConnect_tmpObj.Email
			ZBQQConnect_DB.LoadInfo 3
		End If
	End If
	ZBQQConnect_class.OpenID=ZBQQConnect_DB.OpenID
	ZBQQConnect_class.AccessToken=ZBQQConnect_DB.AccessToken
	Dim tupian
	If ZBQQConnect_DB.openID="" Then Exit Function
	Dim strT,tea,strTemp
	If ZBQQConnect_tmpObj.ID = 0 then Exit Function
    Dim objArticle
	set objArticle = new tarticle
	objArticle.loadinfobyid(ZBQQConnect_tmpObj.Log_Id)
	If Len(ZBQQConnect_tmpObj.Content) <= ZBQQConnect_strLong Then
	    tea=ZBQQConnect_ReplaceXO(UBBCode(replace(replace(replace(ZBQQConnect_tmpObj.Content,vbcrlf," "),vbcr," "),vblf," "),"[link][email][font][face]"))
	Else
	    tea=left(ZBQQConnect_ReplaceXO(UBBCode(replace(replace(replace(ZBQQConnect_tmpObj.Content,vbcrlf," "),vbcr," "),vblf," "),"[link][email][font][face]")),ZBQQConnect_strLong) & "..."
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
	If ZBQQConnect_CommentToZone Then t_add = ZBQQConnect_class.Share(objArticle.Title,objArticle.Url,tea,strTemp,tupian,1)
	If ZBQQConnect_CommentToT Then
		t_Add=objArticle.Meta.GetValue("ZBQQConnect_WBID")
		t_Add=ZBQQConnect_class.fakeQQConnect.r(Replace(Replace(ZBQQConnect_CommentTemplate,"%a",ZBQQConnect_tmpObj.Author),"%c",tea),t_Add)
	End If
	set ZBQQConnect_tmpObj = nothing
End Function



Function ZBQQConnect_ArticleToWb(ByRef objArticle)
	Dim strT ,bolN,objTemp,strTemp
	If objArticle.ID<>0 Then Exit Function
	If objArticle.CateID=0 Then Exit Function
	Call ZBQQConnect_Initialize
	Set ZBQQConnect_DB.objUser=BlogUser
	ZBQQConnect_DB.LoadInfo 2
	ZBQQConnect_Class.OpenID=ZBQQConnect_DB.OpenID
	ZBQQConnect_Class.AccessToken=ZBQQConnect_DB.AccessToken
	If IsObject(objArticle)=False Then Exit Function
		If objArticle.ID=0 then
			bolN=True
			Dim objRS
			Set objRS=objConn.Execute("SELECT TOP 1 log_ID FROM [blog_Article] ORDER BY log_ID desc")
			If (Not objRS.bof) And (Not objRS.eof) Then
				objArticle.ID=objRS(0) + 1 
			Else
				objArticle.ID=1
			End If
		Else
			bolN=False
		End If

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
			t_add = ZBQQConnect_class.Share(objArticle.Title,objArticle.Url,"",strTemp,tupian,1)
			Set t_add=ZBQQConnect_Toobject(t_add)
			If t_add.ret=0 Then
				Call SetBlogHint_Custom("恭喜，同步到QQ空间成功")
			else
				Call SetBlogHint_Custom("同步到QQ空间出现问题" & t_add.ret)
				Call SetBlogHint_Custom("调试信息：<br/>"&ZBQQConnect_class.debugMsg)'&"<br/>URL="&)
			End If
		End If
		ZBQQConnect_class.debugMsg=""
		If ZBQQConnect_SToWb=True Then 
			t_add = ZBQQConnect_class.fakeQQConnect.t(Replace(Replace(Replace(Replace(ZBQQConnect_Content,"%t",ZBQQConnect_r(objArticle.Title)),"%u",objArticle.Url),"%b",ZBQQConnect_r(BlogTitle)),"%i",strTemp),tupian)
			Set t_add=ZBQQConnect_Toobject(t_add)
			If t_add.ret=0 Then
				Call SetBlogHint_Custom("恭喜，同步到腾讯微博成功")
				objArticle.Meta.SetValue "ZBQQConnect_WBID",t_add.data.id
			else
				Call SetBlogHint_Custom("同步到腾讯微博出现问题" & t_add.ret)
			End If
		End If
	End If
		'If Instr(lcase(ZC_USING_PLUGIN_LIST),"autoposturl") Then
		'	If InStr(Split(lcase(ZC_USING_PLUGIN_LIST),"autoposturl")(0),"ZBQQConnect") Then
		'		Call SetBlogHint_Custom("<span style='color:red'>发现自动文章别名插件！请<a href='" & zc_blog_host & "/ZB_SYSTEM/cmd.asp?act=PlugInDisable&name=ZBQQConnect' target='_blank'>点击这里停用</a>ZBQQConnect然后<a href='" & zc_blog_host & "/ZB_SYSTEM/cmd.asp?act=PlugInActive&name=ZBQQConnect' target='_blank'>重新启用以使ZBQQConnect兼容该插件！</a></span>")
		'	End If
		'End If
	if bolN=true Then objArticle.ID=0
end function


Function ZBQQConnect_r(c)
	dim a
	a=c
	a=TransferHTML(UBBCode(a,"[link][email][font][code][face][image][flash][typeset][media][autolink][link-antispam]"),"[nohtml]")
	a=ZBQQConnect_ReplaceXO(a)
	ZBQQConnect_r=a
end function

%>