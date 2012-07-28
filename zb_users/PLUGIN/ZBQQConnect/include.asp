<!-- #include file="function\ZBConnectQQ_Public.asp"-->
<!-- #include file="function\ZBConnectQQ_JSON.asp"-->
<!-- #include file="function\ZBConnectQQ_DB.asp"-->

<%
'Temp 
Const ZBQQConnect_notfoundpic=""
Const ZBQQConnect_PicSendToWb=True
'Temp
Dim ZBQQConnect_CommentCore,ZBQQConnect_ACore,ZBQQConnect_CommentName,ZBQQConnect_EmlMD5

Dim ZBQQConnect_SToWb,ZBQQConnect_SToZone


dim ZBQQConnect_class,ZBQQConnect_DB

Function ZBQQConnect_Initialize
	set ZBQQConnect_class=new ZBQQConnect
	Set ZBQQConnect_DB=New ZBConnectQQ_DB
	ZBQQConnect_class.app_key="100291142"    '设置appkey
	ZBQQConnect_class.app_secret="6e39bee95a58a8c99dce88ad5169a50e"  '设置app_secret
	ZBQQConnect_class.callbackurl="http://www.zsxsoft.com/zblog-1-9/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp"  '设置回调地址
	ZBQQConnect_class.debug=false 'Debug模式设置
End Function



Call RegisterPlugin("ZBQQConnect","ActivePlugin_ZBQQConnect")

Sub ZBQQConnect_RegSave(UID)
	If Not IsEmpty(Request.Form("QQOpenID")) Then
		ZBQQConnect_Initialize
		ZBQQConnect_DB.OpenID=Request.Form("QQOpenID")
		If ZBQQConnect_DB.LoadInfo(4)=True Then
			ZBQQConnect_DB.objUser.LoadInfoById UID
			ZBQQConnect_DB.Bind
		End If
	End If
End Sub


Function ActivePlugin_ZBQQConnect() 
	Dim strQQ,objQQ
	If CheckPluginState("Reg")=True Then
			If IsEmpty(Request.QueryString("QQOPENID"))=False Then
				strQQ=Replace(TransferHTML(FilterSQL(Request.QueryString("QQOPENID")),"[no-html]"),"""","""""")
				If Len(strQQ)=32 Then
					ZBQQConnect_Initialize
					
					Call Add_Response_Plugin("Response_Plugin_RegPage_End","<input type=""hidden"" value="""&strQQ&""" name=""QQOpenID""/>")
					Set objQQ=ZBQQConnect_ToObject(ZBQQConnect_class.API("https://graph.qq.com/user/get_user_info","{'format':'json'}","GET&"))
					Call Add_Action_Plugin("Action_Plugin_RegPage_Begin","dUsername="""&objQQ.nickname&"""")
				End If
			End If
		Call Add_Action_Plugin("Action_Plugin_RegSave_End","Call ZBQQConnect_RegSave(RegUser.ID)")
		'Call Add_Action_Plugin("Action_Plugin_RegSave_End","If isQQLogin=True Then ")
	End If
	'挂上接口
	'Filter_Plugin_PostArticle_Core
	'Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","ZBQQConnect_CommentPst")
	'Call Add_Action_Plugin("Action_Plugin_CommentPost_Succeed","Call ZBQQConnect_SendComment()")
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","ZBQQConnect_SToZone=Request.Form(""syn_qq""):ZBQQConnect_SToWb=Request.Form(""syn_tqq""):Call ZBQQConnect_Main()")
	Call Add_Action_Plugin("Action_Plugin_Edit_ueditor_Begin","Call ZBQQConnect_addForm()")
	'Call Add_Action_Plugin("Action_Plugin_System_Initialize_Succeed","Call Add_Response_Plugin(""Response_Plugin_SiteInfo_SubMenu"",ZBQQConnect_MakeSM)")

	
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
Dim CSS,JS,HTML,ResponseText
CSS="<style type=""text/css"">.syn_qq, .syn_tqq, .syn_qq_check, .syn_tqq_check{display: inline-block;margin-top: 3px;width: 19px;height: 19px;background: transparent url(../../zb_users/plugin/zbqqconnect/connect_post_syn.png) no-repeat 0 0;line-height: 64px;overflow: hidden;vertical-align: top;cursor: pointer;}.syn_tqq{background-position: 0 -22px;margin-left: 5px;}.syn_qq_check{background-position: -22px 0;}.syn_tqq_check{background-position: -22px -22px;margin-left: 5px;}</style>"
JS="<script type='text/javascript'>var a=true,b=true;var d=$('#connectPost_synQQ');var f=$('#connectPost_synT');function c(){if(a){d.removeClass('syn_qq_check');d.addClass('syn_qq');d.attr('title','未设置同步至QQ空间');$('#syn_qq').attr('value','0');a=false}else{d.removeClass('syn_qq');d.addClass('syn_qq_check');d.attr('title','已设置同步至QQ空间');$('#syn_qq').attr('value','1');a=true}};function e(){if(b){f.removeClass('syn_tqq_check');f.addClass('syn_tqq');f.attr('title','未设置同步至腾讯微博');$('#syn_tqq').attr('value','0');b=false}else{f.removeClass('syn_tqq');f.addClass('syn_tqq_check');f.attr('title','已设置同步至腾讯微博');$('#syn_tqq').attr('value','1');b=true}};$(document).ready(function(){d.bind('click',function(){c()});f.bind('click',function(){e()})});</script>"
HTML="<a title='已设置同步至QQ空间' class='syn_qq_check' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a><input type='hidden' name='syn_qq' id='syn_qq' value='1'/><a title='已设置同步至腾讯微博' class='syn_tqq_check' href='javascript:void(0);' id='connectPost_synT'>腾讯微博</a><input type='hidden' name='syn_tqq' id='syn_tqq' value='1'/>"

	'If Request.QueryString("id")="" Then
		ResponseText=CSS&HTML&JS
	'Else

		'If ZBQQConnect_EditPostSend=True Then
		'	ResponseText=TextStart&Text2&TextEnd
'		'Else
'			ResponseText=""
'		'End iF
'	End If

	

Call Add_Response_Plugin("Response_Plugin_Edit_Form3",ResponseText)

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
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Post","ZBQQConnect_ArticlePst")	
End Function


Function ZBQQConnect_CommentPst(objA)
	on error resume next
	Set ZBQQConnect_CommentCore=objA
End Function



Function ZBQQConnect_SendComment()
	on error resume next
	Call ZBQQConnect_Initialize
	Dim strT,tea
	if ZBQQConnect_class.logined=false then exit function
	If ZBQQConnect_CommentCore.ID = 0 then Exit Function
	If ZBQQConnect_CmtSend=False Then Exit Function
	strT=ZBQQConnect_Ttent
    Dim ZBQQConnect_test
	set ZBQQConnect_test = new tarticle
	ZBQQConnect_test.loadinfobyid(ZBQQConnect_CommentCore.Log_Id)
	strT = Replace(strT,"%u",ZBQQConnect_test.url & "#cmt" & ZBQQConnect_CommentCore.id) 
	strT = Replace(strT,"%t",ZBQQConnect_test.title) 
	set ZBQQConnect_test = nothing
	strT = Replace(strT,"%b",ZC_BLOG_NAME)
	strT = Replace(strT,"%s",ZBQQConnect_CommentCore.Author)
	If Len(ZBQQConnect_CommentCore.Content) <= ZBQQConnect_strLong Then
	    tea=ZBQQConnect_ReplaceXO(UBBCode(replace(replace(replace(ZBQQConnect_CommentCore.Content,vbcrlf," "),vbcr," "),vblf," "),"[link][email][font][face]"))
	Else
	    tea=left(ZBQQConnect_ReplaceXO(UBBCode(replace(replace(replace(ZBQQConnect_CommentCore.Content,vbcrlf," "),vbcr," "),vblf," "),"[link][email][font][face]")),ZBQQConnect_strLong) & "..."
	End If
	tea=TransferHTML(tea,"[nohtml]")
	strT=Replace(strT,"%c",tea)
	call ZBQQConnect_class.Run(2,strT,Request.ServerVariables("REMOTE_ADDR"),"","")
	set ZBQQConnect_CommentCore = nothing
	Call ZBQQConnect_Terminate
End Function


Function ZBQQConnect_ArticlePst(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)
	on error resume next
	Dim ZBQQConnect_ACore
	set ZBQQConnect_acore=new tarticle
	ZBQQConnect_ACore.id=id
	ZBQQConnect_ACore.title=title
	ZBQQConnect_ACore.CateID=CateID
	ZBQQConnect_ACore.authorid=authorid
	ZBQQConnect_ACore.intro=intro
	ZBQQConnect_ACore.Content=Content
	ZBQQConnect_ArticleToWb ZBQQConnect_ACore
	set ZBQQConnect_ACore=nothing
end function

function ZBQQConnect_ArticleToWb(ByRef ZBQQConnect_ACore)
	Dim strT ,bolN,objTemp,strTemp
	
	If ZBQQConnect_ACore.CateID=0 Then Exit Function
	Call ZBQQConnect_Initialize
	Set ZBQQConnect_DB.objUser=BlogUser
	ZBQQConnect_DB.LoadInfo 2
	ZBQQConnect_Class.OpenID=ZBQQConnect_DB.OpenID
	ZBQQConnect_Class.AccessToken=ZBQQConnect_DB.AccessToken
	If IsObject(ZBQQConnect_ACore)=False Then Exit Function
		If ZBQQConnect_ACore.ID=0 then
			bolN=True
			Dim objRS
			Set objRS=objConn.Execute("SELECT TOP 1 log_ID FROM [blog_Article] ORDER BY log_ID desc")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ZBQQConnect_ACore.ID=objRS(0) + 1 
			Else
				ZBQQConnect_ACore.ID=1
			End If
		Else
			bolN=False
		End If

	If int(ZBQQConnect_ACore.level)>2 Then
		strTemp=TransferHTML(UBBCode(ZBQQConnect_ACore.Intro,"[link][email][font][code][face][image][flash][typeset][media][autolink][link-antispam]"),"[nohtml]")
		strTemp=Replace(ZBQQConnect_ReplaceXO(strTemp),"'","\'")
		
		dim t_add,tupian
		if ZBQQConnect_PicSendToWb=true then
			tupian=UBBCode(ZBQQConnect_ACore.Content,"[image]")
			tupian=ZBQQConnect_LoadPicture(tupian)
			if tupian="" then tupian=ZBQQConnect_notfoundpic
			tupian=replace(replace(tupian,"\","/"),"'","\'")
		else
			tupian="~"
		end if
		
		Dim strJSON
		
		if tupian<>"~" then
			strJSON="{'title':'"&Replace(ZBQQConnect_ACore.Title,"'","\'")&"','url':'"&Replace(ZBQQConnect_ACore.Url,"'","\'")&"','summary':'"&strTemp&"','images':'"&tupian&"'"
		Else
			strJSON="{'title':'"&Replace(ZBQQConnect_ACore.Title,"'","\'")&"','url':'"&Replace(ZBQQConnect_ACore.Url,"'","\'")&"','summary':'"&strTemp&"'"
		End If
		'Call SetBlogHint_Custom(strjSON)
		strJSON=strJSON & ",'nswb':1"
		strJSON=strJSON & "}"
		t_add = ZBQQConnect_class.API("https://graph.qq.com/share/add_share",strJSON,"POST&")
		Set t_add=ZBQQConnect_Toobject(t_add)
		If t_add.ret=0 Then
			Call SetBlogHint_Custom("恭喜，同步到QQ空间成功")
		else
			Call SetBlogHint_Custom("同步到QQ空间出现问题" & t_add.ret)
			Call SetBlogHint_Custom("调试信息：<br/>OpenID="&ZBQQConnect_Class.OpenID&"<br/>AccessToken="&ZBQQConnect_Class.AccessToken&"<br/>PrivateJSON="&strJSON)'&"<br/>URL="&)
	    End If
		If ZBQQConnect_SToWb=True Then 
			strJSON="{'format':'json','content':'更新了博客：《"&Replace(ZBQQConnect_ACore.Title,"'","\'")&"》，"&Replace(ZBQQConnect_ACore.Url,"'","\'")&"','clientip':'"&Request.ServerVariables("REMOTE_ADDR")&"'}"
			t_add = ZBQQConnect_class.API("https://graph.qq.com/t/add_t",strJSON,"POST&")
			Set t_add=ZBQQConnect_Toobject(t_add)
			If t_add.ret=0 Then
			Call SetBlogHint_Custom("恭喜，同步到腾讯微博成功")
			else
				Call SetBlogHint_Custom("同步到腾讯微博出现问题" & t_add.ret)
				Call SetBlogHint_Custom("调试信息：<br/>OpenID="&ZBQQConnect_Class.OpenID&"<br/>AccessToken="&ZBQQConnect_Class.AccessToken&"<br/>PrivateJSON="&strJSON)'&"<br/>URL="&)
			End If
		End If
	End If
		'If Instr(lcase(ZC_USING_PLUGIN_LIST),"autoposturl") Then
		'	If InStr(Split(lcase(ZC_USING_PLUGIN_LIST),"autoposturl")(0),"ZBQQConnect") Then
		'		Call SetBlogHint_Custom("<span style='color:red'>发现自动文章别名插件！请<a href='" & zc_blog_host & "/ZB_SYSTEM/cmd.asp?act=PlugInDisable&name=ZBQQConnect' target='_blank'>点击这里停用</a>ZBQQConnect然后<a href='" & zc_blog_host & "/ZB_SYSTEM/cmd.asp?act=PlugInActive&name=ZBQQConnect' target='_blank'>重新启用以使ZBQQConnect兼容该插件！</a></span>")
		'	End If
		'End If
	if bolN=true Then ZBQQConnect_ACore.ID=0
	Call ZBQQConnect_Terminate
end function


%>