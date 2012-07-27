<!-- #include file="function\ZBConnectQQ_Public.asp"-->
<!-- #include file="function\ZBConnectQQ_JSON.asp"-->
<!-- #include file="function\ZBConnectQQ_DB.asp"-->

<%
Dim ZBQQConnect_CommentCore,ZBQQConnect_AEnable,ZBQQConnect_ACore,ZBQQConnect_CommentName,ZBQQConnect_EmlMD5

dim ZBQQConnect_class,ZBQQConnect_DB
set ZBQQConnect_class=new ZBQQConnect
Set ZBQQConnect_DB=New ZBConnectQQ_DB
ZBQQConnect_class.app_key="100291142"    '设置appkey
ZBQQConnect_class.app_secret="6e39bee95a58a8c99dce88ad5169a50e"  '设置app_secret
ZBQQConnect_class.callbackurl="http://www.zsxsoft.com/zblog-1-9/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp"  '设置回调地址
ZBQQConnect_class.debug=false 'Debug模式设置


Call RegisterPlugin("ZBQQConnect","ActivePlugin_ZBQQConnect")


Function ActivePlugin_ZBQQConnect() 

	'挂上接口
	'Filter_Plugin_PostArticle_Core
	'Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","ZBQQConnect_CommentPst")
	'Call Add_Action_Plugin("Action_Plugin_CommentPost_Succeed","Call ZBQQConnect_SendComment()")
	'Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","ZBQQConnect_AEnable=Request.Form(""ZBQQConnect_AEnable"")::Call ZBQQConnect_Main()")
	'Call Add_Action_Plugin("Action_Plugin_Edit_ueditor_Begin","Call ZBQQConnect_addForm()")
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
	tmp=BlogPath & replace(tmp,ZC_BLOG_HOST,"")
End Function


Function ZBQQConnect_Main()
	If IsEmpty(ZBQQConnect_AEnable)=True Then
		ZBQQConnect_AEnable=False
	Else
		ZBQQConnect_AEnable=True
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
	ZBQQConnect_ACore.tag=tag
	ZBQQConnect_ACore.cateid=cateid
	ZBQQConnect_ACore.title=title
	ZBQQConnect_ACore.content=content
	ZBQQConnect_ACore.level=level
	ZBQQConnect_ACore.authorid=authorid
	ZBQQConnect_ACore.posttime=posttime
	ZBQQConnect_ACore.commnums=commnums
	ZBQQConnect_ACore.viewnums=viewnums
	ZBQQConnect_ACore.alias=alias
	ZBQQConnect_ACore.istop=istop
	ZBQQConnect_ACore.intro=intro
	ZBQQConnect_ArticleToWb ZBQQConnect_ACore
	set ZBQQConnect_ACore=nothing
end function

function ZBQQConnect_ArticleToWb(ByRef ZBQQConnect_ACore)
	Dim strT ,bolN,objTemp,strTemp
	Call ZBQQConnect_Initialize
	If ZBQQConnect_class.logined=false Or ZBQQConnect_AEnable=False Or IsObject(ZBQQConnect_ACore)=False Then Exit Function
		If ZBQQConnect_ACore.ID=0 then
			bolN=True
			Dim objRS
			Set objRS=objConn.Execute("SELECT TOP 1 log_ID FROM [blog_Article] ORDER BY log_ID desc")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ZBQQConnect_ACore.ID=objRS(0) + 1 'ZBQQConnect_ACore.ID对ZBQQConnect_ACore.URL有影响
			Else
				ZBQQConnect_ACore.ID=1
			End If
		Else
			bolN=False
		End If

	If int(ZBQQConnect_ACore.level)>2 Then
		If bolN Then
			strT=Request.Form("ZBQQConnect_TmpSet1")
		Else
			If ZBQQConnect_EditPostSend=False Then Exit Function
			strT=Request.Form("ZBQQConnect_TmpSet2")
		End If
		strTemp=TransferHTML(UBBCode(ZBQQConnect_ACore.Intro,"[link][email][font][code][face][image][flash][typeset][media][autolink][link-antispam]"),"[nohtml]")
		strTemp=Replace(ZBQQConnect_ReplaceXO(strTemp),"'","\'")
		dim t_add,tupian
		if ZBQQConnect_PicSendToWb=true then
			tupian=UBBCode(ZBQQConnect_ACore.Content,"[image]")
			tupian=ZBQQConnect_RegExp_Execute2("<img [^>]*src=[""']?([^"">]+)['""]?[^>]+>",tupian,1,0)
			if tupian="" then tupian=ZBQQConnect_notfoundpic
			tupian=replace(replace(tupian,"\","/"),"'","\'")
		else
			tupian="~"
		end if
		
		Dim strJSON
		
		if tupian<>"~" then
			strJSON="{'title':'"&Replace(ZBQQConnect_ACore.Title,"'","\'")&"','url':'"&Replace(ZBQQConnect_ACore.Url,"'","\'")&"','summary':'"&strTemp&"','images':'"&tupian&"'}"
		Else
			strJSON="{'title':'"&Replace(ZBQQConnect_ACore.Title,"'","\'")&"','url':'"&Replace(ZBQQConnect_ACore.Url,"'","\'")&"','summary':'"&strTemp&"'}"
		End If
		t_add = ZBQQConnect_class.API("https://graph.qq.com/share/add_share",strJSON,"POST&")
		If instr(t_add,"ok") Then
			Call SetBlogHint_Custom("同步到腾讯微博成功,<a href=""http://t.qq.com/p/t/"&ZBQQConnect_toobject(t_add).data.id&""" target=""_blank"">点此查看</a>")
		else
			Call SetBlogHint_Custom("同步到腾讯微博出现问题" & t_add)
			Call SetBlogHint_Custom("要发的微博已经生成完成，您可以手动<A href='"&ZC_BLOG_HOST&"\ZB_USERS\plugin\ZBQQConnect\index.asp'>点我发微博</a>或到<a href='http://t.qq.com/"& ZBQQConnect_class.username & "' target='_blank'>官方平台发微博</a>。内容在下面<br/>" &  transferhtml(strT,"[no-html]"))
			if tupian<>"~" then Call SetBlogHint_Custom("带图片:"&tupian)
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