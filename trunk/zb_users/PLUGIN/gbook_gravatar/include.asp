<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    myllop-大猪
'// 版权所有:    www.dazhuer.cn
'// 技术支持:    myllop#gmail.com
'// 程序名称:    留言增加gravatar头像
'// 英文名称:    gbook_gravatar
'// 开始时间:    2009-5-10
'// 最后修改:    
'// 备    注:    only for zblog1.8
'///////////////////////////////////////////////////////////////////////////////


Dim DZ_IDS_VALUE	'获取文章ID
Dim DZ_AVATAR_VALUE	'默认头像
Dim DZ_WH_VALUE	'头像宽高
Dim DZ_TITLE_VALUE	'标题长度
Dim DZ_COUNT_VALUE	'调用条数

Dim gbook_gravatar_Config

	

Call RegisterPlugin("gbook_gravatar","ActivePlugin_gbook_gravatar")

Function ActivePlugin_gbook_gravatar()
	Call Add_Action_Plugin("Action_Plugin_CommentPost_Succeed","gbook_gravatar_BlogReBuild_GuestComments()")
	Call Add_Action_Plugin("Action_Plugin_CommentRev_Succeed","gbook_gravatar_BlogReBuild_GuestComments()")
	
End Function


'设置侧边栏
function set_Sidebar(tContent)
Dim t
	Set t=new Tfunction
	t.Name="最新留言"
	t.FileName="gbookgravatar"
	t.IsSystem=false
	t.SidebarID=11111
	t.Order=10
	t.Content=tContent
	t.HtmlID="divGbookGravatar"
	t.Ftype="ul"
	t.post

end function




'安装插件
function InstallPlugin_gbook_gravatar()

	Set gbook_gravatar_Config = New TConfig
	gbook_gravatar_Config.Load("gbook_gravatar")
	If gbook_gravatar_Config.Exists("DZ_VERSION")=False Then
		gbook_gravatar_Config.Write "DZ_VERSION","1.0"
		gbook_gravatar_Config.Write "DZ_IDS_VALUE","1"
		gbook_gravatar_Config.Write "DZ_AVATAR_VALUE","wavatar"
		gbook_gravatar_Config.Write "DZ_WH_VALUE","32"
		gbook_gravatar_Config.Write "DZ_TITLE_VALUE","40"
		gbook_gravatar_Config.Write "DZ_COUNT_VALUE","8"
		gbook_gravatar_Config.Save
		Call SetBlogHint_Custom("您是第一次安装最新留言调用插件，已经为您导入初始配置。")
	End If
	
	set_Sidebar("0")


end function


'卸载插件
Function UnInstallPlugin_gbook_gravatar()
'On Error Resume Next
	Call GetFunction()
	Functions(FunctionMetas.GetValue("gbookgravatar")).Del
'更新侧栏
	BlogReBuild_Functions
End Function

Function gbook_gravatar_Initialize()
	Set gbook_gravatar_Config = New TConfig
	gbook_gravatar_Config.Load("gbook_gravatar")
	DZ_IDS_VALUE = gbook_gravatar_Config.Read ("DZ_IDS_VALUE")
	DZ_AVATAR_VALUE=gbook_gravatar_Config.Read ("DZ_AVATAR_VALUE")
	DZ_WH_VALUE=gbook_gravatar_Config.Read ("DZ_WH_VALUE")
	DZ_TITLE_VALUE=gbook_gravatar_Config.Read ("DZ_TITLE_VALUE")
	DZ_COUNT_VALUE=gbook_gravatar_Config.Read ("DZ_COUNT_VALUE")
	
End Function



'*********************************************************
' 取字符串的前几个字,大于字数时,显示...
'*********************************************************
  function gbook_gravatar_cutTitle(ByVal strtitle,ByVal counts)   
	Dim RegExpObj,ReGCheck
	Set RegExpObj=new RegExp 
	RegExpObj.Pattern="^[\u4e00-\u9fa5]+$" 
	Dim l,t,c,i
	l=Len(strtitle)
	t=0
	For i=1 to l
	c=Mid(strtitle,i,1)   
	ReGCheck=RegExpObj.test(c)
	If ReGCheck Then
	  t=t+2
	Else
	  t=t+1
	End If
	If t>=counts Then
	  gbook_gravatar_cutTitle=left(strtitle,i)&"..."
	  Exit For
	Else
	  gbook_gravatar_cutTitle=strtitle
	End If
	Next
	Set RegExpObj=nothing 
	gbook_gravatar_cutTitle=Replace(gbook_gravatar_cutTitle,chr(10),"")
	gbook_gravatar_cutTitle=Replace(gbook_gravatar_cutTitle,chr(13),"")
end function  
	      



'*********************************************************
' 目的：    最新留言列表
'*********************************************************
Function gbook_gravatar_BlogReBuild_GuestComments()
Call gbook_gravatar_Initialize
	Dim strComments
	Dim gbook_gravatar_objArticle
	Dim s
	Dim i
	Dim t_mail_e
	Dim DZ_Rs
	
	Set DZ_Rs=objConn.Execute("SELECT top "&DZ_COUNT_VALUE&" * FROM [blog_Comment] WHERE [log_ID] in ("&DZ_IDS_VALUE&") ORDER BY [comm_ID] DESC")
	If (Not DZ_Rs.bof) And (Not DZ_Rs.eof) Then
	strComments = strComments & "<link rel=""stylesheet"" href=""" & ZC_BLOG_HOST & "zb_users/PLUGIN/gbook_gravatar/css/gbook_gravatar.css"" type=""text/css"" media=""screen"" />" & vbCrLf
		For i=1 to DZ_COUNT_VALUE
			s=TransferHTML(UBBCode(DZ_Rs("comm_Content"),"[face][link][autolink][font][code][image][media][flash]"),"[nohtml][vbCrlf][upload]")
			s=Replace(s,vbCrlf,"")
			
			s=gbook_gravatar_cutTitle(s,DZ_TITLE_VALUE) 
			if DZ_Rs("comm_Email")<>"" then
			t_mail_e=md5(DZ_Rs("comm_Email"))
			else
			t_mail_e=md5("myllop@163.com")
			end if
			
			Set gbook_gravatar_objArticle=New TArticle
			If gbook_gravatar_objArticle.LoadInfoByID(DZ_Rs("log_ID")) Then
	strComments=strComments & "<li><div class=""n_cmt_gravatar""><img class=""avatar"" title="""&DZ_Rs("comm_Content")&""" alt=""" & DZ_Rs("comm_Author") & " 的头像"" width="""&DZ_WH_VALUE&""" height="""&DZ_WH_VALUE&""" src=""http://www.gravatar.com/avatar/"&t_mail_e&"?s="&DZ_WH_VALUE&"&d="&DZ_AVATAR_VALUE&"&r=G""/></div><div class=""n_cmt_content""> <span class=""n_cmt_auth""><a href="""& gbook_gravatar_objArticle.Url & "#comment-" & DZ_Rs("comm_ID") & """ title=""" & DZ_Rs("comm_PostTime") & " post by " & DZ_Rs("comm_Author") & """>" & DZ_Rs("comm_Author") & "</a></span>  "&s&"  <font class=""n_cmt_time"">"&DZ_Rs("comm_PostTime")&"</font><div style=""clear:both;""></div></div></li>"&vbcrlf
			
			end if
			
			set gbook_gravatar_objArticle = nothing
			

			DZ_Rs.MoveNext
			If DZ_Rs.eof Then Exit For
		Next
	End If
	DZ_Rs.close
	Set DZ_Rs=Nothing

	strComments=TransferHTML(strComments,"[no-asp]")
	
	Call SaveToFile(BlogPath & "/zb_users/include/dzguestcomments.asp",strComments,"utf-8",True)
	
	'更新模块
	Call GetFunction()
		'Functions(FunctionMetas.GetValue("gbookgravatar")).SidebarID=11
		Functions(FunctionMetas.GetValue("gbookgravatar")).Content=strComments
		Functions(FunctionMetas.GetValue("gbookgravatar")).Save
	
	'更新侧栏
	BlogReBuild_Functions
	''''

	
	Call ClearGlobeCache()
	Call LoadGlobeCache()

End Function
'*********************************************************
%>