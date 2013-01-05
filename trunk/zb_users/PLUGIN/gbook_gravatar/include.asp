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
Dim DZ_ISREPLY		'显示回复内容
Dim DZ_USERIDS		'不显示某ID用户的评论
Dim DZ_STYLE_VALUE   '外观样式

Dim gbook_gravatar_Config


Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","$(document).ready(function(){ $(""head"").append(""<link rel='stylesheet' type='text/css' href='"&GetCurrentHost()&"zb_users/plugin/gbook_gravatar/css/gbook_gravatar.css'/>"");});")


Call RegisterPlugin("gbook_gravatar","ActivePlugin_gbook_gravatar")

Function ActivePlugin_gbook_gravatar()

	Call Add_Action_Plugin("Action_Plugin_BlogReBuild_Comments_Begin","gbook_gravatar_BlogReBuild_GuestComments():Exit Function")
	
End Function


'安装插件
function InstallPlugin_gbook_gravatar()

	Set gbook_gravatar_Config = New TConfig
	gbook_gravatar_Config.Load("gbook_gravatar")
	If gbook_gravatar_Config.Exists("DZ_VERSION")=False Then
		gbook_gravatar_Config.Write "DZ_VERSION","1.0"
'		gbook_gravatar_Config.Write "DZ_IDS_VALUE","0"
		gbook_gravatar_Config.Write "DZ_AVATAR_VALUE","wavatar"
		gbook_gravatar_Config.Write "DZ_WH_VALUE","32"
		gbook_gravatar_Config.Write "DZ_TITLE_VALUE","40"
		gbook_gravatar_Config.Write "DZ_COUNT_VALUE","8"
		gbook_gravatar_Config.Write "DZ_ISREPLY","0"
		gbook_gravatar_Config.Write "DZ_USERIDS","0"
		gbook_gravatar_Config.Write "DZ_STYLE_VALUE","1"
		gbook_gravatar_Config.Save
		Call SetBlogHint_Custom("您是第一次安装最新留言调用插件，已经为您导入初始配置。")
	End If
	


end function


'卸载插件
Function UnInstallPlugin_gbook_gravatar()
'更新侧栏
	BlogReBuild_Functions
End Function

Function gbook_gravatar_Initialize()
	Set gbook_gravatar_Config = New TConfig
	gbook_gravatar_Config.Load("gbook_gravatar")
	'DZ_IDS_VALUE = gbook_gravatar_Config.Read ("DZ_IDS_VALUE")
	DZ_AVATAR_VALUE=gbook_gravatar_Config.Read ("DZ_AVATAR_VALUE")
	DZ_WH_VALUE=gbook_gravatar_Config.Read ("DZ_WH_VALUE")
	DZ_TITLE_VALUE=gbook_gravatar_Config.Read ("DZ_TITLE_VALUE")
	DZ_COUNT_VALUE=gbook_gravatar_Config.Read ("DZ_COUNT_VALUE")
	DZ_ISREPLY=gbook_gravatar_Config.Read ("DZ_ISREPLY")
	DZ_USERIDS=gbook_gravatar_Config.Read ("DZ_USERIDS")
	DZ_STYLE_VALUE=gbook_gravatar_Config.Read ("DZ_STYLE_VALUE")
End Function

'过滤掉ubb代码
Function noubb(str) 
 dim ubb  
 Set ubb=new RegExp  
 ubb.IgnoreCase =true 
 ubb.Global=True 
 ubb.Pattern="(\[.[^\[]*\])" 
 str=ubb.replace(str," ")  
 ubb.Pattern="(\[\/[^\[]*\])" 
 str=ubb.replace(str," ")  
 noubb=str 
 set ubb=nothing 
end function 


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
	Call GetFunction
	Call gbook_gravatar_Initialize
	Dim strComments
	Dim gbook_gravatar_objArticle
	Dim s
	Dim i,iii,sqlstr
	Dim t_mail_e
	Dim DZ_Rs,sql1,sql2,sql3
	Dim u,u_t,t_users,uss
	
	if DZ_USERIDS <>"" then
		t_users = split(DZ_USERIDS,",")
		for iii=0 to Ubound(t_users)
			uss = uss & "'"&t_users(iii)&"',"
		next
		uss = left(uss,len(uss)-1)
	end if
	
	
	if DZ_IDS_VALUE <> "0" then sql1= " and log_ID not in("&DZ_IDS_VALUE&")" end if
	if DZ_ISREPLY = "0" then sql2 = " and comm_ParentID=0" end if
	if DZ_USERIDS <> "" then sql3 = " and comm_Author not in ("&uss&")" end if
	
	sqlstr = "SELECT * FROM [blog_Comment] where comm_IsCheck=0 "& sql2 & sql3 &"  ORDER BY [comm_ID] DESC"

	Set DZ_Rs=objConn.Execute(sqlstr)
	If (Not DZ_Rs.bof) And (Not DZ_Rs.eof) Then

		For i=1 to cint(DZ_COUNT_VALUE)
			s=TransferHTML(noubb(DZ_Rs("comm_Content")),"[nohtml]")
			s=Replace(s,vbCrlf,"")
			
			s=gbook_gravatar_cutTitle(s,cint(DZ_TITLE_VALUE)) 
			if DZ_Rs("comm_Email")<>"" then
			t_mail_e=md5(DZ_Rs("comm_Email"))
			else
			t_mail_e=md5("")
			end if
			
			Set gbook_gravatar_objArticle=New TArticle
			If gbook_gravatar_objArticle.LoadInfoByID(DZ_Rs("log_ID")) Then
			
			u=DZ_Rs("comm_HomePage")
			if u="" then
			u="javascript:;"
			u_t=""
			else
			u_t=u
			end if
			
	if DZ_STYLE_VALUE=1 then
 		strComments=strComments & "<li><a href="""& gbook_gravatar_objArticle.Url & "#cmt" & DZ_Rs("comm_ID") & """ title=""" & DZ_Rs("comm_PostTime") & " post by " & DZ_Rs("comm_Author") & """><b>" & DZ_Rs("comm_Author") & "：</b>&nbsp;"&s&"</a></li>"&vbcrlf
 	elseif DZ_STYLE_VALUE=2 then
		strComments=strComments & "<li><a href="""& gbook_gravatar_objArticle.Url & "#cmt" & DZ_Rs("comm_ID") & """ title=""" & DZ_Rs("comm_PostTime") & " post by " & DZ_Rs("comm_Author") & """><img style=""width:"&DZ_WH_VALUE&"px;height:"&DZ_WH_VALUE&"px;margin:0 3.5px -3px 0;"" title=""" & DZ_Rs("comm_Author") & " 的头像"" alt=""" & DZ_Rs("comm_Author") & " 的头像"" src=""http://www.gravatar.com/avatar/"&t_mail_e&"?s="&DZ_WH_VALUE&"&nbsp;d="&DZ_AVATAR_VALUE&"&nbsp;r=G""/>&nbsp;"&s&"</a></li>"&vbcrlf
	elseif DZ_STYLE_VALUE=3 then
		strComments=strComments & "<li class=""C_CMT_32""><span class=""C_CMT_Gravatar""><img style=""border:1px solid #ccc;padding:2px 2px;width:"&DZ_WH_VALUE&"px;height:"&DZ_WH_VALUE&"px;"" title=""" & DZ_Rs("comm_Author") & " 的头像"" alt=""" & DZ_Rs("comm_Author") & " 的头像"" src=""http://www.gravatar.com/avatar/"&t_mail_e&"?s="&DZ_WH_VALUE&"&nbsp;d="&DZ_AVATAR_VALUE&"&nbsp;r=G""/></span><span class=""C_CMT_Content""><a href="""&u&""" title="""&u_t&""" rel=""nofollow"" class=""C_CMT_A"" target=""_blank"">" & DZ_Rs("comm_Author") & "</a>: <br /><a href="""& gbook_gravatar_objArticle.Url & "#cmt" & DZ_Rs("comm_ID") & """ title=""" & DZ_Rs("comm_PostTime") & " post by " & DZ_Rs("comm_Author") & """>"&s&"</a></span></li>"&vbcrlf
	end if
			
			end if
			
			set gbook_gravatar_objArticle = nothing
			

			DZ_Rs.MoveNext
			If DZ_Rs.eof Then Exit For
		Next
	End If
	DZ_Rs.close
	Set DZ_Rs=Nothing
	
	strComments=TransferHTML(strComments,"[no-asp]")
	Functions(FunctionMetas.GetValue("comments")).Content=strComments
	Functions(FunctionMetas.GetValue("comments")).Post()
	Functions(FunctionMetas.GetValue("comments")).SaveFile


End Function
'*********************************************************
%>