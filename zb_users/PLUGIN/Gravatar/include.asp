<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    Gravatar - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
Dim Gravatar_EmailMD5
'注册插件
Call RegisterPlugin("Gravatar","ActivePlugin_Gravatar")
'挂口部分
Function ActivePlugin_Gravatar()	
	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_TemplateTags","Gravatar_Add")
	Call Add_Response_Plugin("Response_Plugin_Admin_Footer","<script type=""text/javascript"">$(""#avatar"").attr(""src"","""&GetCurrentHost&"/ZB_USERS/plugin/Gravatar/LoadHeader.asp"")</script>")
End Function



Function Gravatar_Add(a,b)
	Dim c
	Set c=New TConfig
	c.Load "Gravatar"
	b(13)=Replace(c.Read("c"),"<#article/comment/emailmd5#>",b(11))
	Set c=Nothing
End Function

Function InstallPlugin_Gravatar
	Dim a
	Set a=New TConfig
	a.Load "Gravatar"
	If a.Exists("v")=False Then
		a.Write "v","1.0"
		a.Write "c","http://cn.gravatar.com/avatar/<#article/comment/emailmd5#>?s=32&d=<#ZC_BLOG_HOST#>zb_users/avatar/0.png"
		a.Save
	End If
End Function


%>