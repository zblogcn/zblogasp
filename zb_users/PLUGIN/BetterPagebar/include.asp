<%
'///////////////////////////////////////////////////////////////////////////////
'// 月上之木  12.8.20
'///////////////////////////////////////////////////////////////////////////////


'注册插件
Call RegisterPlugin("BetterPagebar","ActivePlugin_BetterPagebar")

Function ActivePlugin_BetterPagebar()

	'Action_Plugin_TArticleList_ExportBar_Begin 分页条
	Call Add_Action_Plugin("Action_Plugin_TArticleList_ExportBar_Begin","ExportBar=BetterPagebar_ExportBar(intNowPage,intAllPage,Template_PageBar,Template_PageBar_Previous,Template_PageBar_Next,Url,ListType):Exit Function")

End Function


'启用插件
Function InstallPlugin_BetterPagebar()
	Call BetterPagebar_Initialize
	'更新首页
	Call BlogReBuild_Default
End Function


'停用插件
Function UnInstallPlugin_BetterPagebar()
	'更新首页
	Call AddBatch(ZC_MSG259,"Call BlogReBuild_Default")
End Function


'配置
Dim	BetterPagebar_AlwaysShow
Dim	BetterPagebar_FristPage
Dim	BetterPagebar_LastPage
Dim	BetterPagebar_PrvePage
Dim	BetterPagebar_NextPage
Dim	BetterPagebar_FristPage_Tip
Dim	BetterPagebar_LastPage_Tip
Dim	BetterPagebar_PrvePage_Tip
Dim	BetterPagebar_NextPage_Tip
Dim	BetterPagebar_Extend


'初始化配置
Function BetterPagebar_Initialize()
	Dim c
	Set c = New TConfig
	c.Load("BetterPagebar")
	If c.Exists("BetterPagebar_AlwaysShow")=False Then
		c.Write "BetterPagebar_AlwaysShow","True"
		c.Write "BetterPagebar_FristPage","« 首页"
		c.Write "BetterPagebar_LastPage","尾页 »"
		c.Write "BetterPagebar_PrvePage","«"
		c.Write "BetterPagebar_NextPage","»"
		c.Write "BetterPagebar_FristPage_Tip","首页"
		c.Write "BetterPagebar_LastPage_Tip","尾页"
		c.Write "BetterPagebar_PrvePage_Tip","上一页"
		c.Write "BetterPagebar_NextPage_Tip","下一页"
		c.Write "BetterPagebar_Extend","..."
		c.Save
		Call SetBlogHint_Custom("第一次安装分页条优化插件，已经为您导入初始配置。")
	End If
	Set c=Nothing
End Function


'配置读取
Function BetterPagebar_Config()
	Dim c
	Set c = New TConfig
	c.Load("BetterPagebar")
	BetterPagebar_AlwaysShow = c.Read ("BetterPagebar_AlwaysShow")
	BetterPagebar_FristPage=c.Read ("BetterPagebar_FristPage")
	BetterPagebar_LastPage=c.Read ("BetterPagebar_LastPage")
	BetterPagebar_PrvePage=c.Read ("BetterPagebar_PrvePage")
	BetterPagebar_NextPage=c.Read ("BetterPagebar_NextPage")
	BetterPagebar_FristPage_Tip=c.Read ("BetterPagebar_FristPage_Tip")
	BetterPagebar_LastPage_Tip=c.Read ("BetterPagebar_LastPage_Tip")
	BetterPagebar_PrvePage_Tip=c.Read ("BetterPagebar_PrvePage_Tip")
	BetterPagebar_NextPage_Tip=c.Read ("BetterPagebar_NextPage_Tip")
	BetterPagebar_Extend=c.Read("BetterPagebar_Extend")
	Set c=Nothing
End Function



'*********************************************************
' 目的：接管分页条
'*********************************************************
Function BetterPagebar_ExportBar(intNowPage,intAllPage,ByRef Template_PageBar,ByRef Template_PageBar_Previous,ByRef Template_PageBar_Next,Url,ListType)


		Dim strPageBar,strPageBarF,strPageBarL
		
		Dim i
		Dim s
		Dim t

		t=Url

		'读取配置
		Call BetterPagebar_Config

		'ListType="DEFAULT"'CATEGORY'USER'DATE'TAGS
		If ListType="DEFAULT" Then
			If ZC_STATIC_MODE="ACTIVE" Then
				t=t & "?page=%n"
			End If
			If ZC_STATIC_MODE="REWRITE" Then
				t=Replace(t,".html","_%n.html")
			End If
			If ZC_STATIC_MODE="MIX" Then
				t=MixUrl
				t=t & "?page=%n"
			End If
		End If

		If ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS" Then
			If ZC_STATIC_MODE="ACTIVE" Then
				t=t & "&page=%n"
			End If
			If ZC_STATIC_MODE="REWRITE" Then
				t=Replace(t,".html","_%n.html")
			End If
			If ZC_STATIC_MODE="MIX" Then
				t=MixUrl
				t=t & "&page=%n"
			End If
		End If

		If intAllPage>0 Then
			Dim a,b

			If intNowPage=1 Then
				
				Template_PageBar_Previous=""
				strPageBarF="<a title="""& BetterPagebar_FristPage_Tip &""" href="""& Replace(t,"%n",1) &"""><span class=""page first-page"">"& BetterPagebar_FristPage &"</span></a><span class=""extend"">"& BetterPagebar_Extend &"</span><a title="""& BetterPagebar_PrvePage_Tip &""" href="""& Replace(t,"%n",intNowPage)  &"""><span class=""page pagebar-previous"">"& BetterPagebar_PrvePage &"</span></a>"
			Else
				Template_PageBar_Previous="<span class=""page pagebar-previous""><a href="""& Replace(t,"%n",intNowPage-1) &"""><span>"&ZC_MSG156&"</span></a></span>"
				strPageBarF="<a title="""& BetterPagebar_FristPage_Tip &""" href="""& Replace(t,"%n",1) &"""><span class=""page first-page"">"& BetterPagebar_FristPage &"</span></a><span class=""extend"">"& BetterPagebar_Extend &"</span><a title="""& BetterPagebar_PrvePage_Tip &""" href="""& Replace(t,"%n",intNowPage-1) &"""><span class=""page pagebar-previous"">"& BetterPagebar_PrvePage &"</span></a>"
			End If

			If intNowPage=intAllPage Then
				Template_PageBar_Next=""
				strPageBarL="<a title="""& BetterPagebar_NextPage_Tip &""" href="""& Replace(t,"%n",intNowPage) &"""><span class=""page pagebar-next"">"& BetterPagebar_NextPage &"</span></a><span class=""extend"">"& BetterPagebar_Extend &"</span><a title="""& BetterPagebar_LastPage_Tip &""" href="""& Replace(t,"%n",intAllPage) &"""><span class=""page last-page"">"& BetterPagebar_LastPage &"</span></a>"
			Else
				Template_PageBar_Next="<a href="""&Replace(t,"%n",intNowPage+1)  &"""><span class=""page pagebar-next"">"&ZC_MSG155&"</span></a>"
				strPageBarL="<a title="""& BetterPagebar_NextPage_Tip &""" href="""& Replace(t,"%n",intNowPage+1) &"""><span class=""page pagebar-next"">"& BetterPagebar_NextPage &"</span></a><span class=""extend"">"& BetterPagebar_Extend &"</span><a title="""& BetterPagebar_LastPage_Tip &""" href="""&Replace(t,"%n",intAllPage)& """><span class=""page last-page"">"& BetterPagebar_LastPage &"</span></a>"
			End If

			If intAllPage>ZC_PAGEBAR_COUNT Then
				a=intNowPage-Cint((ZC_PAGEBAR_COUNT-1)/2)
				b=intNowPage+ZC_PAGEBAR_COUNT-Cint((ZC_PAGEBAR_COUNT-1)/2)-1
				If a<=1 Then 
					a=1:b=ZC_PAGEBAR_COUNT
					If Not BetterPagebar_AlwaysShow Then strPageBarF=""
				End If
				If b>=intAllPage Then 
					b=intAllPage:a=intAllPage-ZC_PAGEBAR_COUNT+1
					If Not BetterPagebar_AlwaysShow Then strPageBarL=""
				End If
			Else
				a=1:b=intAllPage
				If Not BetterPagebar_AlwaysShow Then strPageBarF="":strPageBarL=""
			End If
			For i=a to b
				s=Replace(t,"%n",i)
				If ListType="DEFAULT" And i=1 Then s=ZC_BLOG_HOST
				If (ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS") And i=1 Then s=Url

				strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
				If i=intNowPage then
					Template_PageBar=Template_PageBar & "<span class=""now-page"">" & i & "</span>"
				Else
					strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
					strPageBar=Replace(strPageBar,"<#pagebar/page/number#>","<span class=""page"">"&i&"</span>")
					Template_PageBar=Template_PageBar & strPageBar
				End If

			Next


			Template_PageBar=strPageBarF & Template_PageBar & strPageBarL



		End If
		
	BetterPagebar_ExportBar=True

End Function
'*********************************************************

%>