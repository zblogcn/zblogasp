<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    PageMeta - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
Dim PageMeta_Meta,PageMeta_ArticleList,PageMeta_Value

'注册插件
Call RegisterPlugin("PageMeta","ActivePlugin_PageMeta")

'挂口部分
Function ActivePlugin_PageMeta()
	Call Add_Action_Plugin("Action_Plugin_TArticle_Export_Begin","PageMeta_Meta=Meta.GetValue(""pagemeta"")")
	Call Add_Filter_Plugin("Filter_Plugin_TArticleList_ExportByMixed","PageMeta_GetMeta2")
	Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_Template","PageMeta_P")
	Call Add_Action_Plugin("Action_Plugin_Edit_ueditor_getArticleInfo","Set PageMeta_ArticleList=EditArticle:Call PageMeta_addForm()")
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_Template","PageMeta_AddMeta")
	
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","PageMeta_Value=Request.Form(""PageMeta_ta""):Call Add_Filter_Plugin(""Filter_Plugin_PostArticle_Core"",""PageMeta_ArticlePst"")")


End Function
Function PageMeta_ArticlePst(Core)
	Core.Meta.setvalue "pagemeta",pAgEmEtA_EsCaPe_(PageMeta_Value)
End Function

Function PageMeta_addForm()

Dim ResponseText,TextStart,TextEnd,Text1,Text2
	Text1="<textarea id='PageMeta_ta' style='width:100%' name='PageMeta_ta'>" & TransferHTML(PageMeta_ArticleList.Meta.GetValue("pagemeta"),"[no-html]") &"</textarea>"
	TextStart="<span class='editinputname' style=""cursor:pointer;"" onClick=""$(this).next().toggleClass('hidden');"">PageMeta</span>"
	TextEnd="</span>"
	ResponseText=TextStart&Text1&TextEnd
	Call Add_Response_Plugin("Response_Plugin_Edit_Form3",ResponseText)
	set PageMeta_ArticleList=nothing
End Function

Function PageMeta_AddMeta(ByRef Ftemplate,Template_Article_Single,Template_Article_Multi, Template_Article_Istop)
		Ftemplate=PageMeta_P(Ftemplate)
End Function	

Function PageMeta_GetMeta2(i,b,d,g,e,h)
	Dim a,c,f
	If Not IsEmpty(b) Then
		Set c=New TCategory
		c.loadinfobyid b
		PageMeta_Meta=c.meta.getValue("pagemeta")&vbcrlf
	End If
	set c=nothing
	If Not IsEmpty(d) Then
		Set c=New TUser
		c.loadinfobyid d
		PageMeta_Meta=c.meta.getValue("pagemeta")&vbcrlf
	End If
	set c=nothing
	If Not IsEmpty(e) Then
		Call GetTagsbyTagNameList(e)
		for each c in tags
			If IsObject(c) Then
				PageMeta_Meta=c.meta.getValue("pagemeta")&vbcrlf
			end if
		next
	End If
	Set c=nothing
End Function


Function PageMeta_P(Ftemplate)
	'Response.Write FTemplate
	'r''esponse.end
	on error resume next
	Dim c,d,e
	c=vbsunescape(PageMeta_Meta)
	If c<>"" Then
		Dim a 
		Set a=New RegExp
		a.pattern="</head>"
		if a.test(FTemplate) Then
			d=split(c,vbcrlf)
			for e=0 to ubound(d)
				if trim(d(e))<>"" and instr(d(e),"---") then
					FTemplate=a.Replace(ftemplate,"<meta name="""&split(d(e),"---")(0)&""" content="""&split(d(e),"---")(1)&""" />"&vbcrlf&"</head>")
				end if
			next
		eNd iF
		set a=nothing
	End If
	PageMeta_P=FTemplate
End Function

%>
<script type="text/javascript"  language="javascript" runat="server">
function pAgEmEtA_EsCaPe_(str){	return escape(str)}
</script>