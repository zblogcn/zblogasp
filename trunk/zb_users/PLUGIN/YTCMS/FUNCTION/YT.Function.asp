<!--#include file="../LIB/YT.Data.asp" -->
<!--#include file="../LIB/YT.Expressions.asp" -->
<!--#include file="../LIB/YT.Template.asp" -->
<!--#include file="../LIB/XML/YT.Model.asp" -->
<!--#include file="../LIB/XML/YT.Block.asp" -->
<!--#include file="../LIB/XML/YT.TPL.asp" -->
<%
Dim YT_CMS_Global_Article
Dim YT_CMS_Global_Bool
Dim YT_CMS_Global_Numeric

Function YT_CMS_Filter_Plugin_TArticle_LoadInfobyID(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)
	YT_CMS_Global_Numeric=CateID
End Function

Function YT_CMS_Filter_Plugin_TArticle_LoadInfoByArray(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)
	YT_CMS_Global_Numeric=CateID
End Function
Sub YT_CMS_Filter_Plugin_TArticle_Build_Template(ByRef html)
	If Not IsEmpty(html) Then html = new YT_Template.AnalysisTab(html)
	'Call SaveToFile(BlogPath & "debug_YT_CMS_Filter_Plugin_TArticle_Build_Template.txt",html,"utf-8",False)
End Sub
Function Filter_Plugin_TArticleList_Build_Template(ByRef html)
	If Not IsEmpty(html) Then html = new YT_Template.AnalysisTab(html)
End Function
Sub YT_CMS_Filter_Plugin_TArticle_Export_Template(ByRef html,ByRef Template_Article_Single,ByRef Template_Article_Multi,ByRef Template_Article_Istop)
	Dim YT_CMS_Template_Article_Single,YT_CMS_Template_Article_Multi
	If Not IsEmpty(html) Then html = new YT_Template.AnalysisTab(html)
	'If Not IsEmpty(Template_Article_Istop) Then Template_Article_Istop = new YT_Template.AnalysisTab(Template_Article_Istop)
	Call new YT_TPL_XML.Load()
	YT_CMS_Template_Article_Single=YT_Template_GetConfigBlock(YT_CMS_Global_Numeric,YTConfig.Single)
	YT_CMS_Template_Article_Multi=YT_Template_GetConfigBlock(YT_CMS_Global_Numeric,YTConfig.Multi)
	If Not IsEmpty(YT_CMS_Template_Article_Single) Then Template_Article_Single=YT_CMS_Template_Article_Single
	If Not IsEmpty(YT_CMS_Template_Article_Multi) Then Template_Article_Multi=YT_CMS_Template_Article_Multi
	'If Not IsEmpty(Template_Article_Single) Then Template_Article_Single = new YT_Template.AnalysisTab(Template_Article_Single)
	'If Not IsEmpty(Template_Article_Multi) Then Template_Article_Multi = new YT_Template.AnalysisTab(Template_Article_Multi)
	'Call SaveToFile(BlogPath & "debug_YT_CMS_Filter_Plugin_TArticle_Export_Template.txt",Template_Article_Multi,"utf-8",False)
End Sub
Sub YT_CMS_Filter_Plugin_TArticle_Export_Template_Sub(ByRef Template_Article_Comment,ByRef Template_Article_Trackback,ByRef Template_Article_Tag,ByRef Template_Article_Commentpost,ByRef Template_Article_Navbar_L,ByRef Template_Article_Navbar_R,ByRef Template_Article_Mutuality)
	'If Not IsEmpty(Template_Article_Comment) Then Template_Article_Comment = new YT_Template.AnalysisTab(Template_Article_Comment)
	'If Not IsEmpty(Template_Article_Tag) Then Template_Article_Tag = new YT_Template.AnalysisTab(Template_Article_Tag)
	'If Not IsEmpty(Template_Article_Commentpost) Then Template_Article_Commentpost = new YT_Template.AnalysisTab(Template_Article_Commentpost)
	'If Not IsEmpty(Template_Article_Navbar_L) Then Template_Article_Navbar_L = new YT_Template.AnalysisTab(Template_Article_Navbar_L)
	'If Not IsEmpty(Template_Article_Navbar_R) Then Template_Article_Navbar_R = new YT_Template.AnalysisTab(Template_Article_Navbar_R)
	'If Not IsEmpty(Template_Article_Mutuality) Then Template_Article_Mutuality = new YT_Template.AnalysisTab(Template_Article_Mutuality)
End Sub
Sub YT_CMS_Filter_Plugin_TArticle_Del(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)
	Dim YTModelXML,Node,Sql
	Set YTModelXML = new YT_Model_XML
	Set Node = YTModelXML.GetModel(CateID)
		If Not Node Is Nothing Then
			Sql = "DELETE FROM ["&Node.selectSingleNode("Table/Name").Text&"] WHERE [log_ID] = "&ID
			objConn.Execute(Sql)
		End If
	Set Node = Nothing
	Set YTModelXML = Nothing
End Sub
Sub YT_CMS_Filter_Plugin_TArticle_Export_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)
	Dim j,YTModelXML,Node,Object,Field
		Set YTModelXML = new YT_Model_XML
		Set Node = YTModelXML.GetModel(aryTemplateTagsValue(12))
			If Not Node Is Nothing Then
				Dim Json
					Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,aryTemplateTagsValue(1))
					If IsEmpty(Json) Then Exit Sub
					Set Object = jsonToObject(Json)
					For Each Field In Object
						j = Ubound(aryTemplateTagsName) + 1
						ReDim Preserve aryTemplateTagsName(j)
						ReDim Preserve aryTemplateTagsValue(j)
						aryTemplateTagsName(j) = "article/model/"&Field.Name
						aryTemplateTagsValue(j) = jsUnEscape(Field.Value)
					Next
					Set Object = Nothing
			End If
		Set Node = Nothing
		Set YTModelXML = Nothing
End Sub
Sub YT_CMS_Filter_Plugin_TArticle_Build_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)
	Dim l
	For l = LBound(aryTemplateTagsName) To UBound(aryTemplateTagsName)
		If InStr(aryTemplateTagsName(l),"TEMPLATE_INCLUDE_") Then
			aryTemplateTagsValue(l) = new YT_Template.AnalysisTab(aryTemplateTagsValue(l))
		End If
	Next
End Sub
Sub YT_CMS_Filter_Plugin_PostArticle_Core(ByRef objArticle)
	Set YT_CMS_Global_Article = objArticle
	If  objArticle.ID = 0 Then YT_CMS_Global_Bool = True
End Sub
Sub YT_CMS_Action_Plugin_ArticlePst_Succeed()
	If YT_CMS_Global_Bool Then
		Dim Sql:Sql = YT_Data_GetSql(YT_CMS_Global_Article.CateID,YT_CMS_Global_Article.ID)
		If YT_CMS_Global_Article.ID > 0 And Not IsEmpty(Sql) Then
			objConn.Execute(Sql)
			Call ScanTagCount(YT_CMS_Global_Article.Tag)
			Call BuildArticle(YT_CMS_Global_Article.ID,True,True)
		End If
	End If
End Sub
Sub YT_CMS_Filter_Plugin_TArticle_Post(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)
	If ID > 0 Then
		Dim Sql:Sql=YT_Data_GetSql(CateID,ID)
		If Not IsEmpty(Sql) Then objConn.Execute(Sql)
	End If
End Sub
Sub YT_CMS_Action_Plugin_MakeBlogReBuild_Core_Begin()
	Call new YT_Block_XML.Build()
End Sub
Function YT_Template_GetConfigBlock(CateID,Block)
	Dim Object,Cate
	Dim isCate:isCate = False
	For Each Object In Block
		If IsObject(Object) Then
			For Each Cate In Object.v
				If Cate = CateID Then
					isCate = True
				Exit For
			End If
			Next
		End If
		If isCate Then
			YT_Template_GetConfigBlock = new YT_Template.GetFile(Object.t)
			Exit For
		End If
	Next
End Function
Function YT_Model_Analysis()
	Dim Script
		Script = Script & "<p id=""model"">"
        Script = Script & "<script>var ZC_BLOG_HOST='"&ZC_BLOG_HOST&"';var YT_CMS_XML_URL='"&ZC_BLOG_HOST&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"&"';</script>"
		Script = Script & "<script src=""../../ZB_USERS/PLUGIN/YTCMS/Config.js""></script>"
		Script = Script & "<script src=""../../ZB_USERS/PLUGIN/YTCMS/SCRIPT/YT.Lib.js""></script>"
		Script = Script & "<script>$(document).ready(function(){YT.Panel.Analysis();});</script>"
		Script = Script & "</p>"
		YT_Model_Analysis = Script
End Function
Function YT_Data_GetRow(TableName,ID)
	If Not new YT_Table.Exist(TableName) Then Exit Function
	Dim R(),Rs,j,Field
	Set Rs = objConn.Execute("SELECT TOP 1 * FROM ["&TableName&"] WHERE [log_ID]="&ID)
		If Not (Rs.EOF and Rs.BOF) Then
			Redim R(-1)
			For Each Field In Rs.Fields
				j = Ubound(R) + 1
				ReDim Preserve R(j)
				R(j) = "{Name:"&Chr(34)&Field.Name&Chr(34)&",Value:"&Chr(34)&Rs(Field.Name)&Chr(34)&"}"
			Next
			YT_Data_GetRow = "["&Join(R,",")&"]"
		End If
	Set Rs = Nothing
End Function
Function YT_Data_GetSql(CateID,ID)
	Dim YTModelXML,Node,Sql,Field
	Dim FieldName(),FieldValue(),j
	Set YTModelXML = new YT_Model_XML
	Set Node = YTModelXML.GetModel(CateID)
		If Not Node Is Nothing Then
			Redim FieldName(0)
			Redim FieldValue(0)
			FieldName(0) = "[log_ID]"
			FieldValue(0) = ID
			For Each Field In Node.selectNodes("Field")
				j = Ubound(FieldName) + 1
				ReDim Preserve FieldName(j)
				ReDim Preserve FieldValue(j)
				FieldName(j) = "["&Field.selectSingleNode("Name").Text&"]"
				FieldValue(j) = "'"&jsEscape(TransferHTML(FilterSQL(Request.Form(Field.selectSingleNode("Name").Text)),"[anti-upload]"))&"'"
			Next
			If Not IsEmpty(YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,FieldValue(0))) Then
				Sql = "UPDATE ["&Node.selectSingleNode("Table/Name").Text&"] SET "
				For j = 1 To UBound(FieldName)
					If j > 1 Then Sql = Sql & ","
					Sql = Sql & FieldName(j)&"="&FieldValue(j)
				Next
				Sql = Sql & " WHERE "&FieldName(0)&"="&FieldValue(0)
			Else
				Sql = "INSERT INTO ["&Node.selectSingleNode("Table/Name").Text&"]("
				Sql = Sql & Join(FieldName,",") & ")VALUES ("&Join(FieldValue,",")&")"
			End If
			YT_Data_GetSql = Sql
		End If
	Set Node = Nothing
	Set YTModelXML = Nothing
End Function
%>