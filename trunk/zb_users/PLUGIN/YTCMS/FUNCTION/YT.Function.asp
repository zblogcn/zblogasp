<!--#include file="../LIB/YT.Data.asp" -->
<!--#include file="../LIB/YT.Expressions.asp" -->
<!--#include file="../LIB/YT.Template.asp" -->
<!--#include file="../LIB/XML/YT.Model.asp" -->
<!--#include file="../LIB/XML/YT.Block.asp" -->
<%
Sub YT_CMS_Filter_Plugin_TArticle_Build_Template(ByRef html)
	If Not IsEmpty(html) Then html = new YT_Template.AnalysisTab(html)
End Sub
Function YT_CMS_Filter_Plugin_TArticleList_Build_Template(ByRef html)
	If Not IsEmpty(html) Then html = new YT_Template.AnalysisTab(html)
End Function
Function YT_CMS_Filter_Plugin_TArticle_Del(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef FType,ByRef MetaString)
	Dim YTModelXML,Node,Sql
	Set YTModelXML = new YT_Model_XML
	Set Node = YTModelXML.GetModel(CateID)
		If Not Node Is Nothing Then
			Sql = "DELETE FROM ["&Node.selectSingleNode("Table/Name").Text&"] WHERE [log_ID] = "&ID
			objConn.Execute(Sql)
		End If
	Set Node = Nothing
	Set YTModelXML = Nothing
End Function
Function YT_CMS_Filter_Plugin_TArticle_Export_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)
	Dim j,YTModelXML,Node,Object,Field
		Set YTModelXML = new YT_Model_XML
		Set Node = YTModelXML.GetModel(aryTemplateTagsValue(12))
			If Not Node Is Nothing Then
				Dim Json
					Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,aryTemplateTagsValue(1))
					If IsEmpty(Json) Then Exit Function
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
End Function
Function YT_CMS_Filter_Plugin_TArticle_Build_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)
	Dim l
	For l = LBound(aryTemplateTagsName) To UBound(aryTemplateTagsName)
		If InStr(aryTemplateTagsName(l),"TEMPLATE_INCLUDE_") Then
			aryTemplateTagsValue(l) = new YT_Template.AnalysisTab(aryTemplateTagsValue(l))
		End If
	Next
End Function
Function YT_CMS_Filter_Plugin_PostArticle_Succeed(ByRef objArticle)
	Dim Sql:Sql = YT_Data_GetSql(objArticle.CateID,objArticle.ID)
	If objArticle.CateID > 0 And Not IsEmpty(Sql) Then objConn.Execute(Sql)
	Call BuildArticle(objArticle.ID,True,True)
End Function
Function YT_CMS_Action_Plugin_MakeBlogReBuild_Core_Begin()
	Call new YT_Block_XML.Build()
End Function
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
				If Field.selectSingleNode("Property").Text = "VARCHAR" Then
					FieldValue(j) = "'"&jsEscape(TransferHTML(FilterSQL(Request.Form(Field.selectSingleNode("Name").Text)),"[anti-upload]"))&"'"
				Else
					FieldValue(j) = "'"&TransferHTML(FilterSQL(Request.Form(Field.selectSingleNode("Name").Text)),"[anti-upload]")&"'"
				End If
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