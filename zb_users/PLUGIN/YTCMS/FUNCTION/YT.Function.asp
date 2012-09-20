<!--#include file="../LIB/YT.Data.asp" -->
<!--#include file="../LIB/YT.TPL.asp" -->
<!--#include file="../LIB/XML/YT.Model.asp" -->
<!--#include file="../LIB/XML/YT.Block.asp" -->
<%
Function YT_CMS_Filter_Plugin_TArticle_Build_Template(ByRef html)
	If Not IsEmpty(html) Then Call YT_TPL_display(html)
End Function
Function YT_CMS_Filter_Plugin_TArticleList_Build_Template(ByRef html)
	If Not IsEmpty(html) Then Call YT_TPL_display(html)
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
					Set Object = YT.eval(Json)
					For Each Field In Object.YTARRAY
						j = Ubound(aryTemplateTagsName) + 1
						ReDim Preserve aryTemplateTagsName(j)
						ReDim Preserve aryTemplateTagsValue(j)
						Execute("aryTemplateTagsName(j) = ""article/model/"&Field&"""")
						Execute("aryTemplateTagsValue(j) = Replace(YT.unescape(Object."&Field&"),VBCRLF,CHR(32))")
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
			Call YT_TPL_display(aryTemplateTagsValue(l))
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
Function YT_TPL_display(Byref s)
	Dim t
	Set t = New YT_TPL
		t.template = s
		s = t.display()
	Set t = Nothing
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
	Dim R(),F(),Rs,j,Field
	Set Rs = objConn.Execute("SELECT TOP 1 * FROM ["&TableName&"] WHERE [log_ID]="&ID)
		If Not (Rs.EOF and Rs.BOF) Then
			Redim R(-1)
			For Each Field In Rs.Fields
				j = Ubound(R) + 1
				ReDim Preserve R(j)
				ReDim Preserve F(j)
				R(j) = Chr(34)&Field.Name&Chr(34)&":"&Chr(34)&Rs(Field.Name)&Chr(34)
				F(j) = Chr(34)&Field.Name&Chr(34)
			Next
			YT_Data_GetRow = "{"&Join(R,",")&",YTARRAY:["&Join(F,",")&"]}"
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
				If Field.selectSingleNode("Property").Text = "VARCHAR" Or Field.selectSingleNode("Property").Text = "TEXT" Then
					FieldValue(j) = "'"&YT.escape(TransferHTML(FilterSQL(Request.Form(Field.selectSingleNode("Name").Text)),"[anti-upload]"))&"'"
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
Function YT_FileJsonList()
	Dim aryFileList,themesDir
	themesDir="ZB_USERS/THEME"&"/"&ZC_BLOG_THEME&"/"&ZC_TEMPLATE_DIRECTORY
	aryFileList=LoadIncludeFiles(themesDir)
	Dim jsonText
	jsonText="["
	If IsArray(aryFileList) Then
		Dim j,i
		j=UBound(aryFileList)
		For i=1 to j
			If i<>1 Then jsonText=jsonText&","
			jsonText=jsonText&Chr(34)&aryFileList(i)&Chr(34)
		Next
	End If
	jsonText=jsonText&"]"
	YT_FileJsonList = jsonText
End Function
Function YT_GetFile(Byval fileName)
	Dim themesDir
	themesDir="ZB_USERS/THEME"&"/"&ZC_BLOG_THEME&"/"&ZC_TEMPLATE_DIRECTORY
	YT_GetFile=LoadFromFile(BlogPath&themesDir&"/"&fileName,"utf-8")
End Function
Function YT_SaveFile(Byval fileName,Byval fileContent)
	Call ClearGlobeCache()
	Dim themesDir
	themesDir="ZB_USERS/THEME"&"/"&ZC_BLOG_THEME&"/"&ZC_TEMPLATE_DIRECTORY
	Call SaveToFile(BlogPath&themesDir&"/"&fileName,fileContent,"utf-8",False)
	Call LoadGlobeCache()
	YT_SaveFile = True
End Function
%>