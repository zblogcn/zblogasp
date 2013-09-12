<!--#include file="../LIB/YT.Data.asp" -->
<!--#include file="../LIB/YT.TPL.asp" -->
<!--#include file="../LIB/XML/YT.Model.asp" -->
<!--#include file="../LIB/XML/YT.Block.asp" -->
<%
Dim GLOBE_MODEL

Function YT_CMS_Filter_Plugin_TArticle_Build_Template_Succeed(ByRef html)
	If Not IsEmpty(html) Then html = YT_TPL_display(html)
End Function

Function YT_CMS_Filter_Plugin_TArticleList_Build_Template_Succeed(ByRef html)
	If Not IsEmpty(html) Then html = YT_TPL_display(array(html,GLOBE_MODEL))
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

function YT_CMS_Action_Plugin_TArticleList_Export_Begin(byref intPage,byref anyCate,byref anyAuthor,byref dtmDate,byref anyTag)
	dim s
	s = array(intPage,anyCate,anyAuthor,dtmDate,anyTag)
	if len(join(s,"")) > 0 then
		if len(s(1)) > 0 then GLOBE_MODEL = "CATEGORY-"&s(0)&"-"&s(1)
		if len(s(2)) > 0 then GLOBE_MODEL = "USER-"&s(0)&"-"&s(2)
		if len(s(3)) > 0 then GLOBE_MODEL = "DATE-"&s(0)&"-"&s(3)
		if len(s(4)) > 0 then GLOBE_MODEL = "TAGS-"&s(0)&"-"&s(4)
	else
		GLOBE_MODEL = "DEFAULT"
	end if
end function

Function YT_CMS_Action_Plugin_TArticle_Export_End(byref html,byref subhtml,byref aryTemplateTagsName,byref aryTemplateTagsValue,byref subhtml_TemplateName)
	if GLOBE_MODEL<>"DEFAULT" then
		Dim j,mx,Node,Object,Field,os
			Execute("aryName = aryTemplateTagsName")
			Execute("aryValue = aryTemplateTagsValue")
			Set mx = new YT_Model_XML
			Set Node = mx.GetModel(aryTemplateTagsValue(12))
				If Not Node Is Nothing Then
					Dim Json,YTARRAY
						Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,aryTemplateTagsValue(1))
						Json = Trim(Json)
						If Len(Json) = 0 Then Exit Function
						Set Object = YT.eval(Json)
						Execute("YTARRAY=Object.YTARRAY")
						For Each Field In Split(YTARRAY,",")
							Execute(Field&" = YT.unescape(Object."&Field&")")
						Next
						os = subhtml
						subhtml = YT_TPL_display(subhtml)
						html = Replace(html,os,subhtml)
						For Each Field In Split(YTARRAY,",")
							Execute(Field&" = Empty")
						Next
						Set Object = Nothing
				End If
			Set Node = Nothing
			Set mx = Nothing
	end if
End Function

Function YT_CMS_Filter_Plugin_PostArticle_Succeed(ByRef objArticle)
	Dim Sql:Sql = YT_Data_GetSql(objArticle.CateID,objArticle.ID)
	If objArticle.CateID > 0 And Not IsEmpty(Sql) Then objConn.Execute(Sql)
	Call BuildArticle(objArticle.ID,True,True)
End Function

Function YT_CMS_Action_Plugin_MakeBlogReBuild_Core_Begin()
	Call new YT_Block_XML.Build()
End Function

Function YT_Model_Analysis()
	Dim Script
		Script = Script & "<p id=""model"">"
        Script = Script & "<script>var ZC_BLOG_HOST='"&ZC_BLOG_HOST&"';var ZC_BLOG_THEME = '"&ZC_BLOG_THEME&"';var YT_CMS_XML_URL=ZC_BLOG_HOST+'ZB_USERS/THEME/'+ZC_BLOG_THEME+'/';</script>"
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
			YT_Data_GetRow = "{"&Join(R,",")&",""YTARRAY"":["&Join(F,",")&"]}"
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

Function YT_DelFile(Byval fileName)
	Call ClearGlobeCache()
	Dim themesDir
	themesDir="ZB_USERS/THEME"&"/"&ZC_BLOG_THEME&"/"&ZC_TEMPLATE_DIRECTORY
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(BlogPath&themesDir&"/"&fileName) Then
		fso.GetFile(BlogPath&themesDir&"/"&fileName).Delete
		YT_DelFile = True
	End If
	Set fso=Nothing
	Call LoadGlobeCache()
End Function

Function YT_TPL_display(block)
	Dim t
	Set t = New YT_TPL
		if isArray(block) then
			if UBound(block) > 0 then
				t.cache = block(1)
				t.template = block(0)
			end if
		else
			t.template = block
		end if
		YT_TPL_display = t.display()
	Set t = Nothing
End Function

Function getModel(CateID,ID)
	Dim x,n,s:s = Empty
	If isNumeric(CateID) And isNumeric(ID) Then
		Set x = new YT_Model_XML
		Set n = x.GetModel(CateID)
			If Not n Is Nothing Then
				s = YT_Data_GetRow(n.selectSingleNode("Table/Name").Text,ID)
			End If
		Set n = Nothing
		Set x = Nothing
	End If
	getModel = s
End Function

Function FilterSQL(strSQL)
	Dim s,t
	s = strSQL
	s = Trim(s)
	if len(s)=0 or isEmpty(s) or s = "" then exit function
	Set t = New YT_TPL
	s = CStr(Replace(s,chr(39),chr(39)&chr(39)))
	s = t.reg_replace("\<\!\-\-\{(.+?)\}\-\-\>","&lt;!--&#123;$1&#125;--&gt;",s)
	s = t.reg_replace("\{\$(.+?)\}", "&#123;&#36;$1&#125;",s)
	s = t.reg_replace("\{foreach\s+(.+?)\s+(.+?)\}", "&#123;foreach&nbsp;$1&nbsp;$2&#125;",s)
	s = t.reg_replace("\{for\s+(.+?)\s+(.+?)\}", "&#123;for&nbsp;$1&nbsp;$2&#125;",s)
	s = Replace(s,"{/next}","&#123;/next&#125;")
	s = t.reg_replace("\{(do|loop)\s+(while|until)\s+(.+?)\}","&#123;$1&nbsp;$2&nbsp;$3&#125;",s)
	s = Replace(s,"{do}","&#123;do&#125;")
	s = Replace(s,"{loop}","&#123;loop&#125;")
	s = t.reg_replace("\{while\s+(.+?)\}","&#123;while&nbsp;$1&#125;",s)
	s = Replace(s,"{/wend}","&#123;/wend&#125;")
	s = t.reg_replace("\{if\s+(.+?)\}","&#123;if&nbsp;$1&#125;",s)
	s = t.reg_replace("\{elseif\s+(.+?)\}","&#123;elseif&nbsp;$1&#125;",s)	
	s = Replace(s,"{/if}","&#123;/if&#125;")
	s = Replace(s,"{else}","&#123;else&#125;")
	s = Replace(s,"{code}","&#123;code&#125;")
	s = Replace(s,"{/code}","&#123;/code&#125;")
	s = t.reg_replace("\{eval\s+(.+?)\}","&#123;eval&nbsp;$1&#125;",s)
	s = t.reg_replace("\{echo\s+(.+?)\}","&#123;echo&nbsp;$1&#125;",s)
	FilterSQL = s
	Set t = Nothing
End Function

'重写ZBLOG系统函数
Function BlogReBuild_Default()
	Call ClearGlobeCache()
	Call LoadGlobeCache()
	Call DelToFile(BlogPath & "zb_users/CACHE/default.asp")
	Dim l,dir,i
	dir = "ZB_USERS/CACHE/"
	l = LoadIncludeFiles(dir)
	for each i in l
		if right(i,4) = ".TPL" then
			Call DelToFile(BlogPath & "zb_users/CACHE/" & i)
		end if
	next
End Function
%>