<%@ CODEPAGE=65001 %>
<!-- #include file="../../../zb_users/c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_users/plugin/p_config.asp" -->
<%
Response.ContentType="application/x-javascript"
%>
<%
Call ActivePlugin()
For Each sAction_Plugin_UEditor_Config_Begin in Action_Plugin_UEditor_Config_Begin
	If Not IsEmpty(sAction_Plugin_UEditor_Config_Begin) Then Call Execute(sAction_Plugin_UEditor_Config_Begin)
Next

        
	Dim strUPLOADDIR

	strUPLOADDIR = Replace(ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())),"\","/")

	Dim Path
	Path=BlogHost & ""& strUPLOADDIR &"/"
	dim strJSContent
	strJSContent="(function(){var URL;URL = '"&BlogHost&"zb_system/admin/ueditor/'; window.UEDITOR_CONFIG = {UEDITOR_HOME_URL : URL,imageUrl:URL+""asp/picUp.asp"",imagePath:"""&Path&""" ,imageFieldName:""edtFileLoad"" ,fileUrl:URL+""asp/fileUp.asp"",filePath:"""&Path&""" ,fileFieldName:""edtFileLoad"",catchRemoteImageEnable:false,imageManagerUrl:URL +""asp/imageManager.asp"" ,imageManagerPath:"""&BlogHost&""",wordImageUrl:URL+""asp/picUp.asp"",scrawlUrl:URL+""asp/picUp.php"",scrawlPath:"""&path&""",wordImagePath:"""&Path&""",wordImageFieldName:""edtFileLoad"",getMovieUrl:URL+""asp/getMovie.asp"",toolbars:[ ['fullscreen', 'source', '|', 'undo', 'redo', '|', 'bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript','|',  'forecolor', 'backcolor', 'insertorderedlist', 'insertunorderedlist','|', 'indent', '|', 'justifyleft', 'justifycenter', 'justifyright', 'justifyjustify',  '|',  'removeformat','autotypeset', 'searchreplace'],[ 'fontfamily', 'fontsize', '|', 'emotion','link','insertimage', 'insertvideo', 'attachment','spechars','|', 'map', 'gmap', '|', 'highlightcode','blockquote', 'pasteplain','wordimage','|','inserttable', 'deletetable', '|','preview']]};})();"


Call Filter_Plugin_UEditor_Config(strJSContent)

For Each sAction_Plugin_UEditor_Config_End in Action_Plugin_UEditor_Config_End
	If Not IsEmpty(sAction_Plugin_UEditor_Config_End) Then Call Execute(sAction_Plugin_UEditor_Config_End)
Next

	response.write strJSContent

%>