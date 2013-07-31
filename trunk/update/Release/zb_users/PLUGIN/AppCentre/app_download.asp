<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../zb_system/admin/ueditor/asp/aspincludefile.asp"-->
<!-- #include file="function.asp"-->
<%

Dim installStep,installPath,bolBatch,strRnd
Randomize
strRnd=MD5(Int(Rnd*10000000))
bolBatch=CBool(Request.QueryString("batch"))
installStep=Request.QueryString("step")
If Not IsNumeric(installStep) Then installStep=0
Pack_For=""
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)

Select Case installStep
	Case 1
		Call Step1_DownloadFromUrl(Request.QueryString("url"))
	Case 2
		result.success=Step2_InstallApp(Request.QueryString("path"))
	Case 0
		Call Step1_DownloadFromUrl(Request.QueryString("url"))
		If result.success Then
			result.success=Step2_InstallApp(strRnd)
		End If
End Select
Response.Write toJson()

	



Function Step1_DownloadFromUrl(strUrl)
	If Left(strURL,Len(APPCENTRE_URL))=APPCENTRE_URL Then 
		Dim objXmlHttp
		Set objXmlHttp=Server.CreateObject("msxml2.serverxmlhttp")
		objXmlhttp.Open "GET",strURL & "?" & strRnd
		objXmlHttp.Send
		If objXmlHttp.ReadyState=4 Then
			If objXmlHttp.Status=200 Then
				Call SaveBinary(objXmlhttp.ResponseBody,BlogPath&"zb_users\cache\temp_" & strRnd & ".zba")
				result.path=strRnd
			Else
				result.success=False
				result.errmsg="下载出错："&objXmlHttp.Status
			End If
		Else
			result.success=False
			result.errmsg="下载出错"
		End If
	Else
		result.success=False
		result.errmsg="非法地址"
	End If
End Function

Function Step2_InstallApp(strRnd)
	If InstallApp(BlogPath&"zb_users\cache\temp_" & strRnd & ".zba") Then
		Call DelToFile(BlogPath&"zb_users\cache\temp_" & strRnd & ".zba")
		Step2_InstallApp=True
	Else
		Step2_InstallApp=False
		result.errmsg="无法解压文件"
	End If
End Function

'Function Step3_

%>
<script language="javascript" runat="server">
var result={
	"success":true,
	"errmsg":"",
	"path":""
}
function toJson(){
	var str="{";
	for(attr in result){
		str+="\""+attr+"\":\""+result[attr]+"\","
	}
	str+="\"antierr\":\"err\"}"
	return str;
}
</script>