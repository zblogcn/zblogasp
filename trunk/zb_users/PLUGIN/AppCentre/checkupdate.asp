<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'未完成功能：
'1.得到需要更新的主题（tname）
'2.用XML来判断是否有app子节点
'3.显示更新列表
'4.下载更新
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp"-->
<%


Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AppCentre")=False Then Call ShowError(48)
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader">应用中心</div>
          <div class="SubMenu">
            <%SubMenu(3)%>
          </div>
          <div id="divMain2">
            <%
Dim bolSilent,bolReDownload,bolCheck,objXml,objXml2,objChildXml,objAppXml,i,objFso

Dim strTemp,strName,strType,dtmModified,dtmLocalModified
Dim aryDownload(),aryName()
Redim aryDownload(0)
Redim aryName(0)

Select Case Request.QueryString("act")
	Case "silent"
		bolSilent=True
		bolReDownload=True
		bolCheck=True
	Case "recheck"
		bolReDownload=True
		bolCheck=True
	Case Else
		Call CheckXml
		Response.Redirect "server.asp?action=update"
End Select

If bolReDownload Then Call ReCheck
If bolCheck Then Call CheckXml
If bolSilent Then Response.End


Function CheckXml()
	Set objXml=Server.CreateObject("Microsoft.XMLDOM")
	objXml.Load BlogPath&"zb_users\cache\appcentre_plugin.xml"
	If objXml.ReadyState=4 Then
		'这里该显示更新列表了
		If objXml.parseError.errorCode=0 Then
			Set objChildXml=objXml.selectNodes("//apps/app")
			For i=0 To objChildXml.length-1
				Set objAppXml=objChildXml(i)'
				If CLng(objAppXml.getAttributeNode("zbversion").value)<=BlogVersion Then
					strName=objAppXml.getAttributeNode("name").value
					strType=objAppXml.getAttributeNode("type").value
					dtmModified=CDate(objAppXml.getAttributeNode("modified").value)
					Set objXml2=Server.CreateObject("Microsoft.XMLDOM")
					objXml2.Load BlogPath&"zb_users\"&strType&"\"&strName&"\plugin.xml"
					If objXml2.ReadyState=4 Then
						If objXml2.parseError.errorCode=0 Then
							dtmLocalModified=CDate(objXml2.documentElement.selectSingleNode("modified").text)
						End If
					End If
					If DateDiff("d","1970-1-1 08:00",dtmModified)>DateDiff("d","1970-1-1 08:00",dtmLocalModified) Then
						Redim Preserve aryDownload(Ubound(aryDownload)+1)
						Redim Preserve aryName(Ubound(aryName)+1)
						aryDownload(Ubound(aryDownload))=objAppXml.getAttributeNode("url").value
						aryName(Ubound(aryName))=strName
					End If
				End If
			Next
		End If
	End If
	For i=0 To Ubound(aryName)
		If Not bolSilent Then Response.Write aryName(i) & "------" & aryDownload(i)
		CheckXml=CheckXml & "," & aryName(i)
	Next
	Call SaveToFile(BlogPath&"zb_users\cache\appcentre_list.lst",CheckXml,"utf-8",False)

End Function


Function GetAllThemeName
	Dim aryReturn()
	Redim aryReturn(0)
	Dim fso,f,fc,f1
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "zb_users/theme" & "/")
	Set fc = f.SubFolders
	For Each f1 In fc
		Redim Preserve aryReturn(Ubound(aryReturn)+1)
		aryReturn(Ubound(aryReturn))=f1.Name
	Next
	GetAllThemeName=aryReturn
End Function

Function ReCheck()
	Dim objXmlHttp,strURL,bolPost,str,bolIsBinary
	Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")

	strUrl=APPCENTRE_UPDATE_URL&"&tname="&Server.URLEncode(Join(GetAllThemeName,","))&"&pname="&Server.URLEncode(Replace(ZC_USING_PLUGIN_LIST,"|",","))
	objXmlHttp.Open "GET",strURL
	objXmlHttp.Send 

	If objXmlHttp.ReadyState=4 Then
		If objXmlhttp.Status=200 Then
		Else
			ShowErr
		End If
		
		
	Else
		ShowErr
	End If
	If Err.Number<>0 Then ShowErr
	'Response.Write strUrl'objXmlHttp.ResponseText
	
	'这里应该用XML来判断是否有app子节点
	Call SaveToFile(BlogPath&"zb_users\cache\appcentre_plugin.xml",objXmlHttp.ResponseText,"utf-8",False)
End Function

Function ShowErr()
%>
            <p>处理<a href='<%=strURL%>' target='_blank'><%=strURL%></a>(method:<%=Request.ServerVariables("REQUEST_METHOD")%>)时出错：</p>
            <p>ASP错误信息：<%=IIf(Err.Number=0,"无",Err.Number&"("&Err.Description&")")%></p>
            <p>HTTP状态码：
              <%If objXmlhttp.readyState<4 Then Response.Write "未发送请求" Else Response.Write objXmlhttp.status%>
            </p>
            <p>&nbsp;</p>
            <p>可能的原因有：</p>
            <p>
            
            <ol>
              <li>您的服务器不允许通过HTTP协议连接到：<a href="<%=APPCENTRE_URL%>" target="_blank"><%=APPCENTRE_URL%></a>；</li>
              <li>您进行了一个错误的请求；</li>
              <li>服务器暂时无法连接，可能是遭到攻击或者检修中。</li>
            </ol>
            <p>请<a href="javascript:location.reload()">点击这里刷新重试</a>，或者到<a href="http://bbs.rainbowsoft.org" target="_blank">Z-Blogger论坛</a>发帖询问。</p>
            <%
	Response.End
End Function
%>
          </div>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script> 
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->