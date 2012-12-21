<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
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
<%
BlogTitle="应用中心"
Dim intHighlight
intHighlight=0
If Request.QueryString("action")="update" Then intHighlight=3
Dim objXmlHttp,strURL,bolPost,str,bolIsBinary,strList,bolFrame,strWrite
strWrite=""
bolFrame=True
bolPost=IIf(Request.ServerVariables("REQUEST_METHOD")="POST",True,False)

Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")

Select Case Request.QueryString("action")
	Case "view"
		strURL="view.asp?"
	Case "catalog"
		strURL="catalog.asp?"
	Case "app"
		strURL="app.asp?"
	Case "vaildcode"
		Response.ContentType="image/gif"
		strURL="zb_system/function/c_validcode.asp?"
		bolIsBinary=True
		bolFrame=False
	Case "cmd"
		strURL="zb_system/cmd.asp?"
		bolFrame=False
	Case "install"
		Response.Redirect "app_download.asp?url=" & Server.URLEncode(Request.QueryString("path"))
	Case "update"
		intHighlight=3
		strList=LoadFromFile(BlogPath&"zb_users\cache\appcentre_list.lst","utf-8")
		If Replace(strList,",","")<>"" Then
			strURL="app.asp?act=checkupdate&updatelist="&Server.URLEncode(strList)&"&"
			
		Else
			Call SetBlogHint_Custom("您没有可以更新的应用")
			strURL="?"
		End If
		Call DelToFile(BlogPath&"zb_users\cache\appcentre_list.lst")
		
	Case Else
		strURL="?"
End Select


'On Error Resume Next
Randomize
strURL=strURL & Request.QueryString & "&rnd="&Rnd

strURL=APPCENTRE_URL & strURL
objXmlHttp.Open Request.ServerVariables("REQUEST_METHOD"),strURL
If bolPost Then objXmlhttp.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"

objXmlHttp.Send Request.Form.Item


If objXmlHttp.ReadyState=4 Then
	If objXmlhttp.Status=200 Then
		If bolIsBinary=False Then
			Dim strResponse
			strResponse=objXmlhttp.ResponseText
			
			strResponse=Replace(strResponse,"$bloghost$",BlogHost)
			strResponse=Replace(strResponse,"$pluginlist$",ZC_USING_PLUGIN_LIST)
			strResponse=Replace(strResponse,"$zbversion$",BlogVersion)
			
			
			
			strResponse=Replace(strResponse,"catalog.asp?","server.asp?action=catalog&")
			strResponse=Replace(strResponse,APPCENTRE_URL&"app.asp?","server.asp?action=app&")
			strResponse=Replace(strResponse,APPCENTRE_URL&"app.asp","server.asp?action=app&")
			strResponse=Replace(strResponse,APPCENTRE_URL&"view.asp?","server.asp?action=view&")
			strResponse=Replace(strResponse,APPCENTRE_URL&"""","server.asp""")
			
			strResponse=Replace(strResponse,APPCENTRE_URL&"zb_system/function/c_validcode.asp?name=commentvalid","server.asp?action=vaildcode")
			strResponse=Replace(strResponse,APPCENTRE_URL&"zb_system/cmd.asp?","server.asp?action=cmd&")
			Dim objRegExp
			Set objRegExp=New RegExp
			'objRegExp.Pattern="<div class=""menu"">([\d\D]+?)</div>"
			'objRegExp.IgnoreCase=True
			'strResponse=objRegExp.Replace(strResponse,"<div class=""menu""><ul><li class=""index""><a href=""../../../zb_system/cmd.asp?act=login"">返回后台</a></li><li><a class=""on"" href=""server.asp"">应用中心</a></li><li><a href=""http://bbs.rainbowsoft.org"" target=""_blank"">Z-Blogger社区</a></li></ul></div>")
			objRegExp.Pattern="<!--client_begin([\d\D]+?)-->"
			objRegExp.Global=True
			strResponse=objRegExp.Replace(strResponse,"$1")
			objRegExp.Pattern="<!--server_begin-->([\d\D]+?)<!--server_end-->"
			strResponse=objRegExp.Replace(strResponse,"")
		Else
			Response.BinaryWrite objXmlHttp.ResponseBody
		End If
	Else
		ShowErr True,"" 
	End If
	'If objXmlHttp.GetRequestHeader("app_zbver")
	
	
Else
	ShowErr True,"" 
End If
If Err.Number<>0 Then ShowErr True,"" 



Function AddHtml(html,stat)
	Select Case stat
	Case 0
		strResponse=Replace(strResponse,"</head>",html&"</head>")
	Case 1
		strResponse=Replace(strResponse,"</body>",html&"</body>")
	Case 2
		strResponse=Replace(strResponse,"<head>","<head>"&html)
	Case 3
		strResponse=Replace(strResponse,"<body>","<body>"&html)
	End Select
End Function


%>
<%
If bolFrame Then%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<%
Dim aryTest
aryTest=Split(Split(strResponse,"</head>")(0),"<head>")
Response.Write aryTest(Ubound(aryTest))
%>

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader">应用中心</div>
          <div class="SubMenu">
            <%SubMenu(intHighlight)%>
          </div>
          <div id="divMain2"> 
<%=strWrite%>
<%
End If
aryTest=Split(Split(strResponse,"</body>")(0),"<body>")
Response.Write aryTest(Ubound(aryTest))
Function ShowErr(isHttp,str)
%>

          <%If isHttp Then%>
            <p>处理<a href='<%=strURL%>' target='_blank'><%=strURL%></a>(method:<%=Request.ServerVariables("REQUEST_METHOD")%>)时出错：</p>
            <p>ASP错误信息：<%=IIf(Err.Number=0,"无",Err.Number&"("&Err.Description&")")%></p>
            <p>HTTP状态码：<%If objXmlhttp.readyState<4 Then Response.Write "未发送请求" Else Response.Write objXmlhttp.status%></p>
            <p>&nbsp;</p>
            <p>可能的原因有：</p>
            <p>
            <ol>
              <li>您的服务器不允许通过HTTP协议连接到：<a href="<%=APPCENTRE_URL%>" target="_blank"><%=APPCENTRE_URL%></a>；</li>
              <li>您进行了一个错误的请求；</li>
              <li>服务器暂时无法连接，可能是遭到攻击或者检修中。</li>
            </ol>
            <p>请<a href="javascript:location.reload()">点击这里刷新重试</a>，或者到<a href="http://bbs.rainbowsoft.org" target="_blank">Z-Blogger论坛</a>发帖询问。</p>
            <%Else%>
            <%=str%>
          <%End If%>
<%
	Response.End
End Function
%>
<%If bolFrame Then%>
          </div>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script> 
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%End If%>