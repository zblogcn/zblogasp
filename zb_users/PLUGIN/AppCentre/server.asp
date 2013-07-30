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
<!-- #include file="server_include.asp"-->
<%
Dim intHighlight,objXmlHttp,strURL,bolPost,str,bolIsBinary,strList,bolFrame,strWrite,aryResponse
Dim strResponse

Call System_Initialize()
Call CheckReference("")
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AppCentre")=False Then Call ShowError(48)
Call LoadPluginXmlInfo("AppCentre")
Call AppCentre_InitConfig
BlogTitle="应用中心"

intHighlight=0


Call Server_Initialize()
Call Server_SendRequest()
Call Server_FormatResponse()
aryResponse=Array(Split(Split(strResponse,"</head>")(0),"<head>"),Split(Split(strResponse,"</body>")(0),"<body>"))

%>
<%If bolFrame Then%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<%=aryResponse(0)(Ubound(aryResponse(0)))%>

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader">应用中心</div>
          <div class="SubMenu">
            <%AppCentre_SubMenu(intHighlight)%>
          </div>
          <div id="divMain2">
            <%=strWrite%>
            
<%End If%>
<%=aryResponse(1)(Ubound(aryResponse(1)))%>
<%If bolFrame Then%>

          </div>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
        <%If login_pw<>"" Then%>
		<script type='text/javascript'>$('div.footer_nav p').html('&nbsp;&nbsp;&nbsp;<b>"&login_un&"</b>您好,欢迎来到APP应用中心!<a href=\'setting.asp?act=logout\'>[退出登录]</a>').css('visibility','inherit');</script>
		<%End If%>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
        
<%End If%>