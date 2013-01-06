<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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

Call AppCentre_InitConfig

If Request.QueryString("last")="now" Then
	Response.Clear
	Response.Write AppCentre_CheckSystemLast
	Response.End
End If


If Request.QueryString("check")="now" Then
	Call AppCentre_CheckSystemIndex(BlogVersion)
End If

Dim PathAndCrc32
Set PathAndCrc32=New TMeta

	Dim objXmlFile,strXmlFile,item
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")

	
	If fso.FileExists(BlogPath & "zb_users/cache/"&BlogVersion&".xml") Then

		strXmlFile =BlogPath & "zb_users/cache/"&BlogVersion&".xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else

				for each item in objXmlFile.documentElement.SelectNodes("file")
					PathAndCrc32.SetValue item.getAttributeNode("name").Value,item.getAttributeNode("crc32").Value
				next

			End If
		End If
	End If

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu">
            <% AppCentre_SubMenu(7)%>
          </div>
          <div id="divMain2">
            <form method="post" action="">
              <table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
                <tr>
                  <th width='50%'>当前版本</th>
                  <th>最新版本</th>
                </tr>
                <tr>
                  <td align='center'>Z-Blog <%=ZC_BLOG_VERSION%></td>
                  <td align='center' id="last">请稍等</td>
                </tr>
              </table>
              <br/>
              <p>
                <input type="button" onClick="location='update.asp?check=now'" value="检查当前版本的系统文件" />
              </p>
              <table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
                <tr>
                  <th width='78%'>文件名</th>
                  <th>状态</th>
                </tr>
				<tr>
                <td>错误计数</td><td><span id="errcount"></span></td>
                </tr>
                <%

Dim a,b,c,d,e,errCount
errCount=0
b=0
For Each a In PathAndCrc32.Names

	
	If b>0 Then
	
		c=vbsunescape(a)
		d=PathAndCrc32.GetValue(c)
		
		If Request.QueryString("check")="now" Then
		
			e=CRC32(BlogPath & c)
			If e=d Then
				e="<span class=""ok"">OK</span>"
			Else
				e="<span class=""fail"">Fail</span>"
				errCount=errCount+1
			End If
		End If
		
		Response.Write "<tr><td>"& c &"</td><td>"& e &"</td></tr>"
		Response.Flush
	
	End If
	b=b+1
Next


%>
              </table>
              <p> </p>
            </form>
          </div>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script> 
        <script type="text/javascript">
   
   $(document).ready(function(){
   
	$.get("update.asp?last=now", function(data){
	  $("#last").html("Z-Blog "+data);
	  var s=$("#errcount");
	  if(s){s.html("<%=errCount%>")}
	});

   });
   
   </script>
        <%
	If login_pw<>"" Then
		Response.Write "<script type='text/javascript'>$('div.SubMenu a[href=\'login.asp\']').hide();$('div.footer_nav p').html('&nbsp;&nbsp;&nbsp;<b>"&login_un&"</b>您好,欢迎来到APP应用中心!').css('visibility','inherit');</script>"
	End If
%>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->