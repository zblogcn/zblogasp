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

If Request.QueryString("update")="now" Then
	'Response.Clear
	'Response.Write AppCentre_CheckSystemLast
	'Response.End
End If

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


If CLng(Request.QueryString("crc32"))>0 Then

	Response.Clear
	If CLng(Request.QueryString("crc32"))<=Round(PathAndCrc32.Count/10) Then

		Dim i,j,k,l,m,n
		k=CLng(Request.QueryString("crc32"))
		i=(k-1)*10+1
		j=k*10
		m="<img src=\'"&BlogHost&"zb_system/image/admin/ok.png\'>"
		n="<img src=\'"&BlogHost&"zb_system/image/admin/exclamation.png\'>"
		For l=i To j
			If l>PathAndCrc32.Count Then Exit For
			If CRC32(BlogPath & vbsunescape(PathAndCrc32.Names(l)))<>PathAndCrc32.Values(l) Then
				Response.Write "$('#td"&l&"').html('"&n&"');"
			Else
				Response.Write "$('#td"&l&"').html('"&m&"');"
			End If
		Next
		
	End If
	Response.End

End If


BlogTitle="系统更新检查"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<% AppCentre_SubMenu(7)%>
  </div>
  <div id="divMain2">
<form method="post" action="">
<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th width='50%'>当前版本</th><th>最新版本</th></tr>

<tr><td align='center'>Z-Blog <%=ZC_BLOG_VERSION%></td>
<td align='center' id="last"></td></tr>


</table>

<p><input type="button" onClick="location='update.asp?update=now'" value="升级新版程序" /></p>
<div class="a"></div>
<div class="divHeader">校验系统核心文件&nbsp;&nbsp;<a href="update.asp?check=now"><img src="Images/refresh.png" width="16" alt="校验" /></a><span id="bar"></sp></div>


<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th width='78%'>文件名</th><th>状态</th></tr>
<!--

 -->
<%

Dim a,b,c,d,e
b=0
For Each a In PathAndCrc32.Names

If b>0 Then

c=vbsunescape(a)

Response.Write "<tr><td><img src='Images/document_empty.png' width='16' alt='' /> "& c &"</td><td id='td"&b&"' align='center'>"& e &"</td></tr>"
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

   function crc32(i){
    $("#bar").prev().hide();
	$.get("update.asp?crc32="+i, function(data){
	  if(data!==""){i=i+1;$("#bar").html($("#bar").html()+"█");eval(data);crc32(i);}else{$("#bar").hide();$("#bar").prev().show();}
	});
   }
   
   $(document).ready(function(){
   
$.get("update.asp?last=now", function(data){
  $("#last").html("Z-Blog "+data);
});

   });


<%
If Request.QueryString("check")="now" Then
	Response.Write "crc32(1)"
End If
%>
   
   </script>
<%
	If login_pw<>"" Then
		Response.Write "<script type='text/javascript'>$('div.SubMenu a[href=\'login.asp\']').hide();$('div.footer_nav p').html('&nbsp;&nbsp;&nbsp;<b>"&login_un&"</b>您好,欢迎来到APP应用中心!').css('visibility','inherit');</script>"
	End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->