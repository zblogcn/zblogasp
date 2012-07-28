<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    login.asp
'// 开始时间:    2004.07.27
'// 最后修改:    
'// 备    注:    登陆页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->

<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->

<%
Call CheckReference("")
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> 
	<link rel="stylesheet" rev="stylesheet" href="../../../zb_system/css/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../../../zb_system/script/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="../../../zb_system/script/md5.js" type="text/javascript"></script>
	<title><%=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG009%></title>
</head>
<body>
<%
If Request.QueryString("act")="verify" Then
	Call System_Initialize
	
	If Login=True Then
		Call ZBQQConnect_RegSave(BlogUser.ID)
		Response.Write "<script>alert('绑定成功！');location.href="""&GETCurrentHost&"/ZB_SYSTEM/ADMIN/ADMIN.ASP?ACT=SiteInfo""</script>"
		Response.End
	End iF
End If
%>
%>
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="../../../zb_system/image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login">
    <form id="frmLogin" method="post" action="">
    <dl>
      <dd><label for="edtUserName"><%=ZC_MSG003%>:</label><input type="text" id="edtUserName" name="edtUserName" size="20" tabindex="1" /></dd>
      <dd><label for="edtPassWord"><%=ZC_MSG002%>:</label><input type="password" id="edtPassWord" name="edtPassWord" size="20" tabindex="2" /></dd>
      <input type="hidden" name="QQOPENID" value="<%=TransferHTML(Request.QueryString("QQOPENID"),"[nohtml]")%>"/>
    </dl>
    <dl>
      <dd class="submit"><input id="btnPost" name="btnPost" type="submit" value="<%=ZC_MSG260%>" class="button" tabindex="4"/></dd>
    </dl>
<input type="hidden" name="username" id="username" value="" />
<input type="hidden" name="password" id="password" value="" />
<input type="hidden" name="savedate" id="savedate" value="30" />
    </form>
  </div>
</div>


<script language="JavaScript" type="text/javascript">

function SetCookie(sName, sValue,iExpireDays) {
	if (iExpireDays){
		var dExpire = new Date();
		dExpire.setTime(dExpire.getTime()+parseInt(iExpireDays*24*60*60*1000));
		document.cookie = sName + "=" + escape(sValue) + "; expires=" + dExpire.toGMTString()+ "; path=/";
	}
	else{
		document.cookie = sName + "=" + escape(sValue)+ "; path=/";
	}
}

if(GetCookie("username")){document.getElementById("edtUserName").value=unescape(GetCookie("username"))};

$("#btnPost").click(function(){

	var strUserName=document.getElementById("edtUserName").value;
	var strPassWord=document.getElementById("edtPassWord").value;
	var strSaveDate=document.getElementById("savedate").value

	if((strUserName=="")||(strPassWord=="")){
		alert("<%=ZC_MSG010%>");
		return false;
	}

	strUserName=escape(strUserName);

	strPassWord=MD5(strPassWord);

	SetCookie("username",strUserName,strSaveDate);
	SetCookie("password",strPassWord,strSaveDate);

	document.getElementById("frmLogin").action="bind.asp?act=verify"
	document.getElementById("username").value=unescape(strUserName);
	document.getElementById("password").value=strPassWord;
	document.getElementById("savedate").value=strSaveDate;
	document.getElementById("QQOPENID").value="<%=TransferHTML(Request.QueryString("QQOPENID"),"nohtml")%>";
})

$(document).ready(function(){ 
	if($.browser.msie){
		$(":checkbox").css("margin-top","4px");
	}
});

$("#chkRemember").click(function(){
	$("#savedate").attr("value",$("#chkRemember").attr("checked")==true?30:0);
})

</script>
</body>
</html>
<%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>