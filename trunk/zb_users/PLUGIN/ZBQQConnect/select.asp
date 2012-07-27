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
<!-- #include file="../../../zb_users/c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
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
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="../../../zb_system/image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login">
<a href="bind.asp?QQOPENID=<%=Server.URLEncode(Request.QueryString("QQOpenID"))%>" target="_blank" title="绑定现有帐号"><img src="data:image/png;base64,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" /></a>
<a href="../Reg/Reg.asp?QQOPENID=<%=Server.URLEncode(Request.QueryString("QQOpenID"))%>" target="_blank" title="新建账户"><img src="data:image/png;base64,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"/></a>
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

	document.getElementById("frmLogin").action="cmd.asp?act=verify"
	document.getElementById("username").value=unescape(strUserName);
	document.getElementById("password").value=strPassWord
	document.getElementById("savedate").value=strSaveDate
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