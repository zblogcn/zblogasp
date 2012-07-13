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
<!-- #include file="../zb_users/c_option.asp" -->
<!-- #include file="../zb_system/function/c_function.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> 
	<link rel="stylesheet" rev="stylesheet" href="css/login.css" type="text/css" media="screen" />
	<script language="JavaScript" src="SCRIPT/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="SCRIPT/md5.js" type="text/javascript"></script>
	<title><%=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG009%></title>
</head>
<body>
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login">
    <form id="frmLogin" method="post" action="">
    <dl>
      <dd><label for="edtUserName"><%=ZC_MSG001%>:</label><input type="text" id="edtUserName" name="edtUserName" size="20" /></dd>
      <dd><label for="edtPassWord"><%=ZC_MSG002%>:</label><input type="password" id="edtPassWord" name="edtPassWord" size="20" /></dd>
    </dl>
    <dl>
      <dd class="checkbox"><input type="checkbox" checked="checked" name="chkRemember" id="chkRemember"></input><label for="chkRemember"><%=ZC_MSG004%></label></dd>
      <dd class="submit"><input id="btnPost" name="btnPost" type="submit" tabindex="6" value="<%=ZC_MSG260%>" onclick="JavaScript:return VerifyMessage()" class="button" /></dd>
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
	alert($("#savedate").attr("value"))
})

</script>
</body>
</html>
<%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>