<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize
'检查非法链接
Call CheckReference("")

If CheckPluginState("RegPage")=False Then Call ShowError(48)

Dim dUsername,dPassword,dEmail,dSite
	
dUsername=Replace(TransferHTML(Request.QueryString("dName"),"[nohtml]"),"""","&quot;")
dPassword=Replace(TransferHTML(Request.QueryString("dPassword"),"[nohtml]"),"""","&quot;")
dEmail=Replace(TransferHTML(Request.QueryString("dEmail"),"[nohtml]"),"""","&quot;")
dSite=Replace(TransferHTML(Request.QueryString("dSite"),"[nohtml]"),"""","&quot;")

For Each sAction_Plugin_RegPage_Begin in Action_Plugin_RegPage_Begin
	If Not IsEmpty(sAction_Plugin_RegPage_Begin) Then Call Execute(sAction_Plugin_RegPage_Begin)
Next
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-CN" />
<title>Z-Blog 注册页面</title>
<link rel="stylesheet" rev="stylesheet" href="<%=GetCurrentHost%>ZB_SYSTEM/css/admin.css" type="text/css" media="screen" />
<link rel="stylesheet" rev="stylesheet" href="style.css" type="text/css" media="screen" />
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/SCRIPT/common.js" type="text/javascript"></script>
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/SCRIPT/md5.js" type="text/javascript"></script>
</head>
<body class="short">
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="<%=GetCurrentHost%>ZB_SYSTEM/image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login">
    <div class="divHeader">注册用户</div>
	<p class="validateTips">以下所有项均需填写。</p>
    <!-- 原form -->
	<form action="reg_save.asp" method="post" id="reg">
      <%=Response_Plugin_RegPage_Begin%>
    <dl>
		<dd><label for="name">名称 </label><input type="text" id="name" name="name" size="30" tabindex="1" value="<%=dUsername%>"/></dd>

		<dd><label for="password">密码 </label><input id="password" name="password" type="password" maxlength="14" size="30" tabindex="2" value="<%=dPassword%>"></dd>

		<dd><label for="email">邮箱 </label><input id="email" name="email" maxlength="32" size="30"  tabindex="4" value="<%=dEMail%>"></dd>

		<dd><label for="site">网站 </label><input id="site" name="site" size="30" tabindex="5"  value="<%=dSite%>" /></dd>

		<dd><label for="edtCheckOut">验证 </label><input  type="text" id="edtCheckOut" name="edtCheckOut" size="30"  tabindex="6"/><img style="border:5px solid #ededed" src="<%=GetCurrentHost%>zb_system/function/c_validcode.asp?name=commentvalid" alt="点击刷新" title=""/></dd>

    <dd class="checkbox" >
      <!-- <dd class="checkbox"><input type="checkbox" checked="checked" name="chkRemember" id="chkRemember"  tabindex="3" /><label for="chkRemember"><%=ZC_MSG114%></label></dd> -->      
	<input type="checkbox" checked="checked" name="chkRemember" id="chkRemember"  tabindex="7" /><label for="chkRemember">阅读并同意本站的《<a target="_blank" href="agreement.txt">使用条款</a>》。</label>
    </dd>
	<dd class="submit"><input id="regButton" class="button" type="button" value="注册"  onClick="chk_reg()" tabindex="8" /></dd>
    </dl>
	<%=Response_Plugin_RegPage_End%>
	</form>
  </div>
</div>
<script language="javascript">
$(document).ready(function(){ 

		var objImageValid=$("img[src^='<%=GetCurrentHost%>zb_system/function/c_validcode.asp?name=commentvalid']");
		if(objImageValid.size()>0){
			objImageValid.css("cursor","pointer");
			objImageValid.click( function() {
					objImageValid.attr("src","<%=GetCurrentHost%>zb_system/function/c_validcode.asp?name=commentvalid"+"&amp;random="+Math.random());
			} );
		};
});

var vname = $("#name"),
	vemail = $("#email"),
	vpassword = $("#password"),
	vsite=$("#site"),
	allFields = $([]).add( vname ).add( vemail ).add( vpassword ).add( vsite ),
	tips = $(".validateTips");

function updateTips( t ) {
	tips
		.text( t )
		.addClass( "state-highlight" );
	setTimeout(function() {
		tips.removeClass( "state-highlight", 1500 );
	}, 1500 );
}

function checkLength( o, n, min, max ) {
	if ( o.val().length > max || o.val().length < min ) {
		o.addClass( "state-error" );
		updateTips(  n + "长度必须介于" + min + " ~ " + max + "位之间" );
		return false;
	} else {
		return true;
	}
}

function checkRegexp( o, regexp, n ) {
	if ( !( regexp.test( o.val() ) ) ) {
		o.addClass( "state-error" );
		updateTips( n );
		return false;
	} else {
		return true;
	}
}

function chk_reg(){
	var bValid = true;
	allFields.removeClass( "state-error" );

	bValid = bValid && checkLength( vname, "用户名", 3, 14 );
	bValid = bValid && checkLength( vpassword, "密码", 8, 16 );
	bValid = bValid && checkLength( vemail, "邮箱", 6, 60 );
	bValid = bValid && checkLength( vsite, "网址", 6, 60 );

	bValid = bValid && checkRegexp( vname, /^[.A-Za-z0-9\u4e00-\u9fa5]+$/i, "用户名只能使用中英文字符及数字。" );
	// From jquery.validate.js (by joern), contributed by Scott Gonzalez: http://projects.scottsplayground.com/email_address_validation/
	bValid = bValid && checkRegexp( vpassword, /^([0-9a-zA-Z])+$/, "密码只能使用英文字母及数字" );
	bValid = bValid && checkRegexp( vemail, /^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$/i, "邮箱格式有误，参考：abc@123.com" );
	bValid = bValid && checkRegexp( vsite, /^[a-zA-Z]+:\/\/[a-zA-Z0-9\\_\\-\\.\\&\\?\/:=#\u4e00-\u9fa5]+?\/*$/ig, "网站地址有误，参考：http://www.site.com" );

	if ( bValid ) {
		$("#reg").submit();
	}
	return true;
}

</script>
<%
For Each sAction_Plugin_RegPage_End in Action_Plugin_RegPage_End
	If Not IsEmpty(sAction_Plugin_RegPage_End) Then Call Execute(sAction_Plugin_RegPage_End)
Next
%>
</body>
</html>