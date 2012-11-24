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
Dim objArticle
Set objArticle = New TArticle
objArticle.FType=ZC_POST_TYPE_PAGE
If GetTemplate("TEMPLATE_REGPAGE")<>empty Then
    objArticle.template = "REGPAGE"
End If
objArticle.Title = "注册"
objArticle.Content="" & vbCrlf & _
"	<p class=""validateTips"">以下用户名密码为必填项。</p>" & vbCrlf & _
"    <!-- 原form -->" & vbCrlf & _
"	<form action=""reg_save.asp"" method=""post"" id=""reg"">" & vbCrlf & _
      Response_Plugin_RegPage_Begin & vbCrlf & _
"    <dl>" & vbCrlf & _
"		<dd><label for=""name"">名称 </label><input type=""text"" id=""name"" name=""name"" size=""30"" tabindex=""1"" value=""" & dUsername & """/>  <span style=""color:red"">(*)</span></dd>" & vbCrlf & _

vbCrlf & _
"		<dd><label for=""password"">密码 </label><input id=""password"" name=""password"" type=""password"" maxlength=""14"" size=""30"" tabindex=""2"" value=""" & dPassword & """>  <span style=""color:red"">(*)</span></dd>" & vbCrlf & _
"		<dd><label for=""repassword"">确认 </label><input id=""repassword"" name=""repassword"" type=""password"" maxlength=""14"" size=""30"" tabindex=""3"" value=""""></dd>" & vbCrlf & _

vbCrlf & _
"		<dd><label for=""email"">邮箱 </label><input type=""email"" id=""email"" name=""email"" maxlength=""32"" size=""30""  tabindex=""4"" value=""" & dEMail & """></dd>" & vbCrlf & _

vbCrlf & _
"		<dd><label for=""site"">网站 </label><input type=""url"" id=""site"" name=""site"" size=""30"" tabindex=""5""  value=""" & dSite & """ /></dd>" & vbCrlf & _

vbCrlf & _
"		<dd><label for=""edtCheckOut"">验证 </label><input type=""number"" min=""10000"" max=""99999"" id=""edtCheckOut"" name=""edtCheckOut"" size=""30""  tabindex=""6""/><img style=""border:5px solid #ededed"" src=""" & GetCurrentHost & "zb_system/function/c_validcode.asp?name=commentvalid"" alt=""验证码"" title=""点击刷新""/></dd>" & vbCrlf & _

vbCrlf & _
"    <dd class=""checkbox"" >" & vbCrlf & _
"      <!-- <dd class=""checkbox""><input type=""checkbox"" checked=""checked"" name=""chkRemember"" id=""chkRemember""  tabindex=""3"" /><label for=""chkRemember"">" & ZC_MSG114 & "</label></dd> -->      " & vbCrlf & _
"	<input type=""checkbox"" checked=""checked"" name=""chkRemember"" id=""chkRemember""  tabindex=""7"" /><label for=""chkRemember"">阅读并同意本站的《<a target=""_blank"" href=""agreement.txt"">使用条款</a>》。</label>" & vbCrlf & _
"    </dd>" & vbCrlf & _
"	<dd class=""submit""><input id=""regButton"" class=""button"" type=""button"" value=""注册""  onClick=""chk_reg()"" tabindex=""8"" /></dd>" & vbCrlf & _
"    </dl>" & vbCrlf & _
    Response_Plugin_RegPage_End & vbCrlf & _
"	</form>" & vbCrlf & _

"<script language=""javascript"">" & vbCrlf & _
"$(document).ready(function(){ " & vbCrlf & _

vbCrlf & _
"		var objImageValid=$(""img[src^='" & GetCurrentHost & "zb_system/function/c_validcode.asp?name=commentvalid']"");" & vbCrlf & _
"		if(objImageValid.size()>0){" & vbCrlf & _
"			objImageValid.css(""cursor"",""pointer"");" & vbCrlf & _
"			objImageValid.click( function() {" & vbCrlf & _
"					objImageValid.attr(""src"",""" & GetCurrentHost & "zb_system/function/c_validcode.asp?name=commentvalid""+""&amp;random=""+Math.random());" & vbCrlf & _
"			} );" & vbCrlf & _
"		};" & vbCrlf & _
"});" & vbCrlf & _

vbCrlf & _
"var vname = $(""#name"")," & vbCrlf & _
"	vemail = $(""#email"")," & vbCrlf & _
"	vpassword = $(""#password"")," & vbCrlf & _
"	vrepassword = $(""#repassword"")," & vbCrlf & _
"	vsite=$(""#site"")," & vbCrlf & _
"	allFields = $([]).add( vname ).add( vemail ).add( vpassword ).add( vsite )," & vbCrlf & _
"	tips = $("".validateTips"");" & vbCrlf & _

vbCrlf & _
"function updateTips( t ) {" & vbCrlf & _
"	tips" & vbCrlf & _
"		.text( t )" & vbCrlf & _
"		.addClass( ""state-highlight"" );" & vbCrlf & _
"	setTimeout(function() {" & vbCrlf & _
"		tips.removeClass( ""state-highlight"", 1500 );" & vbCrlf & _
"	}, 1500 );" & vbCrlf & _
"}" & vbCrlf & _

vbCrlf & _
"function checkLength( o, n, min, max ) {" & vbCrlf & _
"	if ( o.val().length > max || o.val().length < min ) {" & vbCrlf & _
"		o.addClass( ""state-error"" );" & vbCrlf & _
"		updateTips(  n + ""长度必须介于"" + min + "" ~ "" + max + ""位之间"" );" & vbCrlf & _
"		return false;" & vbCrlf & _
"	} else {" & vbCrlf & _
"		return true;" & vbCrlf & _
"	}" & vbCrlf & _
"}" & vbCrlf & _

vbCrlf & _
"function checkRegexp( o, regexp, n ) {" & vbCrlf & _
"	if ( !( regexp.test( o.val() ) ) ) {" & vbCrlf & _
"		o.addClass( ""state-error"" );" & vbCrlf & _
"		updateTips( n );" & vbCrlf & _
"		return false;" & vbCrlf & _
"	} else {" & vbCrlf & _
"		return true;" & vbCrlf & _
"	}" & vbCrlf & _
"}" & vbCrlf & _

vbCrlf & _
"function checkPassword( o, n ) {" & vbCrlf & _
"	if ( o.val()!=n.val()) {" & vbCrlf & _
"		o.addClass( ""state-error"" );" & vbCrlf & _
"		n.addClass( ""state-error"" );" & vbCrlf & _
"		updateTips(  ""请重新确认密码是否正确"" );" & vbCrlf & _
"		return false;" & vbCrlf & _
"	} else {" & vbCrlf & _
"		return true;" & vbCrlf & _
"	}" & vbCrlf & _
"}" & vbCrlf & _

vbCrlf & _
"function chk_reg(){" & vbCrlf & _
"	var bValid = true;" & vbCrlf & _
"	allFields.removeClass( ""state-error"" );" & vbCrlf & _

vbCrlf & _
"	bValid = bValid && checkLength( vname, ""用户名"", " & ZC_USERNAME_MIN & ", " & ZC_USERNAME_MAX & " );" & vbCrlf & _
"	bValid = bValid && checkRegexp( vname, /^[.A-Za-z0-9\u4e00-\u9fa5]+$/i, ""用户名只能使用汉字、英文字母及数字。"" );" & vbCrlf & _

vbCrlf & _
"	bValid = bValid && checkLength( vpassword, ""密码"", " & ZC_PASSWORD_MIN & ", " & ZC_PASSWORD_MAX & " );" & vbCrlf & _
"	bValid = bValid && checkRegexp( vpassword, /^([A-Za-z0-9`~!@#\$%\^&\*\-_])+$/, ""密码只能使用英文字母、数字及部分特殊字符。"" );" & vbCrlf & _
"	bValid = bValid && checkPassword(vpassword,vrepassword);" & vbCrlf & _

vbCrlf & _
"	if (vemail.val().length > 0){" & vbCrlf & _
"		bValid = bValid && checkLength( vemail, ""邮箱"", 6, 60 );" & vbCrlf & _
"		bValid = bValid && checkRegexp( vemail, /^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$/i, ""邮箱格式有误，参考格式：abc@123.com"" );" & vbCrlf & _
"	}" & vbCrlf & _
"	if (vsite.val().length > 0){" & vbCrlf & _
"		bValid = bValid && checkLength( vsite, ""网址"", 6, 60 );" & vbCrlf & _
"		bValid = bValid && checkRegexp( vsite, /^[a-zA-Z]+:\/\/[a-zA-Z0-9\\_\\-\\.\\&\\?\/:=#\u4e00-\u9fa5]+?\/*$/ig, ""网址格式有误，参考格式：http://www.site.com"" );" & vbCrlf & _
"	}" & vbCrlf & _
"	if ( bValid ) {" & vbCrlf & _
"		$(""#reg"").submit();" & vbCrlf & _
"	}" & vbCrlf & _
"	return true;" & vbCrlf & _
"}" & vbCrlf & _

vbCrlf & _
"</script>"
If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
	objArticle.Build
	objArticle.Html=Replace(objArticle.Html,"</head>","<link rel=""stylesheet"" rev=""stylesheet"" href=""style.css"" type=""text/css"" media=""screen"" /></head>")
	Response.Write objArticle.Html
End If
For Each sAction_Plugin_RegPage_End in Action_Plugin_RegPage_End
	If Not IsEmpty(sAction_Plugin_RegPage_End) Then Call Execute(sAction_Plugin_RegPage_End)
Next
%>
</body>
</html>