<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    haphic
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("MusicLink")=False Then Call ShowError(48)

Call CmtN_Initialize
Dim strAct
strAct=Request.QueryString("act")

BlogTitle="博客音乐外链设置"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style>
#totorobox input[type="text"]{width:90%}#totorobox input[type="password"]{width:90%}
.content-box ol li{height: auto;clear:both;float: none;list-style-type: decimal;}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<%If (Not strAct="TestMail") And (Not Request.Cookies("CmtN_licence")="disabled") Then%>
<div id="licence" style="position:absolute;top:26px;left:10px;right:10px;padding:10px 20px;border:3px double #99BBFF;background-color:#EEEEFF;cursor:default;">
<p style=";border-bottom:1px dashed #99BBFF;font-size:15px;text-align:center;">关于发信 SMTP 帐户设置的潜在风险, 规避办法, 及插件作者免责条款.</p>
<p>请仔细阅读以下内容:</p>
<p>尽管插件本身不会向网站访问者暴露插件的设置, 特别是邮箱帐号和密码, 但仍旧存在以下风险.</p>
<ol>
<li>因为发信需要向 SMTP 服务器提供密码, 所以密码必须保存在博客主机上且无法进行有效加密(不可逆加密).</li>
<li>您的服务器管理者可以看到主机上任何源文件并进行修改, 其中包括储存 SMTP 帐户密码的文件.</li>
<li>有博客管理员身份的人可以看到插件设置及空间内任何文件的源代码并进行修改.</li>
</ol>
<p>您可以采取以下办法尽量规避风险.</p>
<ol>
<li>拥有管理员帐户的人必须是您信任的人.</li>
<li>尽量不要使用自己的常用邮箱发信. 应该<a href="http://www.esloy.com/blog/archives/2008/11/CmtN1.2-Released.html#add1" target="_blank">注册一个新邮箱</a>或在自己的邮局里开一个新帐号作为发信专用.</li>
<li>如果是QQ邮箱, 最好使用英文别名和邮箱独立密码.</li>
</ol>
<p>对以下情形, 插件作者没有承担责任的义务.</p>
<ol>
<li>邮箱密码被其它人通过查看服务器文件的方式盗取.</li>
<li>因大量发信而使 SMTP 帐户被服务提供方限制或锁定.</li>
<li>没有妥善设置和保管第三方代发 Key, 被他人利用此功能发信, 而使您的 SMTP 帐户受牵连.</li>
<li>和本插件基本功能无关的由于邮箱使用所带来的风险.</li>
</ol>
<p>当您关闭此页面并进行设置时, 则认为您已了解此插件所带来的风险, 做好应对准备, 并认可插件作者的免责条款. 否则, 请停用并删除此插件.</p>
<div onClick="document.getElementById('licence').style.display = 'none';var date=new Date();var expireDays=30;date.setTime(date.getTime()+expireDays*24*3600*1000);document.cookie = 'CmtN_licence=disabled;expires='+date.toGMTString();" style="width:280px;margin:5px auto;padding:3px 5px;border:3px double #99BBFF;background-color:#FFFFFF;cursor:pointer;text-align:center;">我了解并同意以上内容! (点此关闭页面)</div>
</div>
<%End If%>

			<div id="divMain">
<div class="divHeader">新评论邮件提醒</div>
<div class="SubMenu">
	<a href="setting.asp"><span class="m-left m-now">[插件设置]</span></a>
</div>
<div id="divMain2">
<form id="edit" name="edit" method="post">
<%
If strAct="TestMail" Then

	Response.Write "<div id=""SendingMail""><pre style=""color:blue;font-size:18px;font-weight:bold;font-family:黑体;text-align:center;""><p>尝试发送测试邮件...</p></pre></div>"
	Response.Flush

	Dim MailStatus,MailDesc
	MailStatus = CmtN_SendMessage(CmtN_MailToAddress,"null",CmtN_MailReplyToAddress,CmtN_MailFromName,"新评论提醒插件 - 试发邮件","<p>恭喜您!</p><p>当您收到此邮件时, 您的插件已配置正确, 您可以正常进行提醒邮件的发送.</p>")
	If Err.Number<>0 Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('SendingMail').style.display = 'none';}catch(e){};</script>"
	
		If MailStatus = True Then
			MailDesc = "<font color=""green""> √ 邮件可能发送成功! </font>"
		Else
			MailDesc = "<font color=""red""> × 邮件发送失败! </font>"
		End If
	
		Response.Write "<div onclick=""if(document.getElementById('MailLog').style.display == 'none'){document.getElementById('MailLog').style.display = 'block';}else{document.getElementById('MailLog').style.display = 'none';}"" style=""border:1px solid #99BBFF;background-color:#EEEEFF;cursor:pointer;text-align:center;""><p>"& MailDesc &"点此查看详情. 如果未能收到测试邮件, 请调整设置后重新提交.</p>"
	Else
		Response.Write Err.Number & "<br/><br/>" & Err.Description
	End If

	Response.Write "<div id=""MailLog"" style=""display:none;""><p>" & Replace(Application(ZC_BLOG_CLSID & "CmtN_LastMailLog"),vbcrlf,"<br />") & "</p></div></div>"

Else

	Dim aryFL,sFL,nFL
	aryFL=LoadIncludeFiles("zb_users/PLUGIN/CmtN/OutGoingMails/")
	nFL=0

	For Each sFL In aryFL
		If (sFL<>"") And (Not InStr(LCase(sFL),"index.html")>0) Then nFL=nFL+1
	Next

	Response.Write "<div onclick=""if(document.getElementById('MailLog').style.display == 'none'){document.getElementById('MailLog').style.display = 'block';}else{document.getElementById('MailLog').style.display = 'none';}"" style=""border:1px solid #99BBFF;background-color:#EEEEFF;cursor:pointer;text-align:center;""><p>查看上次发送日志, 剩余待发邮件数("& nFL &")</p>"

	Response.Write "<div id=""MailLog"" style=""display:none;""><p>" & Replace(Application(ZC_BLOG_CLSID & "CmtN_LastMailLog"),vbcrlf,"<br />") & "</p></div></div>"

End If




If checkServerObject("Jmail.Message") Then
	Response.Write "<p><b><font color=""Green"">注意:</font> 您的主机支持Jmail4, 可自主设定发送邮件服务器. <a href=""http://www.zsxsoft.com/archives/255.html"" target=""_blank"">如何选择和设定发件邮箱?</a></b></p>"
Else
	If checkServerObject("cdo.Message") Then
		Response.Write "<p><b><font color=""Green"">注意:</font> 您的主机不支持Jmail4, 插件将使用CDO组件. <a href=""http://www.zsxsoft.com/archives/255.html"" target=""_blank"">如何选择和设定发件邮箱?</a></b></p>"
	Else
		Response.Write "<p><b><font color=""red"">注意:</font> 您的主机啥都不支持。。</b></p>"

	End If
End If
%>
              <div class="content-box">
              <!-- Start Content Box -->
              
              <div class="content-box-header">
                <ul class="content-box-tabs">
                  <li><a href="#tab1" class="default-tab"><span>邮箱地址设置</span></a></li>
                  <li><a href="#tab2"><span>延时发送设置</span></a></li>
                  <li><a href="#tab3"><span>邮件服务器设置</span></a></li>
                  <li><a href="#tab4"><span>帮助</span></a></li>
                </ul>
                <div class="clear"></div>
              </div>
              <!-- End .content-box-header -->
              
              <div class="content-box-content" id="totorobox">
              <div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab1">
              <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
              <tr height="40">
                <td width='40'>序号</td>
                <td width='130' align="center"></td>
                <td width="500" align="center"></td>
                <td align="center">说明</td>
              </tr>
              <tr><td>1</td><td>收件人地址(<font color='red'>*</font>)</td><td><input name="strCmtN_MailToAddress" type="text" value="<%=CmtN_MailToAddress%>"/></td><td>例如: haphic@gmail.com,loybal@gmail.com)(多邮箱用逗号 “,” 隔开, 如果站长不想接收提醒可填写 "null"</td></tr>
              <tr><td>2</td><td>邮件回复地址(<font color="blue">*</font>)</td><td><input name="strCmtN_MailReplyToAddress" type="text" value="<%=CmtN_MailReplyToAddress%>"/></td><td>(only JMail)如果收件人回复提醒邮件,将发到此邮箱</td></tr>
              <tr><td>3</td><td>发件人姓名(<font color="gray">*</font>)</td><td><input name="strCmtN_MailFromName" type="text" value="<%=CmtN_MailFromName%>"/></td><td>(only JMail)选填, 如果需向留有邮箱的评论者发送提醒邮件, 建议设置</td></tr>
              <tr><td>4</td><td>同时提醒评论者(<font color="gray">*</font>)</td><td><input name="strCmtN_NotifyCmtLeaver" type="text" class="checkbox" value="<%=CStr(CmtN_NotifyCmtLeaver)%>"/></td><td>如果评论者留下邮箱, 则在评论被回复时向该评论者发送提醒邮件</td></tr>
              </table>
              </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
              <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
              <tr height="40">
                <td width='40'>序号</td>
                <td width='130' align="center"></td>
                <td width="500" align="center"></td>
                <td align="center">说明</td>
              </tr>
              <tr><td>1</td><td>启用邮件延时发送</td><td><input name="strCmtN_MailSendDelay" type="text" class="checkbox" value="<%=CStr(CmtN_MailSendDelay)%>" id="MailSendDelay"/></td><td>如果评论者留下邮箱, 则在评论被回复时向该评论者发送提醒邮件</td></tr>
              <tr><td>2</td><td>两次发送间隔时间</td><td><input name="strCmtN_MailSendDelayTime" type="text" value="<%=CmtN_MailSendDelayTime%>" <%=IIf(CmtN_MailSendDelay,"","disabled")%>  id="MailSendDelayTime"/></td><td> (单位:秒。推荐 60-120, 最大:7200。启用全静态化时无效。)</td></tr>
              </table>
              </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
              <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
              <tr height="40">
                <td width='40'>序号</td>
                <td width='130' align="center"></td>
                <td width="500" align="center"></td>
                <td align="center">说明</td>
              </tr>
              <tr><td>1</td><td>发送邮件服务器</td><td><input name="strCmtN_MailServerName" type="text" value="<%=CmtN_MailServerName%>"/></td><td>例如: smtp.163.com</td></tr>
              <tr><td>2</td><td>登陆帐号</td><td><input name="strCmtN_MailServerUserName" type="text" value="<%=CmtN_MailServerUserName%>" /></td><td> 例如: haphic@163.com 抑或只填 haphic, 有的邮箱需要身份验证</td></tr>
              <tr><td>3</td><td>登陆密码(<font color="blue">*</font>)</td><td><input name="strCmtN_MailServerUserPwd" type="password" value="<%=CmtN_MailServerUserPwd%>" /></td><td>登陆帐号的密码, 有的邮箱需要身份验证</td></tr>
              <tr><td>4</td><td>发件人地址</td><td><input name="strCmtN_MailFromAddress" type="text" value="<%=CmtN_MailFromAddress%>" /></td><td> 一般需与登陆帐号同名, 否则会导致发送失败</td></tr>
              <tr><td>5</td><td>备用发信服务器(<font color="Blue">*</font>)</td><td><input name="strCmtN_MailServerAlternate" type="text" value="<%=CmtN_MailServerAlternate%>" /></td><td>(only JMail)写法: 用户名:密码@服务器地址(发信地址)</td></tr>

              </table>
              </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab4">
                <ol>
                <li>本插件需要运行环境：安装有JMail4或者安装有CDO.Message的服务器.jMail组件支持的功能更多，插件内将优先使用。</li>
                <li>
                您首先要设置好收信地址, 用作接收提醒邮件.
                </li>
                <li>
                如果选中"同时提醒评论者"选项, 建议指定邮件回复地址和收件人姓名, 以方便评论者在收到提醒邮件后与您联系.最好将邮件回复地址和收件人地址指定为同一邮箱.
                </li>
                <li>
                同时向留下留下邮箱的评论者发送提醒邮件的功能被开启后, 如果评论者留下虚假邮箱等, 可能会对您的SMTP邮件帐号造成影响, 也可能不会, 我不清楚. 如果您不在乎这个功能就不要开启了(但这个功能默认是开启的).
                </li>
                <li>
                可以用此插件实现只提醒评论者而不提醒博主, 只要将收件人地址设为 null 就行. 但前提是您要确定可以发出提醒邮件.
                </li>
                <li>
                您需要确认您的主机是否支持 Jmail4 组件. (在插件设置页中即有提示.)
                </li>
                <li>
                您需要有一个支持SMTP(25端口, 非SSL连接)的邮箱用作邮件发送. (<a href="http://www.esloy.com/blog/archives/2008/11/CmtN1.2-Released.html#add1" target="_blank">查看详情</a>)
                </li>
                <li>
                设置好此邮箱的SMTP服务器, 登陆帐号, 密码, 邮箱地址等信息.
                </li>
                <li>
                您还可以参考以上的资料设置一个备用发信服务器. 当发信失败后, 系统会自动尝试使用备用服务器发信. 如不想设置设留空.
                </li>
                <li>
                延时发送可加速评论提交, 提醒邮件将在动态主页和分类页被访问时发送, 请根据你页面访问量设置合理的发送间隔.
                </li>
                <li>
                延时发送不适用于以下情况: 1,全静态化. 2,博客流量很低. 3,博客评论量很大 (实际上这种情况这个插件都不适用了).
                </li>
                <li>
                插件为发信内容提供了HTML模板, 即插件目录下的三个 .html 文件. 可根据需要在模板中修改邮件内容. (<a href="http://www.esloy.com/blog/archives/2008/11/CmtN1.2-Released.html#add3" target="_blank">查看详情</a>)
                </li>
                </ol>
                
                <hr />
                <p>
                以上只是插件的基本介绍, 由于插件说明复杂且需要时时变动, <a href="http://www.esloy.com/blog/archives/2008/11/CmtN1.2-Released.html" target="_blank">请点此查看更详细的说明.</a>
                </p>
                
                <p>
                如有其它相关问题可 <a href="http://www.esloy.com/blog/archives/2008/11/CmtN1.2-Released.html#comment" target="_blank">到此页提出</a> 或 <a href="mailto:haphic@gmail.com">发我邮件</a>.
                </p>

              </div>
              </div>
              

<%


Response.Write "<br/><p><input type=""submit"" class=""button"" value=""保存插件设置"" id=""btnPost"" onclick='document.getElementById(""edit"").action=""savesetting.asp"";' /> <input name=""TestMail"" id=""sendTestMail"" type=""checkbox"" value=""True""/><label for=""sendTestMail""> 保存设置后发送一封测试邮件!</label></p>"

Response.Write "<br/>"

%>

</form>

</div>
</div>
<script language="JavaScript" type="text/javascript">
$(document).ready(function() {
		$("#MailSendDelay").next().bind("click",function(){
		var checkMailBrige=$('#MailSendDelay');
	
		if(checkMailBrige.val()=="True"){
			$('#MailSendDelayTime').removeAttr("disabled")
		}
		else{
			$('#MailSendDelayTime').attr("disabled","disabled");
		}
	}
	)
});



</script>
<script language="JavaScript" runat="server">
// Check if the object is usable on the server
function checkServerObject(strObjectName){
  try{
    var obj=Server.CreateObject(strObjectName);
  }catch(e){
    return false;
  }
  delete obj;

  return true;
}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>

