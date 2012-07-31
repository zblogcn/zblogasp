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
'代码部分来源于网络，作者未知
Call ActivePlugin
'检查非法链接
Call CheckReference("")

If CheckPluginState("RegPage")=False Then Call ShowError(48)


Dim dUsername,dPassword,dEmail,dSite
	
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
    <div class="divHeader">Z-Blog 注册</div>
    <form name="zblogform" action="reg_save.asp" method="post" onSubmit="return chk_reg()" id="reg">
      <%=Response_Plugin_RegPage_Begin%>
      <ul>
        <li class="r_left">用户名<font color="red">(*)</font>：</li>
        <li class="r_right">
          <input style="width:150px;" id="uname" onBlur="out_uname();" onFocus="on_input('d_uname');" maxlength="30" size="15" name="username" value="<%=dUsername%>">
        </li>
        <li class="r_msg">
          <div class="d_default" id="d_uname"></div>
        </li>
      </ul>
      <ul>
      </ul>
      <ul>
        <li class="r_left">密码<font color="red">(*)</font>：</li>
        <li class="r_right">
          <input style="width:150px;" id="upwd" onBlur="out_upwd1();" onChange="EvalPwdStrength(this.value);" onFocus="EvalPwdStrength(this.value);on_input('d_upwd1');" type="password" maxlength="12" name="password" value="<%=dPassword%>">
        </li>
        <li class="r_msg">
          <div class="d_default" id="d_upwd1"></div>
        </li>
      </ul>
      <ul>
        <li class="r_left"></li>
        <li class="r_right">
          <div class="ob_pws" id="pws">
            <div class="ob_pws0" id="idSM1"><span style="FONT-SIZE: 1px">&nbsp;</span><span id="idSMT1">弱</span></div>
            <div class="ob_pws0" id="idSM2" style="BORDER-LEFT: #dedede 1px solid"><span style="FONT-SIZE: 1px">&nbsp;</span><span id="idSMT2">中</span></div>
            <div class="ob_pws0" id="idSM3" style="BORDER-LEFT: #dedede 1px solid"><span style="FONT-SIZE: 1px">&nbsp;</span><span id="idSMT3">强</span></div>
          </div>
        </li>
      </ul>
      <ul>
        <li class="r_left">确认密码<font color="red">(*)</font>：</li>
        <li class="r_right">
          <input style="width:150px;" id="repassword" onBlur="out_upwd2();" onFocus="on_input('d_upwd2');" type="password" maxlength="12" name="repassword" value="<%=dPassword%>">
        </li>
        <li class="r_msg">
          <div class="d_default" id="d_upwd2"></div>
        </li>
      </ul>
      <ul>
        <li class="r_left">电子邮箱<font color="red">(*)</font>：</li>
        <li class="r_right">
          <input style="width:150px;" id="email" onBlur="out_email();" onFocus="on_input('d_email');" maxlength="32" size="30" name="email" value="<%=dEMail%>">
        </li>
        <li class="r_msg">
          <div class="d_email" id="d_email"></div>
        </li>
      </ul>
      <ul>
        <li class="r_left">网站或博客地址<font color="red">(*)</font>：</li>
        <li class="r_right">
          <input style="width:150px;" id="site" name="site" size="30" onBlur="out_site();" onFocus="on_input('d_site');"  value="<%=dSite%>" />
        </li>
        <li class="r_msg">
          <div class="d_site" id="d_site"></div>
        </li>
        <li class="r_left"><span id="save_stat"></span></li>
      </ul>
      <%=Response_Plugin_RegPage_End%>
      <ul>
        <li class="r_left">验证<font color="red">(*)</font></li>
        <li class="r_right">
          <input style="width:150px;" type="text" id="edtCheckOut" name="edtCheckOut" size="10" />
        </li>
        <li class="r_msg"> &nbsp;&nbsp;<img style="border:1px solid black" src="<%=GetCurrentHost%>zb_system/function/c_validcode.asp?name=commentvalid" height="20" width="60" alt="" title=""/></li>
      </ul>
      <p style="height:32px;text-align:right;">
        <input id="regButton" type="submit" value="同意以下注册条款并提交" name="submit" onClick="return chk_reg()"/>
        </li>
      </p>
      <p>
        <textarea name="textarea" id="textarea" cols="45" rows="5">
一、总则

1．1　用户应当同意本协议的条款并按照页面上的提示完成全部的注册程序。用户在进行注册程序过程中点击"同意"按钮即表示用户与本站达成协议，完全接受本协议项下的全部条款。
1．2　用户注册成功后，本站将给予每个用户一个用户帐号及相应的密码【普通用户权限】，该用户帐号和密码由用户负责保管；用户应当对以其用户帐号进行的所有活动和事件负法律责任。
1．3　用户可以使用本站各个频道单项服务，当用户使用本站各单项服务时，用户的使用行为视为其对该单项服务的服务条款以及本站在该单项服务中发出的各类公告的同意。
1．4　本站会员服务协议以及各个频道单项服务条款和公告可由本站随时更新，且无需另行通知。您在使用相关服务时,应关注并遵守其所适用的相关条款。
　　您在使用本站提供的各项服务之前，应仔细阅读本服务协议。如您不同意本服务协议及/或随时对其的修改，您可以主动取消本站提供的服务；您一旦使用本站服务，即视为您已了解并完全同意本服务协议各项内容，包括本站对服务协议随时所做的任何修改，并成为本站用户。

二、注册信息和隐私保护

2．1　本站帐号（即本站用户ID）的所有权归本站，用户完成注册申请手续后，获得本站帐号的使用权。用户应提供及时、详尽及准确的个人资料，并不断更新注册资料，符合及时、详尽准确的要求。所有原始键入的资料将引用为注册资料。如果因注册信息不真实而引起的问题，并对问题发生所带来的后果，本站不负任何责任。
2．2　用户不应将其帐号、密码转让或出借予他人使用。因黑客行为或用户的保管疏忽导致帐号、密码遭他人非法使用，本站不承担任何责任。
2．3　本站不对外公开或向第三方提供单个用户的注册资料，除非：
（1）事先获得用户的明确授权；
（2）只有透露你的个人资料，才能提供你所要求的产品和服务；
（3）根据有关的法律法规要求；
（4）按照相关政府主管部门的要求；
（5）为维护本站的合法权益。
2．4　在你注册本站帐户，使用其他本站产品或服务，访问本站网页, 或参加促销和有奖游戏时，本站会收集你的个人身份识别资料，并会将这些资料用于：改进为你提供的服务及网页内容。

三、使用规则

3．1　用户在使用本站服务时，必须遵守中华人民共和国相关法律法规的规定，用户应同意将不会利用本服务进行任何违法或不正当的活动，包括但不限于下列行为∶
（1）上载、展示、张贴、传播或以其它方式传送含有下列内容之一的信息：
1） 反对宪法所确定的基本原则的； 2） 危害国家安全，泄露国家秘密，颠覆国家政权，破坏国家统一的； 3） 损害国家荣誉和利益的； 4） 煽动民族仇恨、民族歧视、破坏民族团结的； 5） 破坏国家宗教政策，宣扬邪教和封建迷信的； 6） 散布谣言，扰乱社会秩序，破坏社会稳定的； 7） 散布淫秽、色情、赌博、暴力、凶杀、恐怖或者教唆犯罪的； 8） 侮辱或者诽谤他人，侵害他人合法权利的； 9） 含有虚假、有害、胁迫、侵害他人隐私、骚扰、侵害、中伤、粗俗、猥亵、或其它道德上令人反感的内容； 10） 含有中国法律、法规、规章、条例以及任何具有法律效力之规范所限制或禁止的其它内容的； 
（2）不得为任何非法目的而使用网络服务系统；
（3）不利用本站服务从事以下活动：
1) 未经允许，进入计算机信息网络或者使用计算机信息网络资源的；
2) 未经允许，对计算机信息网络功能进行删除、修改或者增加的；
3) 未经允许，对进入计算机信息网络中存储、处理或者传输的数据和应用程序进行删除、修改或者增加的；
4) 故意制作、传播计算机病毒等破坏性程序的；
5) 其他危害计算机信息网络安全的行为。
3．2　用户违反本协议或相关的服务条款的规定，导致或产生的任何第三方主张的任何索赔、要求或损失，包括合理的律师费，您同意赔偿本站，并使之免受损害。对此，本站有权视用户的行为性质，采取包括但不限于删除用户发布信息内容、暂停使用许可、终止服务、限制使用、回收本站帐号、追究法律责任等措施。对恶意注册本站帐号或利用本站帐号进行违法活动、捣乱、骚扰、欺骗、其他用户以及其他违反本协议的行为，本站有权回收其帐号。同时，本站会视司法部门的要求，协助调查。
3．3　用户不得对本服务任何部分或本服务之使用或获得，进行复制、拷贝、出售、转售或用于任何其它商业目的。
3．4　用户须对自己在使用本站服务过程中的行为承担法律责任。用户承担法律责任的形式包括但不限于：对受到侵害者进行赔偿，以及在本站首先承担了因用户行为导致的行政处罚或侵权损害赔偿责任后，用户应给予本站等额的赔偿。 

四、服务内容

4．1　本站网络服务的具体内容由本站根据实际情况提供。
4．2　除非本服务协议另有其它明示规定，本站所推出的新产品、新功能、新服务，均受到本服务协议之规范。
4．3　为使用本服务，您必须能够自行经有法律资格对您提供互联网接入服务的第三方，进入国际互联网，并应自行支付相关服务费用。此外，您必须自行配备及负责与国际联网连线所需之一切必要装备，包括计算机、数据机或其它存取装置。
4．4　鉴于网络服务的特殊性，用户同意本站有权不经事先通知，随时变更、中断或终止部分或全部的网络服务（包括收费网络服务）。本站不担保网络服务不会中断，对网络服务的及时性、安全性、准确性也都不作担保。
4．5　本站需要定期或不定期地对提供网络服务的平台或相关的设备进行检修或者维护，如因此类情况而造成网络服务（包括收费网络服务）在合理时间内的中断，本站无需为此承担任何责任。本站保留不经事先通知为维修保养、升级或其它目的暂停本服务任何部分的权利。
4．6　本服务或第三人可提供与其它国际互联网上之网站或资源之链接。由于本站无法控制这些网站及资源，您了解并同意，此类网站或资源是否可供利用，本站不予负责，存在或源于此类网站或资源之任何内容、广告、产品或其它资料，本站亦不予保证或负责。因使用或依赖任何此类网站或资源发布的或经由此类网站或资源获得的任何内容、商品或服务所产生的任何损害或损失，本站不承担任何责任。
4．7　用户明确同意其使用本站网络服务所存在的风险将完全由其自己承担。用户理解并接受下载或通过本站服务取得的任何信息资料取决于用户自己，并由其承担系统受损、资料丢失以及其它任何风险。本站对在服务网上得到的任何商品购物服务、交易进程、招聘信息，都不作担保。
4．8　本站有权于任何时间暂时或永久修改或终止本服务（或其任何部分），而无论其通知与否，本站对用户和任何第三人均无需承担任何责任。
4．9　终止服务
您同意本站得基于其自行之考虑，因任何理由，包含但不限于长时间未使用，或本站认为您已经违反本服务协议的文字及精神，终止您的密码、帐号或本服务之使用（或服务之任何部分），并将您在本服务内任何内容加以移除并删除。您同意依本服务协议任何规定提供之本服务，无需进行事先通知即可中断或终止，您承认并同意，本站可立即关闭或删除您的帐号及您帐号中所有相关信息及文件，及/或禁止继续使用前述文件或本服务。此外，您同意若本服务之使用被中断或终止或您的帐号及相关信息和文件被关闭或删除，本站对您或任何第三人均不承担任何责任。

五、青少年用户特别提示

青少年用户必须遵守全国青少年网络文明公约：
要善于网上学习，不浏览不良信息；要诚实友好交流，不侮辱欺诈他人；要增强自护意识，不随意约会网友；要维护网络安全，不破坏网络秩序；要有益身心健康，不沉溺虚拟时空。

六、其他

6．1　本协议的订立、执行和解释及争议的解决均应适用中华人民共和国法律。
6．2　如双方就本协议内容或其执行发生任何争议，双方应尽量友好协商解决；协商不成时，任何一方均可向本站所在地的人民法院提起诉讼。
6．3　本站未行使或执行本服务协议任何权利或规定，不构成对前述权利或权利之放弃。
6．4　如本协议中的任何条款无论因何种原因完全或部分无效或不具有执行力，本协议的其余条款仍应有效并且有约束力。

请您在发现任何违反本服务协议以及其他任何单项服务的服务条款、本站各类公告之情形时，通知本站。</textarea>
      </p>
    </form>
  </div>
</div>
<script language="javascript">
<!--
$(document).ready(function(){ 


		var objImageValid=$("img[src^='<%=GetCurrentHost%>zb_system/function/c_validcode.asp?name=commentvalid']");
		if(objImageValid.size()>0){
			objImageValid.css("cursor","pointer");
			objImageValid.click( function() {
					objImageValid.attr("src","<%=GetCurrentHost%>zb_system/function/c_validcode.asp?name=commentvalid"+"&amp;random="+Math.random());
			} );
		};

});

var msg	;
var bname_m=false;
function init_reg(){
	msg=new Array(
	"请输入14位以内字符，允许汉字。",	
	"请输入8-14位字符，不允许空格。",
	"请重复输入一次密码。",
    "请输入电子邮箱地址。",	
	"只有同意注册条款才能完成注册。",
    "两次输入的密码不一致。",
    "邮箱地址不正确。",
    "同意注册条款。",
	"请输入您的网址。",
	"网址格式不正确。"
	)
	document.getElementById("d_uname").innerHTML=msg[0];	
	document.getElementById("d_upwd1").innerHTML=msg[1];
	document.getElementById("d_upwd2").innerHTML=msg[2];	
	document.getElementById("d_email").innerHTML=msg[3];		
    document.getElementById("d_site").innerHTML=msg[8];
       
	
}
init_reg();
function on_input(objname){
	var strtxt;
	var obj=document.getElementById(objname);
	obj.className="d_on";
	switch (objname){
		case "d_uname":
			strtxt=msg[0];
			break;		
		case "d_upwd1":
			strtxt=msg[1];
		break;
        case "d_upwd2":
			strtxt=msg[2];
		break;	
        case "d_email":
			strtxt=msg[3];
		break;
		case "d_site":
			strtxt=msg[8];
		break;
		
			
	}
	obj.innerHTML=strtxt;
}

function out_uname(){
	var obj=document.getElementById("d_uname");
	var str=sl(document.getElementById("uname").value);
	var chk=true;
	//alert(str);
	if (str<4 || str>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='用户名已经输入。';
		
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[0];
	}
	return chk;
}

function out_upwd1(){
	var obj=document.getElementById("d_upwd1");
	var str=document.getElementById("upwd").value;
	var chk=true;
	if (str=='' || str.length<8 || str.length>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='密码已经输入。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[1];
	}
	return chk;
}

function out_upwd2(){
	var obj=document.getElementById("d_upwd2");
	var str=document.getElementById("repassword").value;
	var chk=true;
	if (str!=document.getElementById("upwd").value||str==''){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='重复密码输入正确。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[5];
	}
	return chk;
}




function out_email(){
	var obj=document.getElementById("d_email");
	var str=document.getElementById("email").value;
	var chk=true;
	if (str==''|| !str.match(/^[\w\.\-]+@([\w\-]+\.)+[a-z]{2,4}$/ig)){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='电子邮箱地址已经输入。';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[6];
	}
	return chk;
}


function out_site(){
	var obj=$("#d_site");
	var str=$("#site").attr("value");
	var chk=true;
	if (str==''|| !str.match(/^[a-zA-Z]+:\/\/[a-zA-Z0-9\\_\\-\\.\\&\\?\/:=#\u4e00-\u9fa5]+?\/*$/ig)){chk=false}
	if (chk){
		obj.attr("class","d_ok");
		obj.html('网址已经输入。');
	}else{
		obj.attr("class","d_err");
		obj.html(msg[9]);
	}
	return chk;
}




function out_passregtext(){
	var chk=true;
	return chk;
}
function chk_reg(){
	var chk=true
	if (!out_uname()){chk=false;return false}	
	if (!out_upwd1()){chk=false;return false}
	if (!out_upwd2()){chk=false;return false}	
	if (!out_email()){chk=false;return false}
	if (!out_site()){chk=false;return false}
    if (!out_passregtext()){chk=false;return false}
	if(chk){
	//document.getElementById('regButton').disabled='disabled';
	var username=document.zblogform.uname.value;
	var password=document.zblogform.upwd.value;
	var repassword=document.zblogform.repassword.value;
    var email=document.zblogform.email.value;
	return true;
	//document.getElementById("reg").submit()
	}

}



function sl(st){
	sl1=st.length;
	strLen=0;
	for(i=0;i<sl1;i++){
		if(st.charCodeAt(i)>255) strLen+=2;
	 else strLen++;
	}
	return strLen;
}






function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9')  &&  (string.charAt(i) < 'a' || string.charAt(i) > 'z') &&  (string.charAt(i)!='-'))
{
return 1;
}
}
return 0;//pass
}

-->
</script> 
<script type="text/javascript">
function GEId(id){return document.getElementById(id);}
function DispPwdStrength(iN,sHL){
	if(iN>3){ iN=3;}
	for(var i=1;i<4;i++){
		var sHCR="ob_pws0";
		if(i<=iN){ sHCR=sHL;}
		if(iN>0){
		GEId("idSM"+i).className=sHCR;
		}
		//GEId("idSMT"+i).className="ob_pwfont2";
		if (iN>0){
			if (i<=iN){
			GEId("idSMT"+i).style.display=((i==iN)?"inline":"none");
			}
		}
		else{
		GEId("idSMT"+i).style.display=((i==iN)?"none":"inline");
		}
	}
}
/*密码强度 来自.Net Passport注册站*/
function EvalPwdStrength(sP){
	if(ClientSideStrongPassword(sP,gSimilarityMap,gDictionary)){
		DispPwdStrength(3,'ob_pws3');
	}else if(ClientSideMediumPassword(sP,gSimilarityMap,gDictionary)){
		DispPwdStrength(2,'ob_pws2');
	}else if(ClientSideWeakPassword(sP,gSimilarityMap,gDictionary)){
		DispPwdStrength(1,'ob_pws1');
	}else{
		DispPwdStrength(0,'ob_pws0');
	}
}


var kNoCanonicalCounterpart = 0;
var kCapitalLetter = 0;
var kSmallLetter = 1;
var kDigit = 2;
var kPunctuation = 3;
var kAlpha =  4;
var kCanonicalizeLettersOnly = true;
var kCananicalizeEverything = false;
var gDebugOutput = null;
var kDebugTraceLevelNone = 0;
var kDebugTraceLevelSuperDetail = 120;
var kDebugTraceLevelRealDetail = 100;
var kDebugTraceLevelAll = 80;
var kDebugTraceLevelMost = 60;
var kDebugTraceLevelFew = 40;
var kDebugTraceLevelRare = 20;
var gDebugTraceLevel = kDebugTraceLevelNone;
function DebugPrint()
{
var string = "";
if (gDebugTraceLevel && gDebugOutput &&
DebugPrint.arguments && (DebugPrint.arguments.length > 1) && (DebugPrint.arguments[0] <= gDebugTraceLevel))
{
for(var index = 1; index < DebugPrint.arguments.length; index++)
{
string += DebugPrint.arguments[index] + " ";
}
string += "<br>\n";
gDebugOutput(string);
}
}
function CSimilarityMap()
{
this.m_elements = "";
this.m_canonicalCounterparts = "";
}
function SimilarityMap_Add(element, canonicalCounterpart)
{
this.m_elements += element;
this.m_canonicalCounterparts += canonicalCounterpart;
}
function SimilarityMap_Lookup(element)
{
var canonicalCounterpart = kNoCanonicalCounterpart;
var index = this.m_elements.indexOf(element);
if (index >= 0)
{
canonicalCounterpart = this.m_canonicalCounterparts.charAt(index);
}
else
{
}
return canonicalCounterpart;
}
function SimilarityMap_GetCount()
{
return this.m_elements.length;
}
CSimilarityMap.prototype.Add = SimilarityMap_Add;
CSimilarityMap.prototype.Lookup = SimilarityMap_Lookup;
CSimilarityMap.prototype.GetCount = SimilarityMap_GetCount;
function CDictionaryEntry(length, wordlist)
{
this.m_length = length;
this.m_wordlist = wordlist;
}
function DictionaryEntry_Lookup(strWord)
{
var fFound = false;
if (strWord.length == this.m_length)
{
var nFirst = 0;
var nLast = this.m_wordlist.length - 1;
while( nFirst <= nLast )
{
var nCurrent = Math.floor((nFirst + nLast)/2);
if( strWord == this.m_wordlist[nCurrent])
{
fFound = true;
break;
}
else if ( strWord > this.m_wordlist[nCurrent])
{
nLast = nCurrent - 1;
}
else
{
nFirst = nCurrent + 1;
}
}
}

return fFound;
}
CDictionaryEntry.prototype.Lookup = DictionaryEntry_Lookup;
function CDictionary()
{
this.m_entries = new Array()
}
function Dictionary_Lookup(strWord)
{
for (var index = 0; index < this.m_entries.length; index++)
{
if (this.m_entries[index].Lookup(strWord))
{
return true;
}
}
}
function Dictionary_Add(length, wordlist)
{
var iL=this.m_entries.length;
var cD=new CDictionaryEntry(length, wordlist)
this.m_entries[iL]=cD;
}
CDictionary.prototype.Lookup = Dictionary_Lookup;
CDictionary.prototype.Add = Dictionary_Add;
var gSimilarityMap = new CSimilarityMap();
var gDictionary = new CDictionary();
function CharacterSetChecks(type, fResult)
{
this.type = type;
this.fResult = fResult;
}
function isctype(character, type, nDebugLevel)
{
var fResult = false;
switch(type)
{
case kCapitalLetter:
if((character >= 'A') && (character <= 'Z'))
{
fResult = true;
}
break;
case kSmallLetter:
if ((character >= 'a') && (character <= 'z'))
{
fResult = true;
}
break;
case kDigit:
if ((character >= '0') && (character <= '9'))
{
fResult = true;
}
break;
case kPunctuation:
if ("!@#$%^&*()_+-='\";:[{]}\|.>,</?`~".indexOf(character) >= 0)
{
fResult = true;
}
break;
case kAlpha:
if (isctype(character, kCapitalLetter) || isctype(character, kSmallLetter))
{
fResult = true;
}
break;
default:
break;
}

return fResult;
}
function CanonicalizeWord(strWord, similarityMap, fLettersOnly)
{
var canonicalCounterpart = kNoCanonicalCounterpart;
var strCanonicalizedWord = "";
var nStringLength = 0;
if ((strWord != null) && (strWord.length > 0))
{
strCanonicalizedWord = strWord;
strCanonicalizedWord = strCanonicalizedWord.toLowerCase();

if (similarityMap.GetCount() > 0)
{
nStringLength = strCanonicalizedWord.length;

for(var index = 0; index < nStringLength; index++)
{
if (fLettersOnly && !isctype(strCanonicalizedWord.charAt(index), kSmallLetter, kDebugTraceLevelSuperDetail))
{
continue;
}

canonicalCounterpart = similarityMap.Lookup(strCanonicalizedWord.charAt(index));
if (canonicalCounterpart != kNoCanonicalCounterpart)
{
strCanonicalizedWord = strCanonicalizedWord.substring(0, index) + canonicalCounterpart +
strCanonicalizedWord.substring(index + 1, nStringLength);
}
}
}
}
return strCanonicalizedWord;
}
function IsLongEnough(strWord, nAtLeastThisLong)
{
if ((strWord == null) || isNaN(nAtLeastThisLong))
{
return false;
}
else if (strWord.length < nAtLeastThisLong)
{
return false;
}

return true;
}
function SpansEnoughCharacterSets(strWord, nAtLeastThisMany)
{
var nCharSets = 0;
var characterSetChecks = new Array(
new CharacterSetChecks(kCapitalLetter, false),
new CharacterSetChecks(kSmallLetter, false),
new CharacterSetChecks(kDigit, false),
new CharacterSetChecks(kPunctuation, false)
);
if ((strWord == null) || isNaN(nAtLeastThisMany))
{
return false;
}

for(var index = 0; index < strWord.length; index++)
{
for(var nCharSet = 0; nCharSet < characterSetChecks.length;nCharSet++)
{
if (!characterSetChecks[nCharSet].fResult && isctype(strWord.charAt(index), characterSetChecks[nCharSet].type, kDebugTraceLevelAll))
{
characterSetChecks[nCharSet].fResult = true;
break;
}
}
}
for(var nCharSet = 0; nCharSet < characterSetChecks.length;nCharSet++)
{
if (characterSetChecks[nCharSet].fResult)
{
nCharSets++;
}
}

if (nCharSets < nAtLeastThisMany)
{
return false;
}

return true;
}
function FoundInDictionary(strWord, similarityMap, dictionary)
{
var strCanonicalizedWord = "";

if((strWord == null) || (similarityMap == null) || (dictionary == null))
{
return true;
}
strCanonicalizedWord = CanonicalizeWord(strWord, similarityMap, kCanonicalizeLettersOnly);

if (dictionary.Lookup(strCanonicalizedWord))
{
return true;
}

return false;
}
function IsCloseVariationOfAWordInDictionary(strWord, threshold, similarityMap, dictionary)
{
var strCanonicalizedWord = "";
var nMinimumMeaningfulMatchLength = 0;

if((strWord == null) || isNaN(threshold) || (similarityMap == null) || (dictionary == null))
{
return true;
}
strCanonicalizedWord = CanonicalizeWord(strWord, similarityMap, kCananicalizeEverything);
nMinimumMeaningfulMatchLength = Math.floor((threshold) * strCanonicalizedWord.length);
for (var nSubStringLength = strCanonicalizedWord.length; nSubStringLength >= nMinimumMeaningfulMatchLength; nSubStringLength--)
{
for(var nSubStringStart = 0; (nSubStringStart + nMinimumMeaningfulMatchLength) < strCanonicalizedWord.length; nSubStringStart++)
{
var strSubWord = strCanonicalizedWord.substr(nSubStringStart, nSubStringLength);

if (dictionary.Lookup(strSubWord))
{
return true;
}
}
}
return false;
}
function Init()
{
gSimilarityMap.Add('3', 'e');
gSimilarityMap.Add('x', 'k');
gSimilarityMap.Add('5', 's');
gSimilarityMap.Add('$', 's');
gSimilarityMap.Add('6', 'g');
gSimilarityMap.Add('7', 't');
gSimilarityMap.Add('8', 'b');
gSimilarityMap.Add('|', 'l');
gSimilarityMap.Add('9', 'g');
gSimilarityMap.Add('+', 't');
gSimilarityMap.Add('@', 'a');
gSimilarityMap.Add('0', 'o');
gSimilarityMap.Add('1', 'l');
gSimilarityMap.Add('2', 'z');
gSimilarityMap.Add('!', 'i');
gDictionary.Add(3,
"oat|not|ken|keg|ham|hal|gas|cpu|cit|bop|bah".split("|"));
gDictionary.Add(4,
"zeus|ymca|yang|yaco|work|word|wool|will|viva|vito|vita|visa|vent|vain|uucp|util|utah|unix|trek|town|torn|tina|time|tier|tied|tidy|tide|thud|test|tess|tech|tara|tape|tapa|taos|tami|tall|tale|spit|sole|sold|soil|soft|sofa|soap|slav|slat|slap|slam|shit|sean|saud|sash|sara|sand|sail|said|sago|sage|saga|safe|ruth|russ|rusk|rush|ruse|runt|rung|rune|rove|rose|root|rick|rich|rice|reap|ream|rata|rare|ramp|prod|pork|pete|penn|penh|pend|pass|pang|pane|pale|orca|open|olin|olga|oldy|olav|olaf|okra|okay|ohio|oath|numb|null|nude|note|nosy|nose|nita|next|news|ness|nasa|mike|mets|mess|math|mash|mary|mars|mark|mara|mail|maid|mack|lyre|lyra|lyon|lynx|lynn|lucy|love|lose|lori|lois|lock|lisp|lisa|leah|lass|lash|lara|lank|lane|lana|kink|keri|kemp|kelp|keep|keen|kate|karl|june|judy|judo|judd|jody|jill|jean|jane|isis|iowa|inna|holm|help|hast|half|hale|hack|gust|gush|guru|gosh|gory|golf|glee|gina|germ|gatt|gash|gary|game|fred|fowl|ford|flea|flax|flaw|finn|fink|film|fill|file|erin|emit|elmo|easy|done|disk|disc|diet|dial|dawn|dave|data|dana|damn|dame|crab|cozy|coke|city|cite|chem|chat|cats|burl|bred|bill|bilk|bile|bike|beth|beta|benz|beau|bath|bass|bart|bank|bake|bait|bail|aria|anne|anna|andy|alex|abcd".split("|"));
gDictionary.Add(5,
"yacht|xerox|wilma|willy|wendy|wendi|water|warez|vitro|vital|vitae|vista|visor|vicky|venus|venom|value|ultra|u.s.a|tubas|tress|tramp|trait|tracy|traci|toxic|tiger|tidal|thumb|texas|test2|test1|terse|terry|tardy|tappa|tapis|tapir|taper|tanya|tansy|tammy|tamie|taint|sybil|suzie|susie|susan|super|steph|stacy|staci|spark|sonya|sonia|solar|soggy|sofia|smile|slave|slate|slash|slant|slang|simon|shiva|shell|shark|sharc|shack|scrim|screw|scott|scorn|score|scoot|scoop|scold|scoff|saxon|saucy|satan|sasha|sarah|sandy|sable|rural|rupee|runty|runny|runic|runge|rules|ruben|royal|route|rouse|roses|rolex|robyn|robot|robin|ridge|rhode|revel|renee|ranch|rally|radio|quark|quake|quail|power|polly|polis|polio|pluto|plane|pizza|photo|phone|peter|perry|penna|penis|paula|patty|parse|paris|parch|paper|panic|panel|olive|olden|okapi|oasis|oaken|nurse|notre|notch|nancy|nagel|mouse|moose|mogul|modem|merry|megan|mckee|mckay|mcgee|mccoy|marty|marni|mario|maria|marcy|marci|maint|maine|magog|magic|lyric|lyons|lynne|lynch|louis|lorry|loris|lorin|loren|linda|light|lewis|leroy|laura|later|lasso|laser|larry|ladle|kinky|keyes|kerry|kerri|kelly|keith|kazoo|kayla|kathy|karie|karen|julie|julia|joyce|jenny|jenni|japan|janie|janet|james|irene|inane|impel|idaho|horus|horse|honey|honda|holly|hello|heidi|hasty|haste|hamal|halve|haley|hague|hager|hagen|hades|guest|guess|gucci|group|grahm|gouge|gorse|gorky|glean|gleam|glaze|ghoul|ghost|gauss|gauge|gaudy|gator|gases|games|freer|fovea|float|fiona|finny|filly|field|erika|erica|enter|enemy|empty|emily|email|elmer|ellis|ellen|eight|eerie|edwin|edges|eatme|earth|eager|dulce|donor|donna|diane|diana|delay|defoe|david|danny|daisy|cuzco|cubit|cozen|coypu|coyly|cowry|condo|class|cindy|cigar|chess|cathy|carry|carol|carla|caret|caren|candy|candi|burma|burly|burke|brian|breed|borax|booze|booty|bloom|blood|bitch|bilge|bilbo|betty|beryl|becky|beach|bathe|batch|basic|bantu|banks|banjo|baird|baggy|azure|arrow|array|april|anita|angie|amber|amaze|alpha|alisa|alike|align|alice|alias|album|alamo|aires|admin|adept|adele|addle|addis|added|acura|abyss|abcde|1701d|123go|!@#$%".split("|"));
gDictionary.Add(6,
"yankee|yamaha|yakima|y7u8i9|xyzxyz|wombat|wizard|wilson|willie|weenie|warren|visual|virgin|viking|venous|venice|venial|vasant|vagina|ursula|urchin|uranus|uphill|umpire|u.s.a.|tuttle|trisha|trails|tracie|toyota|tomato|toggle|tidbit|thorny|thomas|terror|tennis|taylor|target|tardis|tappet|taoist|tannin|tanner|tanker|tamara|system|surfer|summer|subway|stacie|stacey|spring|sondra|solemn|soleil|solder|solace|soiree|soften|soffit|sodium|sodden|snoopy|snatch|smooch|smiles|slavic|slater|single|singer|simple|sherri|sharon|sharks|sesame|sensor|secret|second|season|search|scroll|scribe|scotty|scooby|schulz|school|scheme|saturn|sandra|sandal|saliva|saigon|sahara|safety|safari|sadism|saddle|sacral|russel|runyon|runway|runoff|runner|ronald|romano|rodent|ripple|riddle|ridden|reveal|return|remote|recess|recent|realty|really|reagan|raster|rascal|random|radish|radial|racoon|racket|racial|rachel|rabbit|qwerty|qawsed|puppet|puneet|public|prince|presto|praise|poster|polite|polish|policy|police|plover|pierre|phrase|photon|philip|persia|peoria|penmen|penman|pencil|peanut|parrot|parent|pardon|papers|pander|pamela|pallet|palace|oxford|outlaw|osiris|orwell|oregon|oracle|olivia|oliver|olefin|office|notion|notify|notice|notate|notary|noreen|nobody|nicole|newton|nevada|mutant|mozart|morley|monica|moguls|minsky|mickey|merlin|memory|mellon|meagan|mcneil|mcleod|mclean|mckeon|mchugh|mcgraw|mcgill|mccann|mccall|mccabe|mayfly|maxine|master|massif|maseru|marvin|markus|malcom|mailer|maiden|magpie|magnum|magnet|maggot|lorenz|lisbon|limpid|leslie|leland|latest|latera|latent|lascar|larkin|langur|landis|landau|lambda|kristy|kristi|krista|knight|kitten|kinney|kerrie|kernel|kermit|kennan|kelvin|kelsey|kelley|keller|keenan|katina|karina|kansas|juggle|judith|jsbach|joshua|joseph|johnny|joanne|joanna|jixian|jimmie|jimbob|jester|jeanne|jasmin|janice|jaguar|jackie|island|invest|instar|ingrid|ingres|impute|holmes|holman|hockey|hidden|hawaii|hasten|harvey|harold|hamlin|hamlet|halite|halide|haggle|haggis|hadron|hadley|hacker|gustav|gusset|gurkha|gurgle|guntis|guitar|graham|gospel|gorton|gorham|gorges|golfer|glassy|ginger|gibson|ghetto|german|george|gauche|gasify|gambol|gamble|gambit|friend|freest|fourth|format|flower|flaxen|flaunt|flakes|finley|finite|fillip|fillet|filler|filled|fermat|fender|fatten|fatima|fathom|father|evelyn|euclid|estate|enzyme|engine|employ|emboss|elanor|elaine|eileen|eighty|eighth|effect|efface|eeyore|eerily|edwina|easier|durkin|durkee|during|durham|duress|duncan|donner|donkey|donate|donald|domino|disney|dieter|device|denise|deluge|delete|debbie|deaden|ddurer|dapper|daniel|dancer|damask|dakota|daemon|cuvier|cuddly|cuddle|cuckoo|cretin|create|cozier|coyote|cowpox|cooper|cookie|connie|coneck|condom|coffee|citrus|citron|citric|circus|charon|change|censor|cement|celtic|cecily|cayuga|catnip|catkin|cation|castle|carson|carrot|carrie|carole|carmen|caress|cantor|burley|burlap|buried|burial|brenda|bremen|breezy|breeze|breech|brandy|brandi|border|borden|borate|bloody|bishop|bilbao|bikini|bigred|betsie|berman|berlin|bedbug|became|beavis|beaver|beauty|beater|batman|bathos|barony|barber|baobab|bantus|banter|bantam|banish|bangui|bangor|bangle|bandit|banana|bakery|bailey|bahama|bagley|badass|aztecs|azsxdc|athena|asylum|arthur|arrest|arrear|arrack|arlene|anvils|answer|angela|andrea|anchor|analog|amazon|amanda|alison|alight|alicia|albino|albert|albeit|albany|alaska|adrian|adelia|adduce|addict|addend|accrue|access|abcdef|abcabc|abc123|a1b2c3|a12345|@#$%^&|7y8u9i|1qw23e|1q2w3e|1p2o3i|1a2b3c|123abc|10sne1|0p9o8i|!@#$%^".split("|"));
gDictionary.Add(7,
"yolanda|wyoming|winston|william|whitney|whiting|whatnot|vitriol|vitrify|vitiate|vitamin|visitor|village|vertigo|vermont|venturi|venture|ventral|venison|valerie|utility|upgrade|unknown|unicorn|unhappy|trivial|torrent|tinfoil|tiffany|tidings|thunder|thistle|theresa|test123|terrify|teleost|tarbell|taproot|tapping|tapioca|tantrum|tantric|tanning|takeoff|swearer|suzanne|susanne|support|success|student|squires|sossina|soldier|sojourn|soignee|sodding|smother|slavish|slavery|slander|shuttle|shivers|shirley|sheldon|shannon|service|seattle|scooter|scissor|science|scholar|scamper|satisfy|sarcasm|salerno|sailing|saguaro|saginaw|sagging|saffron|sabrina|russell|rupture|running|runneth|rosebud|receipt|rebecca|realtor|raleigh|rainbow|quarrel|quality|qualify|pumpkin|protect|program|profile|profess|profane|private|prelude|porsche|politic|playboy|phoenix|persona|persian|perseus|perseid|perplex|penguin|pendant|parapet|panoply|panning|panicle|panicky|pangaea|pandora|palette|pacific|olivier|olduvai|oldster|okinawa|oakwood|nyquist|nursery|numeric|number1|nullify|nucleus|nuclear|notused|nothing|newyork|network|neptune|montana|minimum|michele|michael|merriam|mercury|melissa|mcnulty|mcnally|mcmahon|mckenna|mcguire|mcgrath|mcgowan|mcelroy|mcclure|mcclain|mccarty|mcbride|mcadams|mbabane|mayoral|maurice|marimba|manhole|manager|mammoth|malcolm|malaria|mailbox|magnify|magneto|losable|lorinda|loretta|lorelei|lockout|lioness|limpkin|library|lazarus|lathrop|lateran|lateral|kristin|kristie|kristen|kinsman|kingdom|kennedy|kendall|kellogg|keelson|katrina|jupiter|judaism|judaica|jessica|janeiro|inspire|inspect|insofar|ingress|indiana|include|impetus|imperil|holmium|holmdel|herbert|heather|headmen|headman|harmony|handily|hamburg|halifax|halibut|halfway|haggard|hafnium|hadrian|gustave|gunther|gunshot|gryphon|gosling|goshawk|gorilla|gleason|glacier|ghostly|germane|georgia|geology|gaseous|gascony|gardner|gabriel|freeway|fourier|flowers|florida|fishers|finnish|finland|ferrari|felicia|feather|fatigue|fairway|express|expound|emulate|empress|empower|emitted|emerald|embrace|embower|ellwood|ellison|egghead|durward|durrell|drought|donning|donahue|digital|develop|desiree|default|deborah|damming|cynthia|cyanate|cutworm|cutting|cuddles|cubicle|crystal|coxcomb|cowslip|cowpony|cowpoke|console|conquer|connect|comrade|compton|collins|cluster|claudia|classic|citroen|citrate|citizen|citadel|cistern|christy|chester|charles|charity|celtics|celsius|catlike|cathode|carroll|carrion|careful|carbine|carbide|caraway|caravan|camille|burmese|burgess|bridget|breccia|bradley|bopping|blondie|bilayer|beverly|bernard|bermuda|berlitz|berlioz|beowulf|beloved|because|beatnik|beatles|beatify|bassoon|bartman|baroque|barbara|baptism|banshee|banquet|bannock|banning|bananas|bainite|bailiff|bahrein|bagpipe|baghdad|bagging|bacchus|asshole|arrange|arraign|arragon|arizona|ariadne|annette|animals|anatomy|anatole|amatory|amateur|amadeus|allison|alimony|aliases|algebra|albumin|alberto|alberta|albania|alameda|aladdin|alabama|airport|airpark|airfoil|airflow|airfare|airdrop|adenoma|adenine|address|addison|accrual|acclaim|academy|abcdefg|!@#$%^&".split("|"));
gDictionary.Add(8,
"yosemite|y7u8i9o0|wormwood|woodwind|whistler|whatever|warcraft|vitreous|virginia|veronica|venomous|trombone|transfer|tortoise|tientsin|tideland|ticklish|thailand|testtest|tertiary|terrific|terminal|telegram|tarragon|tapeworm|tapestry|tanzania|tantalus|tantalum|sysadmin|symmetry|sunshine|strangle|startrek|springer|sparrows|somebody|solecism|soldiery|softwood|software|softball|socrates|slatting|slapping|slapdash|slamming|simpsons|serenity|security|schwartz|sanctity|sanctify|samantha|salesman|sailfish|sailboat|sagittal|sagacity|sabotage|rushmore|rosemary|rochelle|robotics|reverend|regional|raindrop|rachelle|qwertyui|qwerasdf|qawsedrf|q1w2e3r4|protozoa|prodding|princess|precious|politics|politico|plymouth|pershing|penitent|penelope|pendulum|patricia|password|passport|paranoia|panorama|panicked|pandemic|pandanus|pakistan|painless|operator|olivetti|oleander|oklahoma|notocord|notebook|notarize|nebraska|napoleon|missouri|michigan|michelle|mesmeric|mercedes|mcmullen|mcmillan|mcknight|mckinney|mckinley|mckesson|mckenzie|mcintyre|mcintosh|mcgregor|mcgovern|mcginnis|mcfadden|mcdowell|mcdonald|mcdaniel|mcconnel|mccauley|mccarthy|mccallum|mayapple|masonite|maryland|marjoram|marinate|marietta|maneuver|mandamus|maledict|maladapt|magnuson|magnolia|magnetic|lyrebird|lymphoma|lorraine|lionking|linoleum|limitate|limerick|laterite|landmass|landmark|landlord|landlady|landhold|landfill|kristine|kirkland|kingston|kimberly|khartoum|keystone|kentucky|keeshond|kathrine|kathleen|jubilant|joystick|jennifer|jacobsen|irishman|interpol|internet|insulate|instinct|instable|insomnia|insolent|insolate|inactive|imperial|iloveyou|illinois|hydrogen|hutchins|homework|hologram|holocene|hibernia|hiawatha|heinlein|hebrides|headlong|headline|headland|hastings|hamilton|halftone|halfback|hagstrom|gunsling|gunpoint|gumption|gorgeous|glaucous|glaucoma|glassine|ginnegan|ghoulish|gertrude|geometry|geometer|garfield|gamesman|gamecock|fungible|function|frighten|freetown|foxglove|fourteen|foursome|forsythe|football|flaxseed|flautist|flatworm|flatware|fidelity|exposure|eternity|enthrone|enthrall|enthalpy|entendre|entangle|engineer|emulsion|emulsify|emporium|employer|employee|employed|emmanuel|elliptic|elephant|einstein|eighteen|duration|donnelly|dominion|dlmhurst|delegate|delaware|december|deadwood|deadlock|deadline|deadhead|danielle|cyanamid|cucumber|cristina|criminal|creosote|creation|cowpunch|couscous|conquest|comrades|computer|comprise|compress|colorado|clusters|citation|charming|cerulean|cenozoic|cemetery|cellular|catskill|cationic|catholic|cathodic|catheter|cascades|carriage|caroline|carolina|carefree|cardinal|burgundy|burglary|bumbling|broadway|breeches|bordello|bordeaux|bilinear|bilabial|bernardo|berliner|berkeley|bedazzle|beaumont|beatrice|beatific|bathrobe|baronial|baritone|bankrupt|banister|bakelite|azsxdcfv|asdfqwer|arkansas|appraise|apposite|anything|angerine|ancestry|ancestor|anatomic|anathema|ambiance|alphabet|albright|albrecht|alberich|albacore|alastair|alacrity|airspace|airplane|airfield|airedale|aircraft|airbrush|airborne|aerobics|adrianna|adelaide|additive|addition|addendum|accouter|academic|academia|abcdefgh|abcd1234|a1b2c3d4|7y8u9i0o|7890yuio|1234qwer|0p9o8i7u|0987poiu|!@#$%^&*".split("|"));
gDictionary.Add(9,
"zimmerman|worldwide|wisconsin|wholesale|vitriolic|ventricle|ventilate|valentine|tidewater|testament|territory|tennessee|telephone|telepathy|teleology|telemetry|telemeter|telegraph|tarantula|tarantara|tangerine|supported|superuser|stuttgart|stratford|stephanie|solemnity|softcover|slaughter|slapstick|signature|sheffield|sarcastic|sanctuary|sagebrush|sagacious|runnymede|rochester|receptive|reception|racketeer|professor|princeton|pondering|politburo|policemen|policeman|persimmon|persevere|persecute|percolate|peninsula|penetrate|pendulous|paralytic|panoramic|panicking|panhandle|oligopoly|oligocene|oligarchy|olfactory|oldenburg|nutrition|nurturant|notorious|notoriety|minnesota|microsoft|mcpherson|mcfarland|mcdougall|mcdonnell|mcdermott|mccracken|mccormick|mcconnell|mccluskey|mcclellan|marijuana|malicious|magnitude|magnetron|magnetite|macintosh|lynchburg|louisiana|lissajous|limousine|limnology|landscape|landowner|kinshasha|kingsbury|kibbutzim|kennecott|jamestown|ironstone|invisible|invention|intuitive|intervene|intersect|inspector|insomniac|insolvent|insoluble|impetuous|imperious|imperfect|holocaust|hollywood|hollyhock|headphone|headlight|headdress|headcount|headboard|happening|hamburger|halverson|gustafson|gunpowder|glasswort|glassware|ghostlike|geometric|gaucherie|freewheel|freethink|freestone|foresight|foolproof|extension|expositor|establish|entertain|employing|emittance|ellsworth|elizabeth|eightieth|eightfold|eiderdown|dusenbury|dusenberg|donaldson|dominique|discovery|desperate|delegable|delectate|decompose|decompile|damnation|cutthroat|crabapple|cornelius|conqueror|connubial|commrades|citizenry|christine|christina|chemistry|cellulose|celluloid|catherine|carryover|burlesque|bloodshot|bloodshed|bloodroot|bloodline|bloodbath|bilingual|bilateral|bijective|bijection|bernadine|berkshire|beethoven|beatitude|bakhtiari|asymptote|asymmetry|apprehend|appraisal|apportion|ancestral|anatomist|alexander|albatross|alabaster|alabamian|adenosine|abcabcabc".split("|"));
gDictionary.Add(10,
"washington|volkswagen|topography|tessellate|temptation|telephonic|telepathic|telemetric|telegraphy|tantamount|superstage|slanderous|salamander|qwertyuiop|polynomial|politician|phrasemake|photometry|photolytic|photolysis|photogenic|phosphorus|phosphoric|persiflage|persephone|perquisite|peninsular|penicillin|penetrable|panjandrum|oligoclase|oligarchic|oldsmobile|nottingham|noticeable|noteworthy|mcnaughton|mclaughlin|mccullough|mcallister|malconduct|maidenhair|limitation|lascivious|landowning|landlubber|landlocked|lamination|khrushchev|juggernaut|irrational|invariable|insouciant|insolvable|incomplete|impervious|impersonal|headmaster|glaswegian|geopolitic|geophysics|fourteenth|foursquare|expressive|expression|expository|exposition|enterprise|eightyfold|eighteenth|effaceable|donnybrook|delectable|decolonize|cuttlefish|cuttlebone|compromise|compressor|comprehend|cellophane|carruthers|california|burlington|burgundian|borderline|borderland|bloodstone|bloodstain|bloodhound|bijouterie|biharmonic|bernardino|beaujolais|basketball|bankruptcy|bangladesh|atmosphere|asymptotic|asymmetric|appreciate|apposition|ambassador|amateurish|alimentary|additional|accomplish|1q2w3e4r5t".split("|"));
gDictionary.Add(11,
"yellowstone|venturesome|territorial|telekinesis|sagittarius|safekeeping|politicking|policewoman|photometric|photography|phosphorous|perseverant|persecutory|persecution|penitential|pandemonium|mississippi|marketplace|magnificent|irremovable|interrogate|institution|inspiration|incompetent|impertinent|impersonate|impermeable|headquarter|hamiltonian|halfhearted|hagiography|geophysical|expressible|emptyhanded|eigenvector|deleterious|decollimate|decolletage|connecticut|comptroller|compressive|compression|catholicism|bloodstream|bakersfield|arrangeable|appreciable|anastomotic|albuquerque".split("|"));
gDictionary.Add(12,
"williamsburg|testamentary|qwerasdfzxcv|q1w2e3r4t5y6|perseverance|pennsylvania|penitentiary|malformation|liquefaction|interstitial|inconclusive|incomputable|incompletion|incompatible|incomparable|imperishable|impenetrable|headquarters|geometrician|ellipsometry|decomposable|decommission|compressible|burglarproof|bloodletting|bilharziasis|asynchronous|asymptomatic|ambidextrous|1q2w3e4r5t6y".split("|"));
gDictionary.Add(13,
"ventriloquist|ventriloquism|poliomyelitis|phosphorylate|oleomargarine|massachusetts|jitterbugging|interpolatory|inconceivable|imperturbable|impermissible|decomposition|comprehensive|comprehension".split("|"));
gDictionary.Add(14,
"slaughterhouse|irreproducible|incompressible|comprehensible|bremsstrahlung".split("|"));
gDictionary.Add(15,
"irreconciliable|instrumentation|incomprehension".split("|"));
gDictionary.Add(16,
"incomprehensible".split("|"));
}

function ClientSideStrongPassword()
{
return (IsLongEnough(ClientSideStrongPassword.arguments[0], "10") &&
SpansEnoughCharacterSets(ClientSideStrongPassword.arguments[0], "3") &&
(!(IsCloseVariationOfAWordInDictionary(ClientSideStrongPassword.arguments[0], "0.6",
ClientSideStrongPassword.arguments[1], ClientSideStrongPassword.arguments[2]))));
}

function ClientSideMediumPassword()
{
return (IsLongEnough(ClientSideMediumPassword.arguments[0], "9") &&
SpansEnoughCharacterSets(ClientSideMediumPassword.arguments[0], "2") &&
(!(FoundInDictionary(ClientSideMediumPassword.arguments[0], ClientSideMediumPassword.arguments[1],
ClientSideMediumPassword.arguments[2]))));
}

function ClientSideWeakPassword()
{
return (IsLongEnough(ClientSideWeakPassword.arguments[0], "8") ||
(!(IsLongEnough(ClientSideWeakPassword.arguments[0], "0"))));
}
</script>
<%
For Each sAction_Plugin_RegPage_End in Action_Plugin_RegPage_End
	If Not IsEmpty(sAction_Plugin_RegPage_End) Then Call Execute(sAction_Plugin_RegPage_End)
Next
%>
</body>
</html>