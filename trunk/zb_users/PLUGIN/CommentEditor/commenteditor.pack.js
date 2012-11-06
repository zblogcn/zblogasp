
//document.write("<link rel='stylesheet' type='text/css' href='"+bloghost+"zb_users/plugin/commenteditor/xheditor/xheditor-1.1.14-zh-cn.min.js'/>");
document.write("<script type=\"text/javascript\" src=\""+bloghost+"zb_users/plugin/commenteditor/xheditor/xheditor-1.1.14-zh-cn.min.js\"></scri"+"pt>");
document.write("<script type=\"text/javascript\" src=\""+bloghost+"zb_users/plugin/commenteditor/xheditor/xheditor_plugins/ubb.min.js\"></scri"+"pt>");
var xheditor;
$(document).ready(function(){
	if($("#txaArticle").length>0){
		renderxh()
	}
	try{ReComment_CallBack.add(function(){
		//var m=xheditor.getSource();
		$('#txaArticle').xheditor(false);
		renderxh();
		//xheditor.setSource(m);
	})}catch(e){}
});

function renderxh(){xheditor=$('#txaArticle').xheditor({tools:'Bold,Italic,Underline,Link,Unlink,Emot,|,About',width:'150',clickCancelDialog:false,forcePtag:false,submitID:'btnSumbit',internalStyle:false,inlineStyle:false,html5Upload:false,emotPath:bloghost+"zb_users/emotion/",skin:'nostyle',emotMark:true,beforeSetSource:ubb2html,beforeGetSource:html2ubb});}
function VerifyMessage(){var d=document.getElementById("inpName").value;var c=document.getElementById("inpEmail").value;var b=document.getElementById("inpHomePage").value;var f;f=xheditor.getSource();if(d==""){alert(str01);return false}else{re=new RegExp("^[.A-Za-z0-9\u4e00-\u9fa5]+$");if(!re.test(d)){alert(str02);return false}}if(c==""){}else{re=new RegExp("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$");if(!re.test(c)){alert(str02);return false}}if(typeof(f)=="undefined"){alert(str03);return false}if(typeof(f)=="string"){if(f==""){alert(str03);return false}if(f.length>intMaxLen){alert(str03);return false}}document.getElementById("inpArticle").value=f;document.getElementById("inpLocation").value=parent.window.location.href;if(document.getElementById("frmSumbit").action.search("act=cmt")>0){strFormAction=document.getElementById("frmSumbit").action}var a=document.getElementById("chkRemember").checked;if(a==true){SaveRememberInfo()}else{SetCookie("chkRemember",a,365)}var e=$("#frmSumbit :submit").val();$("#frmSumbit :submit").val("Waiting...");$("#frmSumbit :submit").attr("disabled","disabled");$("#frmSumbit :submit").addClass("btnloading");$.post(document.getElementById("frmSumbit").action,{inpAjax:true,inpID:$("#inpId").val(),inpVerify:(document.getElementById("inpVerify")?$("#inpVerify").val():""),inpEmail:c,inpName:d,inpArticle:f,inpHomePage:b,inpRevID:intRevID},function(k){var h=k;if((h.search("faultCode")>0)&&(h.search("faultString")>0)){alert(h.match("<string>.+?</string>")[0].replace("<string>","").replace("</string>",""))}else{var g=Math.round(Math.random()*1000);var h=k;if(intRevID==0){$(h).insertBefore("#AjaxCommentEnd")}else{$(h).insertBefore("#AjaxCommentEnd"+intRevID);window.location="#cmt"+intRevID}$("#divAjaxComment"+g).fadeIn("slow");if(strFormAction){document.getElementById("frmSumbit").action=strFormAction}$("#txaArticle").val("");xheditor.setSource("");}if(document.getElementById("inpVerify")){$("#inpVerify").val("");var j=$("img[src^='"+str00+"zb_system/function/c_validcode.asp?name=commentvalid']");j.attr("src",str00+"zb_system/function/c_validcode.asp?name=commentvalid&random="+Math.random())}$("#frmSumbit :submit").removeClass("btnloading");$("#frmSumbit :submit").removeAttr("disabled");$("#frmSumbit :submit").val(e)});return false};