
document.write("<link rel='stylesheet' type='text/css' href='"+bloghost+"zb_system/admin/ueditor/themes/default/ueditor.css'/>");
document.write("<script type=\"text/javascript\" src=\""+bloghost+"zb_system/admin/ueditor/editor_all_min.js\"></script>");
(function(){
	var URL;
	URL = bloghost+'zb_system/admin/ueditor/';
	window.UEDITOR_CONFIG = {
		UEDITOR_HOME_URL:URL,
		toolbars:[ [ 'undo', 'redo', 'bold', 'italic', 'underline', 'forecolor', 'emotion','link','spechars','fullscreen']],
		initialContent:'<p></p>',
		initialStyle:'body{font-size:14px;font-family:微软雅黑,宋体,Arial,Helvetica,sans-serif;}',
		wordCount:false,
		elementPathEnabled:false,
		autoHeightEnabled:false,
		sourceEditor:"textarea",
		minFrameHeight:150
	}
}
)();
var UEComment;
$(document).ready(function(){
	if($("#txaArticle").length>0){
	var UEComment1 = new baidu.editor.ui.Editor();
	UEComment1.render('txaArticle');
	UEComment1.ready(function(){
		$('#txaArticle').css('height','auto');
	});
	UEComment=UEComment1;
	try{ReComment_CallBack.add(function(){
		//var m=UEComment.getContent();
		$("div[id='txaArticle']").remove();
		$("#txaArticle").show();
		UEComment.render('txaArticle');
		//UEComment.setContent(m);
	})}catch(e){}
	}
});

function VerifyMessage(){var d=document.getElementById("inpName").value;var c=document.getElementById("inpEmail").value;var b=document.getElementById("inpHomePage").value;var f;f=UEComment.getContent();if(d==""){alert(str01);return false}else{re=new RegExp("^[.A-Za-z0-9\u4e00-\u9fa5]+$");if(!re.test(d)){alert(str02);return false}}if(c==""){}else{re=new RegExp("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$");if(!re.test(c)){alert(str02);return false}}if(typeof(f)=="undefined"){alert(str03);return false}if(typeof(f)=="string"){if(f==""){alert(str03);return false}if(f.length>intMaxLen){alert(str03);return false}}document.getElementById("inpArticle").value=f;document.getElementById("inpLocation").value=parent.window.location.href;if(document.getElementById("frmSumbit").action.search("act=cmt")>0){strFormAction=document.getElementById("frmSumbit").action}var a=document.getElementById("chkRemember").checked;if(a==true){SaveRememberInfo()}else{SetCookie("chkRemember",a,365)}var e=$("#frmSumbit :submit").val();$("#frmSumbit :submit").val("Waiting...");$("#frmSumbit :submit").attr("disabled","disabled");$("#frmSumbit :submit").addClass("btnloading");$.post(document.getElementById("frmSumbit").action,{inpAjax:true,inpID:$("#inpId").val(),inpVerify:(document.getElementById("inpVerify")?$("#inpVerify").val():""),inpEmail:c,inpName:d,inpArticle:f,inpHomePage:b,inpRevID:intRevID},function(k){var h=k;if((h.search("faultCode")>0)&&(h.search("faultString")>0)){alert(h.match("<string>.+?</string>")[0].replace("<string>","").replace("</string>",""))}else{var g=Math.round(Math.random()*1000);var h=k;if(intRevID==0){$(h).insertBefore("#AjaxCommentEnd")}else{$(h).insertBefore("#AjaxCommentEnd"+intRevID);window.location="#cmt"+intRevID}$("#divAjaxComment"+g).fadeIn("slow");if(strFormAction){document.getElementById("frmSumbit").action=strFormAction}$("#txaArticle").val("");UEComment.setContent("");}if(document.getElementById("inpVerify")){$("#inpVerify").val("");var j=$("img[src^='"+str00+"zb_system/function/c_validcode.asp?name=commentvalid']");j.attr("src",str00+"zb_system/function/c_validcode.asp?name=commentvalid&random="+Math.random())}$("#frmSumbit :submit").removeClass("btnloading");$("#frmSumbit :submit").removeAttr("disabled");$("#frmSumbit :submit").val(e)});return false};