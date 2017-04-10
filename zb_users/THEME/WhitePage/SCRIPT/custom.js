$(document).ready(function(){
	var s=document.location;
	$("#divNavBar a").each(function(){
		if(this.href==s.toString().split("#")[0]){$(this).addClass("on");return false;}
	});
});


function ReComment_CallBack(){for(var i=0;i<=ReComment_CallBack.list.length-1;i++){ReComment_CallBack.list[i]()}}
ReComment_CallBack.list=[];
ReComment_CallBack.add=function(s){ReComment_CallBack.list.push(s)};


//重写了common.js里的同名函数
function RevertComment(i){
	$("#inpRevID").val(i);
	var frm=$('#divCommentPost'),cancel=$("#cancel-reply"),temp = $('#temp-frm');


	var div = document.createElement('div');
	div.id = 'temp-frm';
	div.style.display = 'none';
	frm.before(div);


	$('#AjaxCommentEnd'+i).before(frm);

	frm.addClass("reply-frm");
	$('#divCommentPost').find(":submit").bind("click", function(){ $("#cancel-reply").click();return false; });
	cancel.show();

	cancel.click(function(){
		$("#inpRevID").val(0);
		var temp = $('#temp-frm'), frm=$('#divCommentPost');
		if ( ! temp.length || ! frm.length )return;
		temp.before(frm);
		temp.remove();
		$(this).hide();
		frm.removeClass("reply-frm");
		$('#divCommentPost').find(":submit").unbind("click");
		ReComment_CallBack();
		return false;
	});
	try { $('#txaArticle').focus(); }
	catch(e) {}
	ReComment_CallBack();
	return false;
}

//重写GetComments，防止评论框消失
function GetComments(logid,page){
	$('span.commentspage').html("Waiting...");
	$.get(str00+"zb_system/cmd.asp?act=CommentGet&logid="+logid+"&page="+page, function(data){
		$("#cancel-reply").click();
		$('#AjaxCommentBegin').nextUntil('#AjaxCommentEnd').remove();
		$('#AjaxCommentEnd').before(data);
	});
}


//*********************************************************
// 目的：    验证信息
// 输入：    无
// 返回：    无
//*********************************************************

function VerifyMessage() {

	var strName=$("#inpName").val();
	var strEmail=$("#inpEmail").val();
	var strHomePage=$("#inpHomePage").val();
	var strArticle=$("#txaArticle").val();
	var strFormAction=$("#frmSumbit").attr("action");
	var intRevID=$("#inpRevID").val()==""?0:$("#inpRevID").val();

	if(strName==""){
		alert(str01);
		return false;
	}
	else{
		re = new RegExp("^[.A-Za-z0-9\u4e00-\u9fa5]+$");
		if (!re.test(strName)){
			alert(str02);
			return false;
		}
	}

	if(strEmail==""){
		//alert(str01);
		//return false;
	}
	else{
		re = new RegExp("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$");
		if (!re.test(strEmail)){
			alert(str02);
			return false;
		}
	}

	if(typeof(strArticle)=="undefined"){
		alert(str03);
		return false;
	}

	if(typeof(strArticle)=="string"){
		if(strArticle==""){
			alert(str03);
			return false;
		}
		if(strArticle.length>intMaxLen)
		{
			alert(str03);
			return false;
		}
	}

	$("#inpArticle").val(strArticle);

	var bolRemember=document.getElementById("chkRemember").checked;

	if(bolRemember==true){
		SaveRememberInfo();
	}
	else{
		SetCookie("chkRemember",bolRemember,365);
	}

	var strSubmit=$("#frmSumbit :submit").val();
	$("#frmSumbit :submit").val("Waiting...").attr("disabled","disabled").addClass("loading");


	//ajax comment begin
	$.post(strFormAction,
		{
		"inpAjax":true,
		"inpID":$("#inpId").val(),
		"inpVerify":($("#inpVerify").length>0?$("#inpVerify").val():""),
		"inpEmail":strEmail,
		"inpName":strName,
		"inpArticle":strArticle,
		"inpHomePage":strHomePage,
		"inpRevID":intRevID
		},
		function(data){
			var s =data;
			if((s.search("faultCode")>0)&&(s.search("faultString")>0))
			{
				alert(s.match("<string>.+?</string>")[0].replace("<string>","").replace("</string>",""))
			}
			else{
				var i=Math.round(Math.random()*1000);
				var s =data;
				if(intRevID==0){
					$(s).insertBefore("#AjaxCommentEnd");
				}else{
					$(s).insertBefore("#AjaxCommentEnd"+intRevID);
					window.location="#cmt"+intRevID
				}
				$("#divAjaxComment"+i).fadeIn("slow");
				$("#txaArticle").val("");
			}
			if($("#inpVerify").length>0){
				$("#inpVerify").val("");
				var objImageValid=$("img[src^='"+str00+"zb_system/function/c_validcode.asp?name=commentvalid']");
				objImageValid.attr("src",str00+"zb_system/function/c_validcode.asp?name=commentvalid"+"&random="+Math.random());
			}

			$("#frmSumbit :submit").removeClass("loading");
			$("#frmSumbit :submit").removeAttr("disabled");
			$("#frmSumbit :submit").val(strSubmit);

		}
	);

	return false;
	//ajax comment end

}
//*********************************************************
