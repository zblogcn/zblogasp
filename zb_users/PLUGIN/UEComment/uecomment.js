//重写系统自带AJAX提交

function VerifyMessage() {

	var strName=document.getElementById("inpName").value;
	var strEmail=document.getElementById("inpEmail").value;
	var strHomePage=document.getElementById("inpHomePage").value;
	var strArticle;
	

	strArticle=UEComment.getContent();
	
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

	document.getElementById("inpArticle").value=strArticle;
	document.getElementById("inpLocation").value=parent.window.location.href;
	if(document.getElementById("frmSumbit").action.search("act=cmt")>0){
		strFormAction=document.getElementById("frmSumbit").action;
	}

	var bolRemember=document.getElementById("chkRemember").checked;

	if(bolRemember==true){
		SaveRememberInfo();
	}
	else{
		SetCookie("chkRemember",bolRemember,365);
	}

	var strSubmit=$("#frmSumbit :submit").val();
	$("#frmSumbit :submit").val("Waiting...");
	$("#frmSumbit :submit").attr("disabled","disabled");
	$("#frmSumbit :submit").addClass("btnloading");


	//ajax comment begin
	$.post(document.getElementById("frmSumbit").action,
		{
		"inpAjax":true,
		"inpID":$("#inpId").val(),
		"inpVerify":(document.getElementById("inpVerify")?$("#inpVerify").val():""),
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
				//$("#divAjaxComment"+i).fadeTo("normal", 0);
				//$("#divAjaxComment"+i).fadeTo("normal", 1);
				//$("#divAjaxComment"+i).show("slow");
				if(strFormAction){
					document.getElementById("frmSumbit").action=strFormAction;
				}
				$("#txaArticle").val("");
			}
			if(document.getElementById("inpVerify")){
				$("#inpVerify").val("");
				var objImageValid=$("img[src^='"+str00+"zb_system/function/c_validcode.asp?name=commentvalid']");
				objImageValid.attr("src",str00+"zb_system/function/c_validcode.asp?name=commentvalid"+"&random="+Math.random());
			}

			$("#frmSumbit :submit").removeClass("btnloading");
			$("#frmSumbit :submit").removeAttr("disabled");
			$("#frmSumbit :submit").val(strSubmit);

		}
	);



	return false;
	//ajax comment end

}