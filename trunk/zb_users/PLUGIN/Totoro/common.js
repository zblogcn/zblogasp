totoro_statsbar("Loading Totoro....");

function totoro_cmmnginit(){
	$("tr").each(function(i){
		if(i!=0){
		var cmid=$(this).children("td:eq(1)").html();
		$(this).append("<td width=\"5%\" align=\"center\" id=\"totoro_" + cmid + "\"><a href=\"javascript:ThisCmIsSpam(" + cmid + ")\"><img width='16' src='"+str00+"zb_users/plugin/totoro/minus-shield.png' alt='[这是SPAM]' title='[这是SPAM]' /></a></td>");
		}else{
		$(this).append("<td width=\"5%\" align=\"center\"><a href=\"javascript:alert('点击[这是SPAM]将此评论中包含的网址加入TotoroⅡ黑词列表，并按照设置将其删除或进入审核')\">TotoroⅢ</a></td>");
		}
		});
	totoro_statsbar();
}

function totoro_statsbar(stats){
	if(stats){
		if(!$("#totoro_statsbar").length){
			$("body").prepend("<span id=\"totoro_statsbar\" style=\"position:absolute;top:10px;right:10px;height:15px;z-index:999;padding:5px 10px;background:#8B0000;color:#FFFFFF;font-size:12px;\">  </span>");
		}
		$("#totoro_statsbar").html(stats);
	}else{
		$("#totoro_statsbar").remove();
	}
}

function ThisCmIsSpam(cmid){
	$("#totoro_" + cmid).html("<span style=\"color:#800000;\">提交中</span>").prev().html("").prev().html("").prev().html("");
	$.post("../../zb_users/plugin/totoro/ajaxdel.asp", { act: "delcm", id: cmid } ,
	function(data){
	$("#totoro_" + cmid).html("<span style=\"color:#008000;\">已提交</span>").parent().children("td:eq(1)").html(data);
	});
}
