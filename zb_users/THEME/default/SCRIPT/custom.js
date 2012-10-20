function SidebarLoaded(){ 
	$("ul.ul-subcates").prev("a").before("<span class='sh'>-</span>");
	$("span.sh").click(function (){
		$(this).next().next("ul").toggle("fast");
	})
	.toggle(
		function () {
		$(this).html("+");
		},
		function () {
		$(this).html("-");
	});
}

//本条留言DomID,本条留言class,内容class,评论框DomID,指定父ID
function ReComment(comId,comClass,mClass,frmId,i){
	intRevID=i;
	var comm=$('#'+comId),frm=$('#'+frmId),cancel=$("#cancel-reply"),temp = $('#temp-frm');
	if ( ! comm.length || ! frm.length || ! cancel.length)return;
	if ( ! temp.length ) {
			var div = document.createElement('div');
			div.id = 'temp-frm';
			div.style.display = 'none';
			frm.before(div);
	}
	if (comm.has('.'+comClass).length){comm.find('.'+comClass).first().before(frm);}
	else comm.find('.'+mClass).first().append(frm);
	frm.addClass("reply-frm");

	cancel.show();
	cancel.click(function(){
		intRevID=0;
		var temp = $('#temp-frm'), frm=$('#'+frmId);
		if ( ! temp.length || ! frm.length )return;
		temp.before(frm);
		temp.remove();
		$(this).hide();
		frm.removeClass("reply-frm");
		return false;
	});
	try { $('#txaArticle').focus(); }
	catch(e) {}
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