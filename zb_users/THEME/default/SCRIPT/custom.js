	setTimeout(function(){ 
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
	},500);

	//本条留言DomID,类型(如，ul),评论框DomID,指定父ID；
	function moveForm(comId,type,frmId,i){
		intRevID=i;
		var comm=$('#'+comId),frm=$('#'+frmId),cancel=$("#cancel-reply"),temp = $('#temp-frm');
		if ( ! comm.length || ! frm.length || ! cancel.length)return;
		if ( ! temp.length ) {
				var div = document.createElement('div');
				div.id = 'temp-frm';
				div.style.display = 'none';
				frm.before(div);
		}
		if (comm.has(type).length)
		{
			comm.find(type).first().before(frm);
		}else comm.append(frm);

		cancel.show();
		cancel.click(function(){
			intRevID=0;
			var temp = $('#temp-frm'), frm=$('#'+frmId);
			if ( ! temp || ! frm )return;
			temp.before(frm);
			temp.remove();
			$(this).hide();
			return false;
		});
		try { $('#txaArticle').focus(); }
		catch(e) {}
		return false;
	}