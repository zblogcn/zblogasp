function ubbimg(a) {
	var b=/<img src=\"([^\"]*?)zb_users\/emotion\/([^\"]*?)\" data_ue_src=\"([^\"]*?)\">/gi;
	var s=a.replace(b,"[img]$2[/img]");
	return s;
}
$(function() {
	
	$("#txaArticle").before('<textarea id="ueditor" class="form-textarea resizable required"></textarea>');
	$("#txaArticle").hide();
	
	var editor = new baidu.editor.ui.Editor({
		toolbars:[["emotion"]],
		serialize : {
			//黑名单，编辑器会过滤掉以下标签
			blackList:{object:1, applet:1, input:1, meta:1, base:1, button:1, select:1, textarea:1, '#comment':1, 'map':1, 'area':1,p:1,br:1}
		}
	});
	editor.render('ueditor');

	$("input:submit").click(function(){
		$("#txaArticle").html(ubbimg(editor.getPlainTxt()));
		return VerifyMessage();
	});
});