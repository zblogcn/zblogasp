$(document).ready(function(){ 

$("#divMain2").prepend("<form class='search' name='edit' id='edit' method='post' enctype='multipart/form-data' action='"+bloghost+"cmd.asp?act=FileUpload'><p>上传主题zti文件:&nbsp;<input type='file' id='edtFileLoad' name='edtFileLoad' size='40' />&nbsp;&nbsp;&nbsp;&nbsp;<input type='submit' class='button' value='提交' name='B1' />&nbsp;&nbsp;<input class='button' type='reset' value='重置' name='B2' />&nbsp;</p></form>");


$(".theme").each(function(){
	var t=$(this).find("strong").html();
	var s="<p>";
	s=s+"<a href='id="+t+"' title='编辑该主题'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/application_edit.png'/></a>";
	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/theme_pack.asp?id="+t+"' title='导出该主题' target='_blank'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/download.png'/></a>";
	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='id="+t+"' title='更新该主题'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/refresh.png'/></a>";
	if($(this).hasClass("theme-other")){
		s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='id="+t+"' title='删除该主题'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/delete.png'/></a>";
	}
	s=s+"</p>";
	$(this).append(s);
	
});

});