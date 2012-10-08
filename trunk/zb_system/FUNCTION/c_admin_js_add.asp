<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_admin_js.asp
'// 开始时间:    
'// 最后修改:    
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.ContentType="application/x-javascript" %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->

<% Response.Clear %>
var bloghost="<%=BlogHost%>";
var cookiespath="<%=CookiesPath()%>";
var str00="<%=BlogHost%>";



//*********************************************************
// 目的：    全选
// 输入：    无
// 返回：    无
//*********************************************************
function BatchSelectAll() {
	var aryChecks = document.getElementsByTagName("input");

	for (var i = 0; i < aryChecks.length; i++){
		if((aryChecks[i].type=="checkbox")&&(aryChecks[i].id.indexOf("edt")!==-1)){
			if(aryChecks[i].checked==true){
				aryChecks[i].checked=false;
			}
			else{
				aryChecks[i].checked=true;
			};
		}
	}
}
//*********************************************************




//*********************************************************
// 目的：    
// 输入：    无
// 返回：    无
//*********************************************************
function BatchDeleteAll(objEdit) {

	objEdit=document.getElementById(objEdit);
	objEdit.value="";
	var aryChecks = document.getElementsByTagName("input");
	for (var i = 0; i < aryChecks.length; i++){
		if((aryChecks[i].type=="checkbox")&&(aryChecks[i].id.indexOf("edt")!==-1)){
			if(aryChecks[i].checked){
				objEdit.value=aryChecks[i].value+","+objEdit.value;
			}
		}
	}

}
//*********************************************************








//*********************************************************
// 目的：    ActiveLeftMenu
// 输入：    无
// 返回：    无
//*********************************************************
function ActiveLeftMenu(name){

	name="#"+name;
	$("#leftmenu li").removeClass("on");
	$(name).parent().addClass("on");

}
//*********************************************************




//*********************************************************
// 目的：    ActiveTopMenu
// 输入：    无
// 返回：    无
//*********************************************************
function ActiveTopMenu(name){

	name="#"+name;
	$("#topmenu li").removeClass("on");
	$(name).addClass("on");

}
//*********************************************************





//*********************************************************
// 目的：    
//*********************************************************
function bmx2table(){

	//斑马线
	var tables=document.getElementsByTagName("table");
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");
		var b=false;
		if(cells[0].getElementsByTagName("th").length>0){

			cells[0].className="color1";
			for (var i = 1; i < cells.length; i++){
				if(b){
					cells[i].className="color2";
					b=false;
					cells[i].onmouseover=function(){
						this.className="color4";
					}
					cells[i].onmouseout=function(){
						this.className="color2";
					}
				}
				else{
					cells[i].className="color3";
					b=true;
					cells[i].onmouseover=function(){
						this.className="color4";
					}
					cells[i].onmouseout=function(){
						this.className="color3";
					}
				};
			};

		}else{
			var b=true;
			for (var i = 0; i < cells.length; i++){
				if(b){
					cells[i].className="color2";
					b=false;
					cells[i].onmouseover=function(){
						this.className="color4";
					}
					cells[i].onmouseout=function(){
						this.className="color2";
					}
				}
				else{
					cells[i].className="color3";
					b=true;
					cells[i].onmouseover=function(){
						this.className="color4";
					}
					cells[i].onmouseout=function(){
						this.className="color3";
					}
				};
			};
		
		}
	};
};
//*********************************************************




function Batch2Tip(s){$("#batch p").html(s)}
function BatchContinue(){$("#batch p").before("<iframe style='width:20px;height:20px;' frameborder='0' scrolling='no' src='<%=BlogHost%>zb_system/cmd.asp?act=batch'></iframe>");$("#batch img").remove();}
function BatchBegin(){};
function BatchEnd(){};
function BatchNotify(){notify($("#batch p").html())}
function BatchCancel(){$("#batch iframe").remove();$("#batch p").before("<iframe style='width:20px;height:20px;' frameborder='0' scrolling='no' src='<%=BlogHost%>zb_system/cmd.asp?act=batch&cancel=true'></iframe>");};





function changeCheckValue(obj){

	$(obj).toggleClass('imgcheck-on');

	if($(obj).hasClass('imgcheck-on')){
		$(obj).prev('input').val('True');
	}else{
		$(obj).prev('input').val('False');
	}

}



function notify(s){
	if (window.webkitNotifications) {
		if (window.webkitNotifications.checkPermission() == 0) {
			var zb_notifications = window.webkitNotifications.createNotification('<%=BlogHost%>zb_system/IMAGE/ADMIN/logo-16.png', '<%=ZC_MSG257%>', s);
			zb_notifications.onclick = function() {top.focus(),this.cancel();}
			zb_notifications.replaceId = 'Meteoric';
			setTimeout(function(){zb_notifications.cancel()},5000);
			zb_notifications.show();
		} else {
			window.webkitNotifications.requestPermission(notify);
		}
	} 
}


$(document).ready(function(){ 

	// Content box tabs:
	$('.content-box .content-box-content div.tab-content').hide(); // Hide the content divs
	$('ul.content-box-tabs li a.default-tab').addClass('current'); // Add the class "current" to the default tab
	$('.content-box-content div.default-tab').show(); // Show the div with class "default-tab"

	$('.content-box ul.content-box-tabs li a').click( // When a tab is clicked...
		function() { 
			$(this).parent().siblings().find("a").removeClass('current'); // Remove "current" class from all tabs
			$(this).addClass('current'); // Add class "current" to clicked tab
			var currentTab = $(this).attr('href'); // Set variable "currentTab" to the value of href of clicked tab
			$(currentTab).siblings().hide(); // Hide all content divs
			$(currentTab).show(); // Show the content div with the id equal to the id of clicked tab
			return false; 
		}
	);

});


$(document).ready(function(){ 

	bmx2table();

	if($('.SubMenu').find('span').length>0){
		//if($('#leftmenu').find('li.on').length>0){
		//	$('#leftmenu li.on').after('<li class="sub">'+$('.SubMenu').html()+'</li>');
		//}else{
			$('.SubMenu').show();
		//}
	}

	if(!(($.browser.msie)&&($.browser.version)=='6.0')){
		$('input.checkbox').css("display","none");
		$('input.checkbox[value="True"]').after('<span class="imgcheck imgcheck-on"></span>');
		$('input.checkbox[value="False"]').after('<span class="imgcheck"></span>');
	}else{
		$('input.checkbox').attr('readonly','readonly');
		$('input.checkbox').css('cursor','pointer');
		$('input.checkbox').click(function(){  if($(this).val()=='True'){$(this).val('False')}else{$(this).val('True')} })
	}

	$('span.imgcheck').click(function(){changeCheckValue(this)})

	$("#batch a").bind("click", function(){ BatchContinue();$("#batch p").html("<%=ZC_MSG109%>...");});

	$(".SubMenu span.m-right").parent().css({"float":"right"});

});

<%=Response_Plugin_Admin_Js_Add%>