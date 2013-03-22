//配色说明：第一个颜色为页面背景色，第二个为主色，依此类推。
var color_config=[
	{name:"默认",color:["#EEEEEE","#5EAAE4","#A3D0F2","#222222","#333333","#FFFFFF"]},
	{name:"海蓝色",color:["#EEEEEE","#005f92","#A3D0F2","#222222","#333333","#FFFFFF"]},
	{name:"草绿色",color:["#EEEEEE","#76923C","#C3D69B","#003300","#76923C","#FFFFFF"]},
	{name:"黑色",color:["#d8d8d8","#3f3f3f","#bfbfbf","#7f7f7f","#595959","#f2f2f2"]},
	{name:"咖啡色",color:["#d8d8d8","#974806","#fac08f","#262626","#3f3f3f","#f2f2f2"]},
	{name:"紫色",color:["#ccc1d9","#5f497a","#b2a2c7","#262626","#3f3f3f","#f2f2f2"]}
];

function loadConfig(config){
	$('#bodybgc0').colorpicker("val",config.color[0]);
	$('#colorP1').colorpicker("val",config.color[1]);
	$('#colorP2').colorpicker("val",config.color[2]);
	$('#colorP3').colorpicker("val",config.color[3]);
	$('#colorP4').colorpicker("val",config.color[4]);
	$('#colorP5').colorpicker("val",config.color[5]);	
}

$(document).ready(function(){
	
	$.each(color_config, function(i,config){
		$("<div>").attr({"title":config.name,"class":"tc",onclick:"loadConfig(color_config["+i+"]);$('.active').removeClass('active');$(this).addClass('active');","style":"background-color:"+config.color[1]}).appendTo("#loadconfig");
	});

	$("#updatapic1,#updatapic2").parent().css("width","auto");

	$('#bodybgc0').colorpicker();
	$('#bgpx').buttonset();

	$('#bodybgc5').click(function(){
		if($(this).attr("checked")!==undefined){
			$('#bodybgmain').show("fast");}
		else{$('#bodybgmain').hide("fast");} 
	});

	$('#hdbgc6').click(function(){
		if($(this).attr("checked")!==undefined){
			$('#hdbgmain').show("fast");}
		else{$('#hdbgmain').hide("fast");} 
	});

	$( "#hdbgpx").buttonset();

	$('#colorP1').colorpicker();
	$('#colorP2').colorpicker();
	$('#colorP3').colorpicker();
	$('#colorP4').colorpicker();
	$('#colorP5').colorpicker();
	
	$( "#layoutset").buttonset();

	//var ueconfig = window.UEDITOR_CONFIG || {};
	$( "#updatapic1,#updatapic2").click(function(){
		this.callbacks = function(obj,win){
			for(key in obj) {
				$("#url_"+$(this).attr("id")).val(imagePath.replace(ZC_BLOG_HOST,'') + obj[key].url);
				$("#pic_"+$(this).attr("id")).attr("src",imagePath+ obj[key].url+"?"+Math.random());
			}
			win.close();
		};
		window.showModalDialog(ZC_BLOG_HOST+'zb_users/THEME/metro/PLUGIN/image.html',this,'dialogWidth:625px;dialogHeight:340px;resizable:no;scroll:no;status:no;');
	});

});