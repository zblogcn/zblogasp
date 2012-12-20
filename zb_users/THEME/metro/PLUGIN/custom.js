//配色说明：第一个颜色为页面背景色，第二个为主色，依此类推。
var color_config=[
	{name:"默认",color:["#EEEEEE","#5EAAE4","#A3D0F2","#222222","#333333","#FFFFFF"]},
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
		$("<div>").attr({ title:config.name,class:"tc",onclick:"loadConfig(color_config["+i+"]);$('.active').removeClass('active');$(this).addClass('active');",style:"background-color:"+config.color[1]}).appendTo("#loadconfig");
	});

	var myUpload1 = $("#updatapic1").upload();
	myUpload1.set({
		name: 'bg.jpg',
		action: 'saveImage.asp',
		enctype: 'multipart/form-data',
		autoSubmit: true,
		onSelect: function (self, element) {
			this.autoSubmit = false;
			var re = new RegExp("(\.jpg){1}");
			if (this.filename()!==""){
				if (!re.test(this.filename())) {
					alert(this.filename()+"请上传jpg图片");
				}
				else {
					this.submit();
				}
			}
		},
		onComplete: function () {
			$('#bgpic_p').attr("src","../STYLE/images/bg.jpg?"+Math.random());
		}
	});
	var myUpload2 = $("#updatapic2").upload();
	myUpload2.set({
		name: 'headbg.jpg',
		action: 'saveImage.asp',
		enctype: 'multipart/form-data',
		autoSubmit: true,
		onSelect: function (self, element) {
			this.autoSubmit = false;
			var re = new RegExp("(\.jpg){1}");
			if (this.filename()!==""){
				if (!re.test(this.filename())) {
					alert(this.filename()+"请上传jpg图片");
				}
				else {
					this.submit();
				}
			}
		},
		onComplete: function () {
			$('#hbgpic_p').attr("src","../STYLE/images/headbg.jpg?"+Math.random());
		}
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

});