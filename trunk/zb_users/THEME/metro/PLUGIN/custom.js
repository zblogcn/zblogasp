var theme_config={
	default:{ BodyBg:["#EEEEEE","images/bg.jpg","repeat","2","top",""],
					HdBg:["","images/headbg.jpg","repeat  fixed","1","top","100",""],
					color:["#5EAAE4","#A3D0F2","#222222","#333333","#FFFFFF"]
				},
	green:{	BodyBg:["#EEEEEE","images/bg.jpg","repeat  fixed","2","top","True"],
					HdBg:["","images/headbg.jpg","repeat  fixed","2","top","150",""],
					color:["#76923C","#C3D69B","#003300","#76923C","#FFFFFF"]
				}
};

function loadConfig(config){
	$('#bodybgc0').colorpicker("val",config.BodyBg[0]);
	$('#colorP1').colorpicker("val",config.color[0]);
	$('#colorP2').colorpicker("val",config.color[1]);
	$('#colorP3').colorpicker("val",config.color[2]);
	$('#colorP4').colorpicker("val",config.color[3]);
	$('#colorP5').colorpicker("val",config.color[4]);	
}

$(document).ready(function(){
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
			$('#bgpic_p').attr("src","../STYLE/images/bg.jpg");
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
			$('#hbgpic_p').attr("src","../STYLE/images/headbg.jpg");
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