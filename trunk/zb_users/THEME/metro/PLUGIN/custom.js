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

	$("#bgpx input ").removeAttr("checked");
	$("#hdbgpx input ").removeAttr("checked");
	$("#bgpx"+config.BodyBg[3]).attr("checked","checked")
	$("#hdbgpx"+config.HdBg[3]).attr("checked","checked")
	$( "#bgpx").buttonset("refresh");
	$( "#hdbgpx").buttonset("refresh");
	
	$( "#bgurl").val(config.BodyBg[1]);
	$( "#hdbgurl").val(config.HdBg[1]);

	if (config.BodyBg[2].indexOf("repeat")>-1){
		$( "#bodybg2r").attr("checked","checked")
	}else{$( "#bodybg2r").removeAttr("checked");}
	if (config.BodyBg[2].indexOf("fixed")>-1){
		$( "#bodybg2f").attr("checked","checked")
	}else{$( "#bodybg2f").removeAttr("checked");}

	if (config.HdBg[2].indexOf("repeat")>-1){
		$( "#hdbg2r").attr("checked","checked")
	}else{$( "#hdbg2r").removeAttr("checked");}
	if (config.HdBg[2].indexOf("fixed")>-1){
		$( "#hdbg2f").attr("checked","checked")
	}else{$( "#hdbg2f").removeAttr("checked");}



	if (config.BodyBg[5]==""){
		$('#bodybgc5').removeAttr("checked");
		$('#bodybgmain').hide("fast");
	}
	else	{
		$('#bodybgc5').attr("checked","checked");
		$('#bodybgmain').show("fast");
	}

	if (config.HdBg[0]==""){
		$('#hdbgc0').removeAttr("checked");
	}
	else	{
		$('#hdbgc0').attr("checked","checked");
	}
	if (config.HdBg[6]==""){
		$('#hdbgc6').removeAttr("checked");
		$('#hdbgmain').hide("fast");
	}
	else	{
		$('#hdbgc6').attr("checked","checked");
		$('#hdbgmain').show("fast");
	}

	$( "#hdbgph").val(config.HdBg[5]);
	
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
});