jQuery(document).ready(function($){
	$("a[rel=ignition]").colorbox({transition:'elastic',opacity:0.65,current:""}).hover(function(){
		$(this).append('<div id="cboxZoom"></div>');
		var i = $(this).find("img")
		var ol = i.offset().left + i.width() - 13;
		var ot = i.offset().top - 10;
		$("#cboxZoom").css({left: ol, top: ot});
	},
	function(){
		$(this).find("div:last").remove();
	});
	$("body").keydown(function(event) {
			if (event.keyCode == '192') {
				$("input.search-input").focus();
			}
	});
	$("#feeds").colorbox({width:"420px",height:"400px",opacity:0,inline:true,href:"#rss-feeds-box"});
	if($("#comments").length>0){
		$("#commentlist li").each(function() {
			var g = $(this).find("ul.children li .comment-content");
			$(this).find(".comment-reply-link:first").clone().appendTo(g);
		});
		$("#commentlist a.comment-reply-link").click(function(){
			var rename = $(this).parents(".comment-content").find(".comment-author-name").text();
			var reid = '"#comment-' + $(this).parents(".comment-content").find(".comment-id").text() + '"';
			$("#comment").append("&lt;a href=" + reid + "&gt;@" + rename + "&lt;/a&gt;&nbsp;").focus();
		});
		$("#cancel-comment-reply-link").click(function(){
			$("#comment").empty();
		});
		$("#comment").keydown(function(event) {
			if (event.keyCode == '13' && event.ctrlKey) {
				$("#submit").click();
			}
		});
	}
	if($("#loading").length>0){
		$(".kudos-post h4").each(function(){
			var nav = $(this).text();
			var title = $(this).attr("id");
			$(".kudos-nav ul").append("<li><a href='javascript:void(0);' rel='#" + title + "'>" + nav + "</a></li>");
		});
		$("#loading").remove();
	}
	if($(".kudos-frame-left").length>0){
		$(window).scroll(function(){ 
			var g = $(window).scrollTop();
			var p = g;
			var o = 0;
			if(g > 130){
				p = g - 100;
			}
			$(".kudos-frame-left").animate({top: p + "px"},{duration:1000, queue:false});
			if(g > 250){
				o = g - 250;
			}
			$(".kudos-frame-right").css("background-position","25px "+ o +"px");
		});
		$("#search-focus").click(function(){
			$("input.search-input").focus();							  
		});
	}
	$("a[rel*=#]").click(function(){
		var i = $(this).attr("rel");
		var c = $(i).offset().top;
		$("html,body").animate({scrollTop: c}, 1000);
	});
}); 