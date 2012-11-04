// JavaScript Document
//lls for 20120323 to feifei tel:13034063791,mail:88172719@qq.com
var _lls_feifei_max=5;  //
var _lls_feifei_i=Math.floor((Math.random() * _lls_feifei_max)+1);
// 滚屏
jQuery(document).ready(function($){
$('#roll_top').click(function(){$('html,body').animate({scrollTop: '0px'}, 800);}); 
$('#ct').click(function(){$('html,body').animate({scrollTop:$('#comments').offset().top}, 800);});
$('#fall').click(function(){$('html,body').animate({scrollTop:$('.footer').offset().top}, 800);});
});
// context
$(document).ready(function(){
$('.entry_box_s ').hover(
	function() {
		$(this).find('.context_t').stop(true,true).fadeIn();
	},
	function() {
		$(this).find('.context_t').stop(true,true).fadeOut();
	}
);
});
// more
$(document).ready(function(){
$('.entry_box').hover(
	function() {
		$(this).find('.archive_more').stop(true,true).fadeIn();
	},
	function() {
		$(this).find('.archive_more').stop(true,true).fadeOut();
	}
);

	$(".thumbnail img").each(function(){
		var _src=$(this).attr("src");
		if(_src.indexOf("random/tb")!=-1){
		$(this).attr("src","http://www.ilovewz.com/THEMES/Heibai/STYLE/images/random/tb"+_lls_feifei_i+".jpg");
		_lls_feifei_i++;
		if(_lls_feifei_i>_lls_feifei_max){_lls_feifei_i=1;}
		}
		});
});
