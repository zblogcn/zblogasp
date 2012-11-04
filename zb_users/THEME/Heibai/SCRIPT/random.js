// JavaScript Document
//lls for 20120323 to feifei tel:13034063791,mail:88172719@qq.com
var _lls_feifei_max=5;  //
var _lls_feifei_i=Math.floor((Math.random() * _lls_feifei_max)+1);
$(function(){
	$(".thumbnail img").each(function(){
		var _src=$(this).attr("src");
		if(_src.indexOf("random/tb")!=-1){
		$(this).attr("src","http://www.ilovewz.com/THEMES/Heibai/STYLE/images/random/tb"+_lls_feifei_i+".jpg");
		_lls_feifei_i++;
		if(_lls_feifei_i>_lls_feifei_max){_lls_feifei_i=1;}
		}
		});
	})