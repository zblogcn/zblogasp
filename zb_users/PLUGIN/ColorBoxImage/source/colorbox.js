$(document).ready(function(){
$('img').each(function(){
	//var maxwidth=$(this).parent(this).width();
	if ($(this).parent("a").length == 0){
			//if (this.width > maxwidth){
			//this.width = maxwidth - 60;
			$(this).wrap("<a rel='ignition' href="+this.src+" />");
		//}
	}else{
		$(this).parent("a").attr("rel","ignition");
	};
});
})