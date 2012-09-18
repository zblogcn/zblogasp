	setTimeout(function(){ 
	$("ul.ul-subcates").prev("a").before("<span class='sh'>-</span>");
	$("span.sh").click(function (){
		$(this).next().next("ul").toggle("fast");
	})
	.toggle(
		function () {
		$(this).html("+");
		},
		function () {
		$(this).html("-");
	});
	},500);