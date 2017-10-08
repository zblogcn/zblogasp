	function changeval(a,b){
		if(a==1){
			a="#edtZC_ARTICLE_REGEXT";
			if(b==1){b="{%host%}/{%post%}/{%alias%}.html"};
			if(b==2){b="{%host%}/{%year%}/{%month%}/{%alias%}.html"};
			if(b==3){b="{%host%}/{%category%}/{%alias%}.html"};
			if(b==4){b="{%host%}/{%post%}/{%alias%}/default.html"};
			if(b==5){b="{%host%}/{%category%}/{%id%}/default.html"};
		}
		if(a==2){
			a="#edtZC_PAGE_REGEX";
			if(b==1){b="{%host%}/{%alias%}.html"};
			if(b==2){b="{%host%}/{%alias%}/default.html"};
			if(b==3){b="{%host%}/{%id%}.html"};
			if(b==4){b="{%host%}/{%id%}/default.html"};
		}
		if(a==3){
			a="#edtZC_CATEGORY_REGEX";
			if(b==1){b="{%host%}/category/{%alias%}/default.html"};
			if(b==2){b="{%host%}/category/{%id%}/default.html"};
			if(b==3){b="{%host%}/category-{%alias%}.html"};
			if(b==4){b="{%host%}/category-{%id%}.html"};
		}
		if(a==4){
			a="#edtZC_TAGS_REGEX";
			if(b==1){b="{%host%}/tags/{%name%}/default.html"};
			if(b==2){b="{%host%}/tags/{%alias%}/default.html"};
			if(b==3){b="{%host%}/tags/{%id%}/default.html"};
			if(b==4){b="{%host%}/tags-{%name%}.html"};
			if(b==5){b="{%host%}/tags-{%alias%}.html"};
			if(b==6){b="{%host%}/tags-{%id%}.html"};
		}
		if(a==5){
			a="#edtZC_DATE_REGEX";
			if(b==1){b="{%host%}/date/{%date%}/default.html"};
			if(b==2){b="{%host%}/date-{%date%}.html"};
		}
		if(a==6){
			a="#edtZC_DEFAULT_REGEX";
			if(b==1){b="{%host%}/page/{%page%}/default.html"};
			if(b==2){b="{%host%}/default.html"};
		}
		if(a==7){
			a="#edtZC_USER_REGEX";
			if(b==1){b="{%host%}/author/{%alias%}/default.html"};
			if(b==2){b="{%host%}/author/{%id%}/default.html"};
			if(b==3){b="{%host%}/author-{%alias%}.html"};
			if(b==4){b="{%host%}/author-{%id%}.html"};
		}

		$(a).val(b);
	}

	function flashradio(){
		if($("#edtZC_POST_STATIC_MODE").val()=="STATIC"){
			$("input[name='radio'],input[name='radio2']").removeAttr("disabled");
		};
		if($("#edtZC_POST_STATIC_MODE").val()=="ACTIVE"){
			$("input[name='radio'],input[name='radio2']").attr("disabled","disabled");
		};
		if($("#edtZC_POST_STATIC_MODE").val()=="REWRITE"){
			$("input[name='radio'],input[name='radio2']").removeAttr("disabled");
		};
		if($("#edtZC_STATIC_MODE").val()=="ACTIVE"){
			$("input[name='radio3'],input[name='radio4'],input[name='radio5'],input[name='radio6'],input[name='radio7']").attr("disabled","disabled");
		};
		if($("#edtZC_STATIC_MODE").val()=="REWRITE"){
			$("input[name='radio3'],input[name='radio4'],input[name='radio5'],input[name='radio6'],input[name='radio7']").removeAttr("disabled");
		};
	}

function enable(list){
	$("[_enable]").each(function(index,element){
		var Element=$(element);
		Element.addClass("disable");
		if(list!="none"&&(((Element.attr("_enable")=="{%host%}"||Element.attr("_enable")=="{%post%}")&&(list.indexOf("%")>-1||list==""))||list=="all"||list.indexOf(Element.attr("_enable"))>-1)){Element.removeClass("disable")}
		
		/*这句话逻辑翻译如下：
		列表不是none的同时{
			如果li是{%host%}或{%post%}，判断{
				如果它是空的，就允许显示。
				如果它不是空的，但不含%符号，就不允许显示。
			}
			如果列表是all{
				一定让它显示
			}
			如果在列表里找到了当前ID{
				让它显示
			}
		}
		
		*/
	})
}