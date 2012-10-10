var urlRewrite = {
	mark:{
		host:/(.*?)/,
		post:/([a-z,\d]+)/,
		category:/([a-z,\d]+)(_([0-9]+))?/,
		alias:/([a-z,\d]+)(_([0-9]+))?/,
		user:/(.+?)(_([0-9]+))?/,
		year:/(\d{4})/,
		month:/(\d{2})/,
		day:/(\d{2})/,
		id:/(\d+)(_([0-9]+))?/,
		date:/\d{4}-(\d{2})-(\d{2})/
	},
	active:{
		ZC_ARTICLE_REGEX:'$1/zb_system/view\.asp\?id=$2',
		ZC_PAGE_REGEX:'$1/zb_system/view\.asp\?id=$2',
		ZC_CATEGORY_REGEX:'$1/catalog\.asp\?cate=$2&page=$4',
		ZC_USER_REGEX:'$1/catalog\.asp\?auth=$2&page=$4',
		ZC_TAGS_REGEX:'$1/catalog\.asp\?tags=$2&page=$4',
		ZC_DATE_REGEX:'$1/catalog\.asp\?date=$2&page=$4',
		ZC_DEFAULT_REGEX:'$1/catalog.asp?page=$2'
	},
	get:function(url,zc,type){
		var z = {};
		for(var key in this.mark){
			url = url.replace('{%'+key+'%}',
			this.mark[key].toString().replace(/^\/(.*?)\/$/g,'$1'));
		}
		url = url.replace(/(\/default\.html)$/ig,'(/)?');
		if(url=='(.*?)(/)?'){
			url = url.replace('(/)?','default_(\d+).html');	
		}
		url = '^'+url+'$';
		url = url.replace(/\.([a-z]+)/ig,'\\.$1');
		z.rule = url;
		z.action = eval('this.active.'+zc.replace('edt',''));
		switch(type){
			case 'ISAPI2':
			break;
			case 'ISAPI3':
			break;
			case 'IIS7':
				z.action = z.action.replace(/\$(\d+)/g,'{R:$1}');
			break;
		}
		return z;
	},
	display:function(){
		var s = '';zl = [];zc = {};
		var type = $('#type').val();
		$('select,input').each(function(){
			var url = $(this).val();
			var zc = $(this).attr('name');
			if(/ZC_[A-Z]+_REGEX/.test(zc)){
				zl.push(urlRewrite.get(url,zc,type));
			}	
		});	
		$.ajax({
			url: 'iis7.html',
			type: "GET",
			data: {t:Math.random()},
			dataType: "html",
			success: function(html) {
				switch(type){
					case 'ISAPI2':
						html = '[ISAPI_Rewrite]\n';
						for(var i=0;i<zl.length;i++){
							html+='RewriteRule '+zl[i].rule+' '+zl[i].action+'\n';
						}
					break;
					case 'ISAPI3':
						html = '';
						for(var i=0;i<zl.length;i++){
							html+='RewriteRule '+zl[i].rule+' '+zl[i].action+'\n';
						}
					break;
					case 'IIS7':
						var div=$('<div></div>').html(html);
						div.find('rules').html('');
						for(var i=0;i<zl.length;i++){
							var rule = $(html).find('rules').children().eq(0).clone();
								rule.attr('name','Rule '+i);
								rule.find('match').attr('url',zl[i].rule);
								rule.find('action').attr('url',zl[i].action);
								div.find('rules').append(rule);
						}
						html = div.html().replace(/<!--(.*?)-->/,'<$1>');
					break;
				}
				urlRewrite.call(html);
			}
		});
	}
};