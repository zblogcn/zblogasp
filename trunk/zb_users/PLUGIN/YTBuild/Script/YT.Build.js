var YT={
	Thread:[],
	MaxThread:10,
	Delay:300,
	loadIMG:'/zb_system/IMAGE/ADMIN/loading.gif',
	Insert:function(obj,str){  
		if(str == '') return;
		var _length = obj.value.length;  
			obj.focus();  
		if(typeof document.selection!='undefined'){  
			document.selection.createRange().text = str;  
		}else {
			obj.value = obj.value.substr(0,obj.selectionStart)+str+obj.value.substring(obj.selectionEnd,_length);  
		}  
	},
	Show:function(){
		$("#Panel").css('opacity', '0.8').css("height",document.documentElement.clientHeight).animate({opacity:'show'});
	},
	Catalog:function(e,aC,Key){
		$("#"+e).click(function(){
			YT.Thread[0] = 0;
			$.ajax({
				url: 'YT.Ajax.asp',
				type: 'POST',
				dataType: 'json',
				data: {Act: "Catalog" ,Key: Key ,aC: YT.GSI(aC).join(",") , t:Math.random()},
				error: function(e){YT.SettingsPanel("系统错误","<font color=red>"+e.responseText+"</font>");},
				beforeSend: function(){YT.SettingsPanel("系统提示","正在获取数据集合,请稍候...",true);},
				success: function(d) {
					YT.SettingsPanel("系统提示","正在启动线程发布,请稍候...",true);
					if(d && d.length > 0){
						YT.SettingsPanel($("#"+e).val(),null);
						//获取需要创建的线程数,初始化
						for(var i = 0; i < d.length; i++){
							YT.Thread[i] = 0;
							//绘制线程进度条
							$("#content").append($("#Template").attr("class","T"+i).clone()).find("#Template").attr("id","Te").show();
							$('#content').find('.T'+i).find('#UserShow').attr("title",$("#"+aC+">option[value="+d[i].Type+"]").get(0).text);	
							//启动线程
							YT.ThreadCatalog(d[i],i);
						}
					}else{
						YT.Thread = [];
						YT.SettingsPanel(null,"<font color=red>数据获取失败</font>");
					}
				}
			});							   								   
		});
	},
	View:function(e,aC,Key){
		$("#"+e).click(function(){
			YT.Thread[0] = 0;
			$.ajax({
				url: 'YT.Ajax.asp',
				type: 'POST',
				dataType: 'json',
				data: {Act: "View" ,Key: Key ,aC: YT.GSI(aC).join(",") , t:Math.random()},
				error: function(e){YT.SettingsPanel("系统错误","<font color=red>"+e.responseText+"</font>");},
				beforeSend: function(){YT.SettingsPanel("系统提示","正在获取数据集合,请稍候...",true);},
				success: function(d) {
					if(d && d.length > 0){
						YT.SettingsPanel("设置线程",'<input type="text" value="'+YT.MaxThread+'" /> <input type="button" value="开始任务" />');
						$("#content").find("input[type=button]").click(function(){
							var t = $(this).parent().find("input[type=text]");
							if(/^(?!0)\d+$/.test(t.val())){
								YT.SettingsPanel("系统提示","正在启动线程发布,请稍候...",true);
								var tl = parseInt(t.val());
								//检查线程数量,防止客户端假死
								if(tl > YT.MaxThread){
									tl = YT.MaxThread;	
								}
								//根据线程分割数组
								var dt = YT.SegmentationArray(d,tl);
								//检查线程数是否合法
								if(tl > dt.length){
									tl = dt.length;	
								}
								YT.SettingsPanel($("#"+e).val(),null);
								//开启线程数
								for(var i = 0; i < tl; i++){
									YT.Thread[i] = 0;
									//绘制线程进度条
									$("#content").append($("#Template").attr("class","T"+i).clone()).find("#Template").attr("id","Te").show();
									
									//启动线程
									YT.ThreadView(dt[i],i);
								}
							}else{
								t.attr("style","border:red 1px solid");	
							}												
						});
/*						
						YT.SettingsPanel("系统提示","正在启动线程发布,请稍候...",true);
						var tl = YT.MaxThread;
						//根据线程分割数组
						var dt = YT.SegmentationArray(d,tl);
						//检查线程数是否合法
						if(tl > dt.length){
							tl = dt.length;	
						}
						YT.SettingsPanel($("#"+e).val(),null);
						//开启线程数
						for(var i = 0; i < tl; i++){
							YT.Thread[i] = 0;
							//绘制线程进度条
							$("#content").append($("#Template").attr("class","T"+i).clone()).find("#Template").attr("id","Te").show();
							
							//启动线程
							YT.ThreadView(dt[i],i);
						}*/
					}else{
						YT.Thread = [];
						YT.SettingsPanel(null,"<font color=red>数据获取失败</font>");
					}
				}
			});							   								   
		});
	},
	SettingsPanel:function(title,content,b){
		YT.Show();
		var sload='<img src="'+YT.loadIMG+'" /> ';
		if(title != null) $("#title").html(title);
		if(content != null) $("#content").html((b?sload:'')+content);
	},
	SegmentationArray:function(d,tl){
		var t = Math.ceil(d.length / tl);
		var ds = new Array();
		for(var i = 0; i < tl; i++){
			var dt = new Array();
			var u = 0;
			for(var j = (i * t); j < d.length; j++){
				if(u < t){
					dt.push(d[j]);
					u++;
				}
			}
			if(dt.length > 0){
				ds.push(dt);	
			}
		}
		return ds;
	},
	Default:function(){
		//注册线程,以防止用户强行关闭层
		YT.Thread[0] = 0;
		$("#default").click(function(){
			$.ajax({
				url: 'YT.Ajax.asp',
				type: 'POST',
				dataType: 'json',
				data: { Act: "Default" , t:Math.random() },
				error: function(e) { YT.Thread[0] = -1;YT.SettingsPanel("系统错误","<font color=red>"+e.responseText+"</font>");},
				beforeSend: function(){YT.SettingsPanel("发布首页","正在发布首页,请稍候...",true);},
				success: function(d) {
					YT.SettingsPanel(null,d?"发布成功":"<font color=red>发布失败</font>");	
					YT.Thread[0] = -1;
				}
			});	
		});	
	},
	ThreadCatalog:function(d,i) {
		var $p = Math.round(parseFloat(YT.Thread[i]) / parseFloat(d.intPageCount) * 10000) / 100.00;
		var $w = $p;
		if (!isNaN($p)) {
			$w += "%";
			$("#content .T"+i+" #Status font").text($w);
			$("#content .T"+i+" #Status div").css("width", $w);
		}
		if (YT.Thread[i] < d.intPageCount) {
			$.ajax({
				url: "YT.Ajax.asp",
				type: "POST",
				global: false,
				data: {Act:"ThreadCatalog",Key:d.Key,Page:(YT.Thread[i]+1),ID:d.ID},
				dataType: "json",
				error: function(e) {YT.SettingsPanel("系统错误","<font color=red>"+e.responseText+"</font>");},
				success: function(o) {
					if (typeof(o) == "boolean" && Boolean(o)) {
						setTimeout(function(){
							YT.Thread[i]++;
							$('#content').find('.T'+i).find('#UserShow').text("共"+d.intPageCount+"页 正在创建["+$('#content').find('.T'+i).find('#UserShow').attr("title")+"]第"+YT.Thread[i]+"页");
							YT.ThreadCatalog(d,i);
						},YT.Delay);
					}
				}
			});
		}else{
			YT.Thread[i] = -1;
			//释放线程
			if(YT.CompleteThread()){
				YT.SettingsPanel(null,"发布成功");
			}
		}
	},
	ThreadView:function(d,i) {
		var $p = Math.round(parseFloat(YT.Thread[i]) / parseFloat(d.length) * 10000) / 100.00;
		var $w = $p;
		if (!isNaN($p)) {
			$w += "%";
			$("#content .T"+i+" #Status font").text($w);
			$("#content .T"+i+" #Status div").css("width", $w);
		}
		if (YT.Thread[i] < d.length) {
			$.ajax({
				url: "YT.Ajax.asp",
				type: "POST",
				global: false,
				data: {Act:"ThreadView",ID:d[YT.Thread[i]],t:Math.random()},
				dataType: "json",
				error: function(e) {YT.SettingsPanel("系统错误","<font color=red>"+e.responseText+"</font>"); },
				success: function(o) {
					if (typeof(o) == "boolean" && Boolean(o)) {
						setTimeout(function(){
							YT.Thread[i]++;
							$('#content').find('.T'+i).find('#UserShow').text("共"+d.length+"条数据 线程["+(i+1)+"]正在创建第"+YT.Thread[i]+"条数据");
							YT.ThreadView(d,i);
						},YT.Delay);
					}
				}
			});
		}else{
			YT.Thread[i] = -1;
			//释放线程
			if(YT.CompleteThread()){
				YT.SettingsPanel(null,"发布成功");
			}
		}
	},
	Close:function(){
		//检查线程是否工作
		if(YT.CompleteThread()){
			$("#Panel").animate({opacity:'hide'});	
		}
	},
	CompleteThread:function(){
		var b=true;
		for(var j = 0; j < YT.Thread.length; j++){
			if(YT.Thread[j] > -1){
				b=false;
				break;
			}
		}
		return b;
	},
	GSI:function(a){
		var ary = new Array();
		if($("#"+a).val() != null){
			ary = $("#"+a).val();
		}else{
			$("#"+a+" option").each(function(){
				ary.push($(this).val());
			});
		}
		return ary;
	},
	Config:function(s){
		return s=='REWRITE'?YT.ZC_STATIC_MODE.REWRITE:YT.ZC_STATIC_MODE.ACTIVE;
	},
	ZC_STATIC_MODE:{
		REWRITE:{
			ZC_ARTICLE_REGEX:'{%host%}/{%post%}/{%alias%}.html',
			ZC_PAGE_REGEX:'{%host%}/{%alias%}.html',
			ZC_CATEGORY_REGEX:'{%host%}/category-{%id%}.html',
			ZC_USER_REGEX:'{%host%}/author-{%id%}.html',
			ZC_TAGS_REGEX:'{%host%}/tags-{%alias%}.html',
			ZC_DATE_REGEX:'{%host%}/{%year%}-{%month%}.html',
			ZC_DEFAULT_REGEX:'{%host%}/default.html'
		},
		ACTIVE:{
			ZC_ARTICLE_REGEX:'{%host%}/{%post%}/{%alias%}.html',
			ZC_PAGE_REGEX:'{%host%}/{%alias%}.html',
			ZC_CATEGORY_REGEX:'{%host%}/catalog.asp?cate={%id%}',
			ZC_USER_REGEX:'{%host%}/catalog.asp?user={%id%}',
			ZC_TAGS_REGEX:'{%host%}/catalog.asp?tags={%alias%}',
			ZC_DATE_REGEX:'{%host%}/catalog.asp?date={%year%}-{%month%}',
			ZC_DEFAULT_REGEX:'{%host%}/catalog.asp'
		}
	}
}