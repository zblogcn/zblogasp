///////////////////////////////////////////////////////////////////////////////
// 作	 者:    	瑜廷
// 技术支持:     33195@qq.com
// 程序名称:    	YT.CMS Script
// 开始时间:    	2011-05-28
// 最后修改:    	2012-11-05
// 备	 注:    	only for YT.CMS
///////////////////////////////////////////////////////////////////////////////
var YT = {
	CMS:[],
	ESC:null,
	Panel:{
		TPL:{
			C:function(){
				(function(){
					var but = $('#tpl').find('input[type="button"]');
						but.first().click(function(){
							var ir = $('#tpl').find('li.d_li');
							var r = ir.last().clone();
								r.find('h4').text('').hide();
							var nav = r.find('.d_adviewcon').show();
								nav2 = nav.clone();
								nav.before(nav2);
								$.ajax({
									url: YT_CMS_XML_URL+YTConfig.Block,
									type: 'GET',
									dataType: 'xml',
									data: { t:Math.random() },
									success: function(xml) {
										var _b = document.createElement('select');
											_b.options.add(new Option('「模块」','-1'));
										$('Block', xml).each(function(i) {
											var Block = $('Block', xml).get(i);
											_b.options.add(new Option($('Name',Block).text(),'{cache:'+$('Name',Block).text().toUpperCase()+'}'));
										});
										_b.onchange = function(){
											YT.InsertText(nav.parent().find('textarea')[0],this.value,true);
											_b.selectedIndex = 0;
										}
										nav.append(_b);	
										YT.S(nav[0]);
									}
								});
								var txt = document.createElement('input');
								$(txt).css('width','95%');
								nav2.append(txt);
							r.find('span.d_check input[type="checkbox"]').each(function(j){
								$(this).click(function(){
									r.find('span.d_check input[type="checkbox"]').attr('checked',false);
									$(this).attr('checked',true);
									if(j==1){
										r.fadeOut("slow");	
									}
								});
							}).eq(0).trigger('click');
							r.addClass('tpl').insertBefore(ir.first().next()).fadeIn("slow");
						});
						but.last().click(function(){
							var ary = [];
							$('#tpl').find('li.tpl').each(function(){
								var i2,data = {};
								$(this).find('span.d_check input[type="checkbox"]').each(function(i){
									if($(this).attr('checked')){
										i2 = i;	
									}
								});
								var name = $(this).find('h4');
								var si = name.text();
								if(i2==0){
									data.content = $(this).find('textarea').val();
									if(si==''){
										data.name = $(this).find('.d_adviewcon').eq(0).find('input').val()+'.html';
										data.action = 'SaveFile';	
									}else{
										data.name = $(this).find('.d_check label').last().text();
										data.action = 'UpdateFile';
									}
								}else if(i2==1){
									data.name = name.text();
									data.action = 'DelFile';		
								}
								if(data.action&&data.name!=''){
									ary.push(data);
								}
							});
							var j = 0;
							if(ary.length>0){send(ary);}
							function send(data){
								if(j<data.length){
									$.ajax({
										url: 'YT.Ajax.asp',
										type: 'POST',
										dataType: 'json',
										data: { Action: data[j].action!='DelFile'?'SaveFile':data[j].action ,
										Name: data[j].name, Content: data[j].content , t:Math.random() },
										success: function(b) {
											if(b){
												var div = $('#tpl').find('div').last();
													div.find('span').remove();
												var span = document.createElement('span');
												$(span).addClass('loading').text(data[j].action.toUpperCase()+'-'+data[j].name+'√');
												div.append(span);
												setTimeout(function(){j++;send(data);},500);
											}
										}
									});	
								}else{
									setTimeout(function(){
										$('#tpl').find('div').last().find('span').remove();
										YT.Panel.TPL.M();		
									},1000);	
								}
							}
						});	
				})();
			},
			M:function(){
				(function(){
					lang = [
						{tpl:'default.html',text:'首页主模板文件'},
						{tpl:'catalog.html',text:'列表页模板文件'},
						{tpl:'b_article-multi.html',text:'摘要文章模板'},
						{tpl:'b_article-istop.html',text:'置顶文章模板'},
						{tpl:'b_pagebar.html',text:'分页条模板'},
						{tpl:'page.html',text:'独立页面模板'},
						{tpl:'b_function.html',text:'侧栏模块模板'},
						{tpl:'b_article_comment_pagebar.html',text:'评论分页条模板'},
						{tpl:'single.html',text:'日志页主模板文件'},
						{tpl:'b_article-single.html',text:'日志页文章模板'},
						{tpl:'b_article_trackback.html',text:'引用通告显示模板'},
						{tpl:'b_article_mutuality.html',text:'相关文章显示模板'},
						{tpl:'b_article_comment.html',text:'评论内容显示模板'},
						{tpl:'b_article_commentpost.html',text:'评论发送表单模板'},
						{tpl:'b_article_commentpost-verify.html',text:'评论验证码显示模板'},
						{tpl:'b_article_tag.html',text:'每个tag 的显示样式'}
					];
					$('#tpl').find('div').last().append('<span class="loading">loading...</span>');
					$('.tpl').remove();
					$.ajax({
						url: 'YT.Ajax.asp',
						type: 'POST',
						dataType: 'json',
						data: { Action: 'tplList',t:Math.random() },
						success: function(ary) {
							var i = 0;
							function stpl(data){
								if(i<data.length){
									var ir = $('#tpl').find('li').last();
									var r = ir.clone();
										for(var j=0;j<lang.length;j++){
											if(lang[j].tpl==ary[i].toLowerCase()){
												r.find('h4').text(lang[j].text);
												break;
											}	
										}
									if(r.find('h4').text()==''){r.find('h4').text(ary[i]);}
									var nav = r.find('.d_adviewcon').hide();
									var label = document.createElement('label');
									$(label).text(ary[i]).click(function(){
											if(nav.html().search(/<.*?>/)!=-1){
												if(nav.css('display')=='none'){
													nav.fadeIn("slow");
													r.find('textarea').fadeIn("slow");	
												}else{
													nav.fadeOut("slow");
													r.find('textarea').fadeOut("slow");	
												}
											}
									});
									r.find('.d_check').append(label);
									r.find('textarea').hide();
									r.find('span.d_check input[type="checkbox"]').attr('checked',false).each(function(j){
										$(this).click(function(){
											r.find('span.d_check input[type="checkbox"]').attr('checked',false);
											$(this).attr('checked',true);
											if(j==0){
												if(nav.html()==''){
													$.ajax({
														url: YT_CMS_XML_URL+YTConfig.Block,
														type: 'GET',
														dataType: 'xml',
														data: { t:Math.random() },
														success: function(xml) {
															var _b = document.createElement('select');
																_b.options.add(new Option('「模块」','-1'));
															$('Block', xml).each(function(i) {
																var Block = $('Block', xml).get(i);
																_b.options.add(new Option($('Name',Block).text(),'{cache:'+$('Name',Block).text().toUpperCase()+'}'));
															});
															_b.onchange = function(){
																YT.InsertText(nav.parent().find('textarea')[0],this.value,true);
																_b.selectedIndex = 0;
															}
															nav.append(_b).fadeIn("slow");	
															YT.S(nav[0]);
														}
													});
												}else{
													nav.fadeIn("slow");
												}
												$.get(
													'YT.Ajax.asp',{ 
														Action: 'GetFile',
														File:r.find('.d_check label').last().html(),
														t:Math.random() 
													},
													function(txt){
														var area = r.find('textarea').css('height','auto').val(txt).fadeIn("slow");
														var height = $(area)[0].scrollHeight;
														area.css('height',(height+20)>600?600:(height+20)+'px');
													}
												);
											}else{
												nav.fadeOut("slow");
												r.find('textarea').fadeOut("slow");
											}
										});
									});
									r.addClass('tpl').insertBefore(ir).fadeIn("slow");
									(function(){i++;stpl(ary);})();
								}
							}
							if(ary.length>0){stpl(ary);}
							$('#tpl').find('div').last().find('span').remove();
						}
					});
				})();	
			}	
		},
		Model:{
			C:function(){
				(function(){
					var but = $('#model').find('input[type="button"]');
					but.first().click(function(){
						var ir = $('#model').find('li.d_li');
						var r = ir.last().clone();
							r.find('h4').text('').hide();
						var r2 = r.find('div.d_adviewcon');
						var rl = r2.find('li');
						r.find('span.d_check input[type="checkbox"]').each(function(j){
							$(this).click(function(){
								r.find('span.d_check input[type="checkbox"]').attr('checked',false);
								$(this).attr('checked',true);
								if(j==0){
									var s = $(this).parent().html();
									s = s.replace('修改','保存');
									$(this).parent().html(s);
									r2.fadeIn("slow");
								}else if(j==1){
									r.fadeOut("slow");	
								}
							});
						}).eq(0).trigger('click');
						rl.eq(0).find('input[type="text"]').val('');
						rl.first().next().hide();
						rl.eq(0).find('input').eq(2).unbind('click').click(function(){
							var irl = rl.first().next();
							var rl2 = irl.clone();
							rl2.find('input[type="text"]').val('');
							rl2.find('select').val(0);
							rl2.addClass('fields').hover(function(){
								$(this).find('em').fadeIn("slow")
								.css('cursor','pointer').unbind('click').click(function(){
									$(this).parent().remove();	
								});
							},function(){
								$(this).find('em').fadeOut("slow");
							}).insertBefore(irl).show();
						}).trigger('click');
						r.find('div.d_status').hide();
						r.addClass('model').insertBefore(ir.first().next()).fadeIn("slow");
					});
					but.last().click(function(){
						var ary = [];
						var h4 = 0;
						$('#model').find('li.model').each(function(){
							var i2,data = {};
							$(this).find('span.d_check input[type="checkbox"]').each(function(i){
								if($(this).attr('checked')){
									i2 = i;	
								}
							});
							data.json = {};
							data.index = h4;
							data.json.table = {};
							data.json.fields = [];
							var name = $(this).find('h4');
							var si = name.text();
							if(si!=''){h4++;};
							if(i2==0){
								var row = $(this).find('li');
								var tables = row.eq(0).find('input');
								data.json.table.name = tables.eq(0).val();
								var ab = [];
								row.last().find('select option').each(function(){
									if($(this).attr('selected')){
										ab.push($(this).val());	
									}	
								});
								data.json.table.bind = ab.join(',');
								data.json.table.description = tables.eq(1).val();
								row.each(function(){
									if($(this).attr('class')=='fields'){
										data.json.fields.push({
											name:$(this).find('input').eq(0).val(),
											description:$(this).find('input').eq(1).val(),
											value:$(this).find('input').eq(2).val(),
											property:$(this).find('select').eq(0).val(),
											type:$(this).find('select').eq(1).val()
										});
									}	
								});
								if(si==''){
									data.index = '';
									data.action = 'SaveModel';	
								}else{
									data.action = 'UpdateModel';
								}
							}else if(i2==1){
								data.json.table.name = name.text();
								data.action = 'DelModel';		
							}
							if(data.action&&data.json.table.name!=''){
								if(data.action=='DelModel'){
									ary.push(data);
								}else{
									if(data.json.fields.length>0){
										ary.push(data);	
									}
								}
							}
						});
						Array.prototype.seq = function(s){
							var ary = [];
							for(var j=0;j<s.length;j++){
								for(var i=0;i<this.length;i++){
									if(this[i].action==s[j]){
										ary.push(this[i]);
									}
								}
							}
							return ary;
						}
						var j = 0,index = 0;
						data = ary.seq(['UpdateModel','DelModel','SaveModel']);
						if(data.length>0){send(data);}
						function send(data){
							if(j<data.length){
								data[j].index = data[j].index - index;
								$.ajax({
									url: YT_CMS_XML_URL+YTConfig.Model,
									type: 'GET',
									dataType: 'xml',
									data: { t:Math.random() },
									success: function(xml) {
										var b=false;
										$('Model', xml).each(function(i) {
											var Model = $('Model', xml).get(i);
											if(data[j].json.table.name==$('Table>Name',Model).text()){
												b=true;
												return;
											}
										});
										if(b&&data[j].action=='SaveModel'){
											var div = $('#model').find('div').last();
												div.find('span').remove();
											var span = document.createElement('span');
											$(span).addClass('loading').css('color','red').text(data[j].action.toUpperCase()+'-'+data[j].json.table.name+'×');
											div.append(span);
											setTimeout(function(){j++;send(data);},1000);
										}else{
											$.ajax({
												url: 'YT.Ajax.asp',
												type: 'POST',
												dataType: 'json',
												data: { Action: data[j].action ,Index:data[j].index,Json:$.toJSONString(data[j].json), t:Math.random() },
												success: function(b) {
													if(b){
														if(data[j].action.toLowerCase()=='delmodel'){
															index++;	
														}
														var div = $('#model').find('div').last();
															div.find('span').remove();
														var span = document.createElement('span');
														$(span).addClass('loading').text(data[j].action.toUpperCase()+'-'+data[j].json.table.name+'√');
														div.append(span);
														setTimeout(function(){j++;send(data);},500);
													}
												}
											});
										}
									}
								});	
							}else{
								setTimeout(function(){
									$('#model').find('div').last().find('span').remove();
									YT.Panel.Model.M();		
								},1000);
							}
						}
					});
				})();
			},
			M:function(){
				(function(){
					$('#model').find('div').last().append('<span class="loading">loading...</span>');
					$('.model').remove();
					$.ajax({
						url: YT_CMS_XML_URL+YTConfig.Model,
						type: 'GET',
						dataType: 'xml',
						data: { t:Math.random() },
						success: function(xml) {
							$('Model', xml).each(function(i) {
								var Model = $('Model', xml).get(i);
								var ir = $('#model').find('li.d_li').last();
								var r = ir.clone();
								var table = $('Table>Name',Model).text();
								var desc = $('Table>Description',Model).text();
								var bind = $('Table>Bind',Model).text();
								r.find('h4').text(desc);
								r.find('.fields').remove();
								var r2 = r.find('div.d_adviewcon');
								var rl = r2.find('li');
								rl.eq(0).find('input').eq(0).val(table);
								rl.eq(0).find('input').eq(1).val(desc);
								rl.first().next().hide();
								rl.eq(0).find('input').eq(2).unbind('click').click(function(){
									var irl = rl.first().next();
									var rl2 = irl.clone();
									rl2.find('input[type="text"]').val('');
									rl2.find('select').val(0);
									rl2.addClass('fields').hover(function(){
										$(this).find('em').fadeIn("slow")
										.css('cursor','pointer').unbind('click').click(function(){
											$(this).parent().remove();	
										});
									},function(){
										$(this).find('em').fadeOut("slow");
									}).insertBefore(irl);
								});
								rl.last().find('select').attr('selected',false);	
								var b2 = bind.split(',');
								rl.last().find('input[type="checkbox"]').attr('selected',false);
								for(var l=0;l<b2.length;l++){
									rl.last().find('select option').each(function(){
										if(b2[l]==$(this).val()){
											$(this).attr('selected',true);	
										}
									});	
								}
								$('Field',Model).each(function(){
									var irl = rl.first().next();
									var rl2 = irl.clone().show();
									rl2.find('input').eq(0).val($(this).find('Name').text());
									rl2.find('input').eq(1).val($(this).find('Description').text());
									rl2.find('input').eq(2).val($(this).find('Value').text());
									rl2.find('select').eq(0).val($(this).find('Property').text());
									rl2.find('select').eq(1).val($(this).find('Type').text());
									rl2.addClass('fields').hover(function(){
										$(this).find('em').fadeIn("slow")
										.css('cursor','pointer').unbind('click').click(function(){
											$(this).parent().remove();	
										});
									},function(){
										$(this).find('em').fadeOut("slow");
									}).insertBefore(irl);
								});	
								r.find('span.d_check input[type="checkbox"]').each(function(j){
									$(this).click(function(){
										r.find('span.d_check input[type="checkbox"]').attr('checked',false);
										$(this).attr('checked',true);
										if(j==0){
											r2.fadeIn("slow");
										}else if(j==1){
											r2.hide();	
										}
									});
								});
								$.ajax({
									url: 'YT.Ajax.asp',
									type: 'POST',
									dataType: 'json',
									data: { Action:'Exist',Name:table, t:Math.random() },
									success: function(bool) {
										var status = function(s,b){
											var s2 = [];
											s2.push('-');
											s2.push(s==''?'系统表':'用户表');
											s2.push('-');
											s2.push(b?'已安装[√]<font color="red">→点击此处卸载此表,注意执行此操作会删除此表</font>':
											'<font color="#999999">未安装[×]→点击此处安装此表</font>')
											return s2.join('');
										};
										r.find('div.d_status').css('cursor','pointer').last()
										.html(status(bind,bool)).attr('rel',bool).unbind('click').click(function(){
											var ib = $(this).attr('rel');
											//变更模型状态
											$.ajax({
												url: 'YT.Ajax.asp',
												type: 'POST',
												dataType: 'html',
												data: { Action:ib=='true'?'UnInstall':'Install',Index:i,t:Math.random() },
												success: function(ib2) {
													r.find('div.d_status').last().html(
													status(bind,(ib2=='install'?true:false))
													).attr('rel',(ib2=='install'?'true':'false'));
												}
											});	
										});
									}
								});	
								r.addClass('model').insertBefore(ir).fadeIn("slow");
							});
							$('#model').find('div').last().find('span').remove();
						}
					});
				})();
			}
		},
		Analysis:function(){
			$('#cmbCate').change(function(){
				var ele,val;
				var _Cate = $(this).val();
				$('#model').find('p').remove();
				$.ajax({
					url:YT_CMS_XML_URL+YTConfig.Model,
					type: 'GET',
					dataType: 'xml',
					data: { t:Math.random() },
					success: function(xml) {
						$('Model', xml).each(function(i) {
							var Model = $('Model', xml).get(i);
							var _Bind = $('Table>Bind',Model).text().split(',');
							var _isBind = false;
							for(var j=0;j<_Bind.length;j++){
								if(parseInt(_Cate) == parseInt(_Bind[j])){
									_isBind = true;
									break;
								}
							}
							if(_isBind){
								$('Field',Model).each(function(){
									switch($(this).find('Type').text()){
										case 'text':
											ele = document.createElement('input');
											ele.type = $(this).find('Type').text();
											ele.value = $(this).find('Value').text();
											ele.name = $(this).find('Name').text();
											ele.style.width = '50%';
											$('<p>'+$(this).find('Description').text()+':</p>')
											.attr('title',$(this).find('Type').text()).append(ele).appendTo($('#model'));
										break;
										case 'select':
											ele = document.createElement('select');
											ele.name = $(this).find('Name').text();
											try{
												val = eval($(this).find('Value').text());
												for(var j=0;j<val.length;j++){
													ele.options.add(new Option(val[j].t,val[j].v));	
												}
											}catch(e){
												val = $(this).find('Value').text().split(',');	
												for(var j=0;j<val.length;j++){
													ele.options.add(new Option(val[j],val[j]));	
												}
											}
											$('<p>'+$(this).find('Description').text()+':</p>')
											.attr('title',$(this).find('Type').text()).append(ele).appendTo($('#model'));
										break;
										case 'checkbox':
											var row = $('<p>'+$(this).find('Description').text()+':</p>')
											.attr('title',$(this).find('Type').text());
											try{
												var val = eval($(this).find('Value').text());
												for(var j=0;j<val.length;j++){
														ele = document.createElement('input');
														ele.type = $(this).find('Type').text();
														ele.value = val[j].v;
														ele.name = $(this).find('Name').text();
														row.append(val[j].t).append(ele);
												}
											}catch(e){
												val = $(this).find('Value').text().split(',');	
												for(var j=0;j<val.length;j++){
														ele = document.createElement('input');
														ele.type = $(this).find('Type').text();
														ele.value = val[j];
														ele.name = $(this).find('Name').text();
														row.append(val[j]).append(ele);
												}	
											}
											row.appendTo($('#model'));
										break;
										case 'textarea':
											ele = document.createElement('textarea');
											ele.value = $(this).find('Value').text();
											ele.name = $(this).find('Name').text();
											ele.style.width = '50%';
											$('<p>'+$(this).find('Description').text()+':</p>')
											.attr('title',$(this).find('Type').text()).append(ele).appendTo($('#model'));
										break;
										default:
											var ueconfig = window.UEDITOR_CONFIG || {};
											ele = document.createElement('input');
											$(ele).attr('type','text').attr('name',$(this).find('Name').text())
											.css('width','50%').val($(this).find('Value').text())
											.addClass($(this).find('Type').text()).click(function(){
												this.callbacks = function(obj,win){
													this.value = '';
													for(key in obj) {
														this.value += ueconfig.imagePath.replace(ZC_BLOG_HOST,'') + obj[key].url;
													}
													win.close();
												};
												var file = this.className.replace('upload-','');
												window.showModalDialog(ZC_BLOG_HOST+'zb_users/plugin/ytcms/'+
												file+'.html',this,file=='image'?'dialogWidth:625px;dialogHeight:340px;resizable:no;scroll:no;status:no;'
												:'dialogWidth:480px;dialogHeight:360px;resizable:no;scroll:no;status:no;');
											});
											$('<p>'+$(this).find('Description').text()+':</p>')
											.attr('title',$(this).find('Type').text()).append(ele).appendTo($('#model'));
										break;
									}
								});
								if($('#edtID').val()!=0){
									$.ajax({
										url: ZC_BLOG_HOST+'ZB_USERS/PLUGIN/YTCMS/YT.Ajax.asp',
										type: 'POST',
										dataType: 'json',
										data: { Action:'GetData', Name:$('Table>Name',Model).text(), ID:$('#edtID').val(), t:Math.random() },
										success: function(r) {
											if(r!=null){
												$('#model').find('p').each(function(j){
													switch($(this).attr('title')){
														case 'text':
															val=k(r,$(this).find('input')[0].name);
															if(val!=null){$(this).find('input').val(unescape(val));}
														break;
														case 'select':
															val=k(r,$(this).find('select')[0].name);
															if(val!=null){$(this).find('select').val(unescape(val));}
														break;
														case 'checkbox':
															val=k(r,$(this).find('input[type="checkbox"]')[0].name);
															if(val!=null){
																val = unescape(val).split(',');
																$(this).find('input[type="checkbox"]').each(function(){
																	for(var l=0;l<val.length;l++){
																		if($(this).val().toLowerCase().replace(/\s+/ig,'')
																		== val[l].toLowerCase().replace(/\s+/ig,'')){
																			$(this).attr('checked',true);
																			break;
																		}	
																	}
																});
															}
														break;
														case 'textarea':
															var val=k(r,$(this).find('textarea')[0].name);
															if(val!=null){$(this).find('textarea').val(unescape(val));}
														break;
														default:
															val=k(r,$(this).find('input')[0].name);
															if(val!=null){$(this).find('input').val(unescape(val));}
														break;
													}		  
												});
											}
										}
									});	
								}
								return false;
							}
						});	
					}
				});
				function k(j,key){
					for(var i=0;i<j.YTARRAY.length;i++){
						if(j.YTARRAY[i]==key){return eval('j.'+key);}	
					}
					return null;
				}
			});
			try{$('#cmbCate').trigger('change');}catch(e){}
/*			(function(){  
				var inp = null;
				var b = setInterval(function(){
						$('p[title="text"] input,textarea').each(function(){
							$(this).focus(function(){
								inp = this;	
							});
							window.clearInterval(b);
						});   		   
					},500);
				var e = $('#divEditTitle');
					e.append('<div class="YTCMS"></div>');
				var obs = [];
				setInterval(function(){
					if(editor.hasContents()){
						$(editor.document).find('img').each(function(){
							var b2=false,ob = new Image();
								ob.src = this.src;
								ob.onload = function(){
									this.id = this.src.replace(/[\:\/\.]/g,'');
									for(var i=0;i<obs.length;i++){
										if(obs[i]==this.id){
											b2 = true;
											break;
										}
									}
									if(this.complete&&!b2){
										obs.push(this.id);
										$(this).css({
											width:'5%',
											height:'5%',
											cursor:'pointer',
											padding:'1px',
											margin:'1px',
											border:'black 1px solid'
										}).click(function(){
											if(inp!=null){
												var s='';
												if(inp.type=='textarea'){
													s=$(inp).val();
													if(s!=''){s+='\n';}
												}
												$(inp).val(s+this.src);
											}
										}).hover(function(){
											$(this).css({border:'red 1px solid'});
										},function(){
											$(this).css({border:'black 1px solid'});
										})
										e.find('div.YTCMS').append($(this).clone(true))
									}
								};
						});
					}					   
				},800);
			})();*/
		},
		SQL:function(){
			$('#sql').find('div').last().append('<span class="loading">loading...</span>');
			$('.sql').remove();
			$.ajax({
				url: 'YT.Ajax.asp',
				type: 'POST',
				dataType: 'json',
				data: { Action: 'ImportList' , t:Math.random() },
				success: function(json) {
					for(var i=0;i<json.length;i++){
						var ir = $('#sql').find('li');
						var r = ir.last().clone();
						var span = document.createElement('span');
						$(span).text(json[i]);
						r.find('label').append(span);
						r.addClass('sql').insertBefore(ir.last()).fadeIn("slow");;
					}
					$('#sql').find('div').last().find('span').hide();
					var j = 0;
					$('#sql').find('input[type="button"]').click(function(){
						var ary = [];
						$('.sql').each(function(){
							if($(this).find('input').attr('checked')){
								var data = {};
								data.action = 'Import';
								data.t = Math.random();
								data.name = $(this).find('.d_check span').text();	
								ary.push(data);
							}	
						});
						j = 0;
						if(ary.length>0){send(ary);}
					});
					function send(data){
						if(j<data.length){
							$.ajax({
								url: 'YT.Ajax.asp',
								type: 'POST',
								dataType: 'json',
								data: data[j],
								success: function(b) {
									var div = $('#sql').find('div').last();
										div.find('span').remove();
									var span = document.createElement('span');
									$(span).addClass('loading').css('color',b?'':'red')
									.text(data[j].action.toUpperCase()+'-'+data[j].name+(b?'√':'×'));
									div.append(span);
									setTimeout(function(){j++;send(data);},500);
								}
							});	
						}else{
							setTimeout(function(){
								$('#sql').find('div').find('span').hide();
							},500);	
						}
					}
				}
			});
		},
		Block:{
			C:function(){
				(function(){
					var but = $('#block').find('input[type="button"]');
					but.first().click(function(){
						var ir = $('#block').find('li');
						var r = ir.last().clone();
						var	txt = document.createElement('input');
							r.find('h4').text('').append($(txt).attr('type','text').addClass('d_inp'));
							var div = document.createElement('div');
							$(div).addClass('d_adviewcon');
							YT.S(div);
							txt = document.createElement('textarea');
							$(txt).addClass('d_tarea');
							r.find('span.d_check input[type="checkbox"]').each(function(i){
								$(this).click(function(){
									r.find('span.d_check input[type="checkbox"]').attr('checked',false);
									$(this).attr('checked',true);
									if(i==1){
										r.remove();	
									}
								});
							}).first().trigger('click');
							r.append(div).append(txt).addClass('block').insertBefore(ir.first().next()).fadeIn("slow");
							
					});
					but.last().click(function(){
						var ary = [];
						var h4 = 0;
						$('#block').find('li.block').each(function(){
							var i2,data = {};
							$(this).find('span.d_check input[type="checkbox"]').each(function(i){
								if($(this).attr('checked')){
									i2 = i;	
								}
							});
							data.json = {};
							data.index = h4;
							var name = $(this).find('h4');
							var si = name.html().search(/<.*?>/);
							if(si==-1){h4++;};
							if(i2==0){
								if(si==-1){
									var content = $(this).find('textarea').val();
									if(typeof(content)=='undefined'){
										data.json.content = $(this).find('input[type="text"]').val();	
									}else{
										data.json.content = content;
									}
									data.json.name = name.text();
									data.action = 'UpdateBlock';
								}else{
									data.index = '';
									data.json.content = $(this).find('textarea').val();
									data.json.name = name.find('input').val();
									data.action = 'SaveBlock';	
								}
							}else if(i2==1){
								data.json.name = name.text();
								data.action = 'DelBlock';		
							}
							if(data.action&&data.json.name!=''){
								ary.push(data);
							}
						});
						Array.prototype.seq = function(s){
							var ary = [];
							for(var j=0;j<s.length;j++){
								for(var i=0;i<this.length;i++){
									if(this[i].action==s[j]){
										ary.push(this[i]);
									}
								}
							}
							return ary;
						}
						//先更新,再删除,后添加,否则节点顺序错误会更新异常
						var j = 0,index = 0;
						data = ary.seq(['UpdateBlock','DelBlock','SaveBlock']);
						if(data.length>0){send(data);}
						function send(data){
							if(j<data.length){
								data[j].index = data[j].index - index;
								$.ajax({
									url: YT_CMS_XML_URL+YTConfig.Block,
									type: 'GET',
									dataType: 'xml',
									data: { t:Math.random() },
									success: function(xml) {
										var b=false;
										$('Block', xml).each(function(i) {
											var Block = $('Block', xml).get(i);
											if(data[j].json.name==$('Name',Block).text()){
												b=true;
												return;
											}
										});
										if(b&&data[j].action=='SaveBlock'){
											var div = $('#block').find('div').last();
												div.find('span').remove();
											var span = document.createElement('span');
											$(span).addClass('loading').css('color','red').text(data[j].action.toUpperCase()+'-'+data[j].json.name+'×');
											div.append(span);
											setTimeout(function(){j++;send(data);},1000);
										}else{
											$.ajax({
												url: 'YT.Ajax.asp',
												type: 'POST',
												dataType: 'json',
												data: { Action: data[j].action ,Index:data[j].index,Json:$.toJSONString(data[j].json), t:Math.random() },
												success: function(b) {
													if(b){
														if(data[j].action.toLowerCase()=='delblock'){
															index++;	
														}
														var div = $('#block').find('div').last();
															div.find('span').remove();
														var span = document.createElement('span');
														$(span).addClass('loading').text(data[j].action.toUpperCase()+'-'+data[j].json.name+'√');
														div.append(span);
														setTimeout(function(){j++;send(data);},500);
													}
												}
											});
										}
									}
								});	
							}else{
								setTimeout(function(){
									$('#block').find('div').last().find('span').remove();
									YT.Panel.Block.M();		
								},1000);
							}
						}
					});
				})();
			},
			M:function(){
				(function(){
					$('#block').find('div').last().append('<span class="loading">loading...</span>');
					$('.block').remove();
					$.ajax({
						url: YT_CMS_XML_URL+YTConfig.Block,
						type: 'GET',
						dataType: 'xml',
						data: { t:Math.random() },
						success: function(xml) {
							$('Block', xml).each(function(i) {
								var Block = $('Block', xml).get(i);
								var ir = $('#block').find('li').last();
								var name = $('Name',Block).text();
								var content = $('Content',Block).text();
								var r = ir.clone();
									r.find('h4').addClass('d_inp').text(name);
									var txt;
									if(content.search(/\n+|\r+/)>0){
										txt = document.createElement('textarea');
										$(txt).addClass('d_tarea').val(content);
									}else{
										txt = document.createElement('input');
										$(txt).attr('type','text').addClass('d_inp').val(content);
									}
									r.find('span.d_check input[type="checkbox"]').attr('checked',false).click(function(){
										r.find('span.d_check input[type="checkbox"]').attr('checked',false);
										$(this).attr('checked',true);
									});
									r.append(txt).addClass('block').insertBefore(ir).fadeIn("slow");
							});
							$('#block').find('div').last().find('span').remove();
						}
					});	
				})();
			}
		}
	},
	Copyright:function(){
		$.ajax({
			url: ZC_BLOG_HOST+'ZB_USERS/PLUGIN/YTCMS/plugin.xml',
			type: 'GET',
			dataType: 'xml',
			data: { t:Math.random() },
			success: function(xml) {
				$('#version').text($('version',xml).text()+' '+$('pubdate',xml).text().replace(/^\d{2}|-/g,''));
				$('#author').attr('href',$('author>url',xml).text()).text($('author>name',xml).text());
				$('#email').text($('author>email',xml).eq(0).text());
				$('#give').attr('href',$('author>url',xml).eq(0).text());		
				$('#bug').attr('href',$('url',xml).eq(0).text());	
			}
		});
	},
	Demo:function(){
		$('#demo').find('li').first().next().html('');
		$('#demo').find('div').last().append('<span class="loading">loading...</span>');
		$.get('YT.Ajax.asp', { Action: 'Demo',t:Math.random() },function(txt){
			$('#demo').find('li').first().next().html(txt);
			$('#demo .loading').remove();
		});	
	},
	InsertText:function(obj,str,bool){  
		if(str == '') return;
		//为了兼容火狐
		var _length = obj.value.length;  
			obj.focus();  
		if(typeof document.selection!='undefined'){  
			if(bool){
				document.selection.createRange().text = str;  
			}else{
				document.selection.createRange().text = str.replace('@T','\n'+document.selection.createRange().text+'\n'); 	
			}
		}else {
			var restoreTop = obj.scrollTop;
			if(!bool){
				str = str.replace('@T','\n'+obj.value.substring(obj.selectionStart,obj.selectionEnd)+'\n'); 	
			}
			obj.value = obj.value.substr(0,obj.selectionStart)+str+obj.value.substring(obj.selectionEnd,_length);  
			if(restoreTop > 0){obj.scrollTop = restoreTop;}  
		}  
	},
	S:function(t){
		var _s = document.createElement('select');
		var _cms = YT.CMS;
		_s.className = 'YT';
		_s.options.add(new Option('「类型」','-1'));
		for(var _i=0;_i<_cms.length;_i++){
			_s.options.add(new Option(_cms[_i].YT.Text,_i));
		}
		_s.onchange = function(){
			$('.Parameters,.DataSource,.Save,.Fields').remove();
			if(this.value != -1){
				var _ds = _cms[this.value].DataSource;
				var _fields = _cms[this.value].YT.Fields;
				var __s = document.createElement('select');
					__s.className = 'DataSource';
					__s.options.add(new Option('「数据源」','-1'));
					for(var _i=0;_i<_ds.length;_i++){
						var _group = document.createElement('OPTGROUP');   
							_group.label = _ds[_i].Group;
							__s.appendChild(_group);
							for(var __i=0;__i<_ds[_i].DataSource.length;__i++){
								var _option = new Option();
									_option.value = _i+'|'+__i;
									_option.innerHTML = _ds[_i].DataSource[__i].Text;
									_group.appendChild(_option);
							}
					}
					__s.onchange = function(){
						$('.Parameters,.Save,.Fields').remove();
						if(this.value != -1){
							var _Parameters = _ds[parseInt(this.value.split('|')[0])].DataSource[parseInt(this.value.split('|')[1])].Parameters;
							for(var _i=0;_i<_Parameters.length;_i++){
								if(_Parameters[_i].Text.indexOf('分类')==0){
									var _select = $('#model').find('select').last().clone();
										_select.attr('class','Parameters').attr('lang',typeof(_Parameters[_i].Value)).attr('title',_Parameters[_i].Text).attr('size',5).css({width:'100%'});
										$(t).append(_select[0]);
								}
								if(_Parameters[_i].Text.indexOf('分类')==-1){
									var _input = document.createElement('input');
										_input.className = 'Parameters';
										_input.type = 'text';
										_input.lang = typeof(_Parameters[_i].Value);
										_input.title = _Parameters[_i].Text+',类型'+_input.lang;
										_input.value = _Parameters[_i].Value;
										$(t).append(_input);
								}
							}
							var ___s = document.createElement('select');
								___s.className = 'Fields';
								___s.options.add(new Option('「字段」','-1'));
								for(var _i=0;_i<_fields.length;_i++){
									var _group = document.createElement('OPTGROUP');   
										_group.label = _fields[_i].Group;
										___s.appendChild(_group);
										for(var __i=0;__i<_fields[_i].Fields.length;__i++){
											var _option = new Option();
												_option.value = _fields[_i].Fields[__i].Value;
												_option.innerHTML = _fields[_i].Fields[__i].Text;
												_group.appendChild(_option);
										}
								}
								___s.onchange = function(){
									YT.InsertText($(t).parent().find('textarea')[0],this.value,true);
									___s.selectedIndex = 0;
								}
								$(t).append(___s);
							var _but = document.createElement('input');
								_but.className = 'Save';
								_but.type = 'button';
								_but.value = 'CODE';
								_but.onclick = function(){
									var _t = '{YT:@YT DataSource="@D"}@T{/YT:@YT}';
									var _obj = _cms[$('.YT').val()];
									var _DataSource = _obj.DataSource[parseInt($('.DataSource').val().split('|')[0])].DataSource[parseInt($('.DataSource').val().split('|')[1])];
									var _Parameters = [];
										$('.Parameters').each(function(){
											if($(this).attr('lang') == 'string'){
												_Parameters.push("'"+$(this).val()+"'");
											}else{
												_Parameters.push($(this).val());	
											}
										});
										_t = _t.replace(/\@YT/ig,_obj.YT.Value);
										_t = _t.replace(/\@D/ig,_DataSource.Value+'('+_Parameters.join(',')+')');
										YT.InsertText($(t).parent().find('textarea')[0],_t,false);
								};
								$(t).append(_but);
						}
					};
					$(t).append(__s);
			}
		};
		$(t).append(_s);
	}
};
jQuery.extend({
	toJSONString:function(object){
		var type=typeof object;
		if('object'==type){
			if(Array==object.constructor){
				type='array';
			}else if(RegExp==object.constructor){
				type = 'regexp';
			}else{
				type = 'object';
			}
		}
		switch(type){
			case 'undefined':
			case 'unknown':
				return;
			break;
			case 'function':
			case 'boolean':
			case 'regexp':
				return object.toString();
			break;
			case 'number':
				return isFinite(object) ? object.toString() : 'null';
			break;
			case 'string':
				return '"' + object.replace(/(\\|\")/g, "\\$1").replace(/\n|\r|\t/g, function() {
					var a = arguments[0];
					return (a == '\n') ? '\\n': (a == '\r') ? '\\r': (a == '\t') ? '\\t': ""
				}) + '"';
				break;
			case 'object':
				if (object === null) return 'null';
				var results = [];
				for (var property in object) {
					var value = jQuery.toJSONString(object[property]);
					if (value !== undefined){
						results.push(jQuery.toJSONString(property) + ':' + value);
					}
				}
				return '{' + results.join(',') + '}';
				break;
			case 'array':
				var results = [];
				for (var i = 0; i < object.length; i++) {
					var value = jQuery.toJSONString(object[i]);
					if (value !== undefined){
						results.push(value);
					}
				}
				return '[' + results.join(',') + ']';
				break;
			}
		}
});