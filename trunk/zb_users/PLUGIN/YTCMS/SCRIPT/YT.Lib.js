///////////////////////////////////////////////////////////////////////////////
// 作	 者:    	瑜廷
// 技术支持:     33195@qq.com
// 程序名称:    	YT.CMS Script
// 开始时间:    	2011-05-28
// 最后修改:    	2012-08-08
// 备	 注:    	only for YT.CMS
///////////////////////////////////////////////////////////////////////////////
var YT = {
	Panel:{
		Load:function(){
			$.get('YT.Ajax.asp', { Action: 'tplList',t:Math.random() },
			function(json){
				var d = eval('('+json+')');
				for(var i=0;i<d.length;i++){
					var r=$('#tplList').find('div.row').hide().clone();
						r.text(d[i]).attr('title',d[i]).mousedown(function(e){
							var _e=this,evt=e;
							evt.stopPropagation();
							$(this).mouseup(function(e){
								e.stopPropagation();
								$(this).unbind('mouseup');
								if(evt.button==2) {
									var menu = $('#currPanel');
									/b_article\-(multi|single)\-[a-z]+\.html/i.test(_e.title)?menu.find('.diyBlock').show():menu.find('.diyBlock').hide();
									var o = {}, x, y;
									if( self.innerHeight ) {
										o.pageYOffset = self.pageYOffset;
										o.pageXOffset = self.pageXOffset;
										o.innerHeight = self.innerHeight;
										o.innerWidth = self.innerWidth;
									} else if( document.documentElement &&
										document.documentElement.clientHeight ) {
										o.pageYOffset = document.documentElement.scrollTop;
										o.pageXOffset = document.documentElement.scrollLeft;
										o.innerHeight = document.documentElement.clientHeight;
										o.innerWidth = document.documentElement.clientWidth;
									} else if( document.body ) {
										o.pageYOffset = document.body.scrollTop;
										o.pageXOffset = document.body.scrollLeft;
										o.innerHeight = document.body.clientHeight;
										o.innerWidth = document.body.clientWidth;
									}
									(e.pageX) ? x = e.pageX : x = e.clientX + o.scrollLeft;
									(e.pageY) ? y = e.pageY : y = e.clientY + o.scrollTop;
									
									$(document).unbind('click');
									$(menu).css({ top: y, left: x }).show();
									$(menu).find('LI.tree').unbind('click');
									$(menu).find('LI.tree').each(function(){
										$(this).click(function(){
											if($(this).find('div').css('display')=='none'){
												$(this).find('div').show();
											}else{
												$(this).find('div').hide();
											}
											return false;
										});
									});
									$(menu).find('LI:not(.tree) A').unbind('click');
									$(menu).find('LI:not(.tree) A').each(function(){
										$(this).click(function(){
											$(menu).hide();
											switch($(this).attr('href').substr(1)){
												case 'CBLOCK':
													YT.Panel.Block.C();
												break;
												case 'MBLOCK':
													YT.Panel.Block.M();
												break;
												case 'CMODEL':
													YT.Panel.Model.C();
												break;
												case 'MMODEL':
													YT.Panel.Model.M();
												break;
												case 'CTPL':
													YT.Panel.TPL.C({title:_e.title,type:-1});
												break;
												case 'MTPL':
													YT.Panel.TPL.M();
												break;
											}
											return false;
										});
									});
									setTimeout( function() {
										$(document).click( function() {
											$(document).unbind('click').unbind('keypress');
											$(menu).hide();
											return false;
										});
									}, 0);
								}
								if( $.browser.mozilla ) {
									$('#currPanel').css({ 'MozUserSelect' : 'none' });
								} else if( $.browser.msie ) {
									$('#currPanel').bind('selectstart.disableTextSelect', function() { return false; });
								} else {
									$('#currPanel').bind('mousedown.disableTextSelect', function() { return false; });
								}
								$(document).bind('contextmenu', function() { return false; });
							});
						}).click(function(){
							var t = YT.Panel.ModalDialog($('#Template').html());
							var e = this;
							$.get('YT.Ajax.asp', { Action: 'GetFile',File:e.title,t:Math.random() },function(txt){
								$($(t).find('li')[0]).text(e.title);
								YT.S(t);
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
											_b.options.add(new Option($('Name',Block).text(),'<#CACHE_INCLUDE_'+$('Name',Block).text().toUpperCase()+'#>'));
										});
										_b.onchange = function(){
											YT.InsertText($(t).find('textarea')[0],this.value,true);
											_b.selectedIndex = 0;
										}
										$(t).find('li')[1].appendChild(_b);	
									}
								});	
								$(t).find('textarea').css({height:'380px'}).text(txt);
								$(t).find('input').click(function(){
									$.ajax({
										url: 'YT.Ajax.asp',
										type: 'POST',
										dataType: 'html',
										data: { Action: 'SaveFile' ,Name: $($(t).find('li')[0]).text(), Content: $(t).find('textarea').val() , t:Math.random() },
										success: function() {
											$(t).find('span').trigger('click');
										}
									});						  
								});
							});
						});
						r.attr('class','ready').fadeIn('slow').hover(function() {
							$(this).addClass('color2')
						}, function() {
							$(this).removeClass('color2')
						}).appendTo($('#tplList'));
				}
			});
		},
		Model:{
			C:function(n){
				var _Panel = YT.Panel.ModalDialog($('.Model').eq(0).html());
					_Panel.find('#Step1').find('div').html(YT.Panel.Model.row());
					_Panel.find('#Step2').hide();
					_Panel.find('#Add').click(function(){
						_Panel.find('#Step1').find('div').append(YT.Panel.Model.row()).find('em').unbind('click').click(function(){
							if(_Panel.find('#Step1').find('div').find('p').length>1){
								$(this).parent().remove();
							}
						});
					});
					_Panel.find('label>select').change(function(){
						var e=$(this);
						var b=false;
						_Panel.find('#Step1').find('div').find('p').each(function(){
							var l=$(this).find('input').eq(0).val();
							if(e.val()==0){
								if(l.toUpperCase()=='YT_Money'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Type'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Fee'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Payment'.toUpperCase()){$(this).remove();}
							}else if(e.val()==1){
								if(l.toUpperCase()=='YT_Logistics_Type'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Fee'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Payment'.toUpperCase()){$(this).remove();}
							}
							if(l=='YT_Money'){
								b=true;
								return;	
							}
						});
						if(!b){
							_Panel.find('#Add').trigger('click');
							var p=_Panel.find('#Step1').find('div').find('p:last');
								p.find('input').eq(0).val('YT_Money');
								p.find('input').eq(1).val('金额');
								p.find('input').eq(2).val('100');
								p.find('select').eq(0).val('VARCHAR');
						}
						if(e.val()==2){
							_Panel.find('#Add').trigger('click');
							var jsonAry=[];
							var p=_Panel.find('#Step1').find('div').find('p:last');
								p.find('input').eq(0).val('YT_Logistics_Type');
								p.find('input').eq(1).val('物流类型');
								jsonAry.push({t:'平邮',v:'POST'});
								jsonAry.push({t:'EMS',v:'EMS'});
								jsonAry.push({t:'快递',v:'EXPRESS'});
								p.find('input').eq(2).val($.toJSONString(jsonAry));
								p.find('select').eq(0).val('VARCHAR');
								p.find('select').eq(1).val('select');
							_Panel.find('#Add').trigger('click');
							var p=_Panel.find('#Step1').find('div').find('p:last');
								p.find('input').eq(0).val('YT_Logistics_Fee');
								p.find('input').eq(1).val('物流费用');
								p.find('input').eq(2).val(0);
								p.find('select').eq(0).val('VARCHAR');
								p.find('select').eq(1).val('text');
							_Panel.find('#Add').trigger('click');
							var p=_Panel.find('#Step1').find('div').find('p:last');
								p.find('input').eq(0).val('YT_Logistics_Payment');
								p.find('input').eq(1).val('物流支付方式');
								jsonAry=[];
								jsonAry.push({t:'卖方付款',v:'SELLER_PAY'});
								jsonAry.push({t:'买方付款',v:'BUYER_PAY'});
								p.find('input').eq(2).val($.toJSONString(jsonAry));
								p.find('select').eq(0).val('VARCHAR');
								p.find('select').eq(1).val('select');
								
						}
					});
					if(/\d+/.test(n)){
						_Panel.find('#Step1').find('div').html('');	
					}
					_Panel.find('#Next').click(function(){
						var _Json = {
								Table:{Name:_Panel.find('input').eq(0).val(),Description:_Panel.find('input').eq(1).val(),Bind:''},
								Fields:[]
							};
							if(_Json.Table.Name==''){
								_Panel.find('input').eq(0).attr('title','请输入模型名称').css('border','red 1px solid').focus();
								return;
							}
							$.ajax({
								url: YT_CMS_XML_URL+YTConfig.Model,
								type: 'GET',
								dataType: 'xml',
								data: { t:Math.random() },
								success: function(xml) {
									var b=false;
									$('Model', xml).each(function(i) {
										var Model = $('Model', xml).get(i);
										if(_Json.Table.Name==$('Table>Name',Model).text()){
											b=true;
											return;
										}
									});
									if(b&&!/\d+/.test(n)){
										_Panel.find('input').eq(0).attr('title','存在相同名称的模型').css('border','red 1px solid').focus();
									}else{
										_Panel.find('input').eq(0).attr('title','').css('border','');
										var f=false;
										_Panel.find('#Step1').find('p').each(function(){
											if($(this).find('input').eq(0).val()==''){
												f=true;
												$(this).find('input').eq(0).attr('title','请输入字段名称').css('border','red 1px solid').focus();
												return;	
											}else{
												$(this).find('input').eq(0).attr('title','').css('border','')
												_Json.Fields.push({
													Name:$(this).find('input').eq(0).val(),
													Description:$(this).find('input').eq(1).val(),
													Value:$(this).find('input').eq(2).val(),
													Property:$(this).find('select').eq(0).val(),
													Type:$(this).find('select').eq(1).val()
												});	
											}
										});
										if(!f){
											_Panel.find('#Step1').hide();
											_Panel.find('#Step2').fadeIn('slow');
											_Panel.find('#Step2').find('input').click(function(){
													var _Bind = [];
													$(this).parent().find('select').eq(0).children().each(function(){
														if($(this).attr('selected')){
															_Bind.push($(this).val())	
														}													 
													});
													_Json.Table.Bind = _Bind.join(',');
													if(_Json.Table.Bind==''){
														$(this).parent().find('select').attr('title','请绑定栏目,支持多选.').css('border','red 1px solid');
													}else{
														$.ajax({
															url: 'YT.Ajax.asp',
															type: 'POST',
															dataType: 'json',
															data: { Action: (/\d+/.test(n)?'Update':'Save')+'Model' ,
															Index:n,Json:$.toJSONString(_Json), t:Math.random() },
															success: function(result) {
																if(result){
																	$(_Panel).find('span').trigger('click');
																}
															}
														});	
													}							
											});
										}	
									}
								}
							});	
				});	
				if(/\d+/.test(n)){
					$.ajax({
						url: YT_CMS_XML_URL+YTConfig.Model,
						type: 'GET',
						dataType: 'xml',
						data: { t:Math.random() },
						success: function(xml) {
							$('Model', xml).each(function(i) {
								if(n==i){
									var Model = $('Model', xml).get(i);
									_Panel.find('input').eq(0).val($('Table>Name',Model).text());
									_Panel.find('input').eq(1).val($('Table>Description',Model).text());
									$('Field',Model).each(function(i){
										_Panel.find('#Add').trigger('click');
										_Panel.find('#Step1').find('div').find('p').eq(i).find('input').eq(0).val($(this).find('Name').text()); 
										_Panel.find('#Step1').find('div').find('p').eq(i).find('input').eq(1).val($(this).find('Description').text());
										_Panel.find('#Step1').find('div').find('p').eq(i).find('input').eq(2).val($(this).find('Value').text());
										_Panel.find('#Step1').find('div').find('p').eq(i).find('select').eq(0).val($(this).find('Property').text());
										_Panel.find('#Step1').find('div').find('p').eq(i).find('select').eq(1).val($(this).find('Type').text());
									});
									var s=$('Table>Bind',Model).text().split(',');
									for(var i=0;i<s.length;i++){
										_Panel.find('#Step2').find('select').eq(0).children().each(function(){
											if($(this).val()==s[i]){
												$(this).attr('selected','selected');	
											}														  
										});	
									}
									return;	
								}						  
							});
						}
					});
				}
			},
			M:function(){
				var _Panel = YT.Panel.ModalDialog($('.Model').eq(1).html());
				$('.readyModel').remove();
				$.ajax({
					url: YT_CMS_XML_URL+YTConfig.Model,
					type: 'GET',
					dataType: 'xml',
					data: { t:Math.random() },
					success: function(xml) {
						$('Model', xml).each(function(i) {
							var Model = $('Model', xml).get(i);
							var r = _Panel.find('tr').eq(1).clone().attr('class','readyModel');
								r.find('td').eq(0).text($('Table>Name',Model).text())			//表
								r.find('td').eq(1).text($('Table>Description',Model).text())	//描述
							var _Bind = $('Table>Bind',Model).text().split(',');
								r.find('td').eq(2).text('');
								var j=[];
								for(var _i=0;_i<_Bind.length;_i++){
									$('#Step2').find('select').eq(0).find('option').each(function(){
										if(parseInt($(this).val()) == parseInt(_Bind[_i])){
											j.push('<em>'+$(this).text()+'</em>');
										}												
									});
								}
								r.find('td').eq(2).append(j.join(','));	//绑定
							var a=false;
								for(var i2=0;i2<YTConfig.Table.length;i2++){
									if(YTConfig.Table[i2]==$('Table>Name',Model).text()){a=true;break;}	
								}
								r.find('td').eq(3).text(a?'系统':'用户');
								if(a){
									r.find('td').eq(4).html('<em>已安装</em>');
								}else{
									//检测模型状态
									$.ajax({
										url: 'YT.Ajax.asp',
										type: 'POST',
										dataType: 'json',
										data: { Action:'Exist',Name:r.find('td').eq(0).text(), t:Math.random() },
										success: function(bool) {
											r.find('td').eq(4).html('<a href="#" rel="'+!bool+'">'+(bool?'已安装':'未安装')+'</a> <a href="#">更新</a> <a href="#">删除</a>');
											r.find('td').eq(4).find('a').eq(0).click(function(){
												var _Install = $(this).attr('rel');
												if(_Install.toString() != 'true' && !confirm('卸载表将会删除当前表的所有数据,仍进行此操作?')){
													return false;	
												}
												//变更模型状态
												$.ajax({
													url: 'YT.Ajax.asp',
													type: 'POST',
													data: { Action:bool?'UnInstall':'Install',Index:i,t:Math.random() },
													success: function() {
														YT.Panel.Model.M();
													}
												});	
											});
											r.find('td').eq(4).find('a').eq(1).click(function(){
												YT.Panel.Model.C(i);
											});
											r.find('td').eq(4).find('a').eq(2).click(function(){
												if($(this).parent().find('a').eq(0).attr('rel')=='true'){
													$.ajax({
														url: 'YT.Ajax.asp',
														type: 'POST',
														dataType: 'html',
														data: { Action: 'DelModel' ,Index:i, t:Math.random() },
														success: function() {
															YT.Panel.Model.M();
														}
													});	
												}else{
													$(this).parent().parent().css('background','#CCC').attr('title','已安装的表无法进行此操作,请先卸载表!');
												}											  
											});
										}
									});	
								}
								r.hover(function() {
									$(this).addClass('color1')
								}, function() {
									$(this).removeClass('color1')
								}).insertBefore(_Panel.find('tr').eq(1))
						});
					}
				});	
					
			},
			row:function(){return $('<p>'+$('.Model').eq(0).find('#Step1').find('div').html()+'<em>x</em></p>');}
		},
		Analysis:function(){
			$('#cmbCate').change(function(){
				$('#model').parent().find('p').each(function(){
					if($(this).attr('id') != 'model'){
						$(this).remove();	
					}
				});
				var _Cate = $(this).val();
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
							for(var _i=0;_i<_Bind.length;_i++){
								if(parseInt(_Cate) == parseInt(_Bind[_i])){
									_isBind = true;
									break;
								}
							}
							if(_isBind){
								$('Field',Model).each(function(ii){
									switch($(this).find('Type').text()){
										case 'text':
											var _i = document.createElement('input');
												_i.type = $(this).find('Type').text();
												_i.value = $(this).find('Value').text();
												_i.name = $(this).find('Name').text();
												_i.style.width = '50%';
												$('#model').parent().append('<p>'+$(this).find('Description').text()
												+':</p>').find('p').eq(ii+1).attr('title',$(this).find('Type').text()).append(_i);
										break;
										case 'select':
											var _s = document.createElement('select');
												_s.name = $(this).find('Name').text();
												try{
													var _v = jsonToObject($(this).find('Value').text());
													for(var _i=0;_i<_v.length;_i++){
														_s.options.add(new Option(_v[_i].t,_v[_i].v));	
													}
												}catch(e){
													_v = $(this).find('Value').text().split(',');	
													for(var _i=0;_i<_v.length;_i++){
														_s.options.add(new Option(_v[_i],_v[_i]));	
													}
												}
												$('#model').parent().append('<p>'+$(this).find('Description').text()
												 +':</p>').find('p').eq(ii+1).attr('title',$(this).find('Type').text()).append(_s);
										break;
										case 'checkbox':
											var row = $('#model').parent().append('<p>'+$(this).find('Description').text()
												 +':</p>').find('p').eq(ii+1).attr('title',$(this).find('Type').text());
												try{
													var __v = jsonToObject($(this).find('Value').text());
													for(var _i=0;_i<__v.length;_i++){
														var __i = document.createElement('input');
															__i.type = $(this).find('Type').text();
															__i.value = __v[_i].v;
															__i.name = $(this).find('Name').text();
															row.append(__v[_i].t).append(__i);
													}
												}catch(e){
													__v = $(this).find('Value').text().split(',');	
													for(var _i=0;_i<__v.length;_i++){
														var __i = document.createElement('input');
															__i.type = $(this).find('Type').text();
															__i.value = __v[_i];
															__i.name = $(this).find('Name').text();
															row.append(__v[_i]).append(__i);
													}	
												}
										break;
										case 'textarea':
											var _t = document.createElement('textarea');
												_t.value = $(this).find('Value').text();
												_t.name = $(this).find('Name').text();
												_t.style.width = '50%';
												$('#model').parent().append('<p>'+$(this).find('Description').text()
												 +':</p>').find('p').eq(ii+1).attr('title',$(this).find('Type').text()).append(_t);
										break;
									}
								});
								if($('#edtID').val()!=0){
									$.ajax({
										url: ZC_BLOG_HOST+'ZB_USERS/PLUGIN/YTCMS/YT.Ajax.asp',
										type: 'POST',
										dataType: 'html',
										data: { Action:'GetData', Name:$('Table>Name',Model).text(), ID:$('#edtID').val(), t:Math.random() },
										success: function(r) {
											r=eval('('+r+')');
											$('#model').parent().find('p').each(function(j){
												if($(this).attr('id')!='model'){
													switch($(this).attr('title')){
														case 'text':
															var v=k(r,$(this).find('input')[0].name);
															if(v!=null){$(this).find('input').val(unescape(v));}
														break;
														case 'select':
															var v=k(r,$(this).find('select')[0].name);
															if(v!=null){$(this).find('select').val(unescape(v));}
														break;
														case 'checkbox':
															var v=k(r,$(this).find('checked')[0].name);
															if(v!=null){
																var _c = unescape(v).split(',');
																	$(this).find('checked').each(function(){
																		for(var _j=0;_j<_c.length;_j++){
																			if($(this).val().toLowerCase().replace(/\s+/ig,'') == _c[_j].toLowerCase().replace(/\s+/ig,'')){
																				$(this).attr('checked',true);
																				break;
																			}	
																		}
																	});
															}
														break;
														case 'textarea':
															var v=k(r,$(this).find('textarea')[0].name);
															if(v!=null){$(this).find('textarea').val(unescape(v));}
														break;
													}
												}			  
											});
										}
									});	
								}
								return false;
							}
						});	
					}
				});
				function k(j,key){
					for(var i=0;i<j.length;i++){
						if(j[i].Name==key&&j[i].Value!=''){return j[i].Value;}	
					}
					return null;
				}
			});
			$('#cmbCate').trigger('change');
		},
		Block:{
			C:function(n){
				var t = YT.Panel.ModalDialog($('#Template').html());
					$(t).find('li').eq(0).html('<input type="text" />');
					YT.S(t);
					$(t).find('textarea').css({height:'380px'}).text();
					$(t).find('input').eq(1).click(function(){
						var _Json={};
							_Json.Name=$(t).find('input').eq(0).val();
							_Json.Content=$(t).find('textarea').val();
						if(_Json.Name==''){
							$(t).find('input').eq(0).attr('title','请输入模块名称').css('border','red 1px solid').focus();
							return;
						}
						$.ajax({
							url: YT_CMS_XML_URL+YTConfig.Block,
							type: 'GET',
							dataType: 'xml',
							data: { t:Math.random() },
							success: function(xml) {
								var b=false;
								$('Block', xml).each(function(i) {
									var Block = $('Block', xml).get(i);
									if(_Json.Name==$('Name',Block).text()){
										b=true;
										return;
									}
								});
								if(b&&!/\d+/.test(n)){
									$(t).find('input').eq(0).attr('title','存在相同名称的模块').css('border','red 1px solid').focus();
								}else{
									$(t).find('input').eq(0).attr('title','').css('border','');	
									$.ajax({
										url: 'YT.Ajax.asp',
										type: 'POST',
										dataType: 'json',
										data: { Action: (/\d+/.test(n)?'Update':'Save')+'Block' ,Index:n,Json:$.toJSONString(_Json), t:Math.random() },
										success: function(result) {
											if(result){
												$(t).find('span').trigger('click');
											}
										}
									});
								}
							}
						});				  
					});
					if(/\d+/.test(n)){
						$.ajax({
							url: YT_CMS_XML_URL+YTConfig.Block,
							type: 'GET',
							dataType: 'xml',
							data: { t:Math.random() },
							success: function(xml) {
								$('Block', xml).each(function(i) {
									if(n==i){
										var Block = $('Block', xml).get(i);
										$(t).find('input').eq(0).val($('Name',Block).text());
										$(t).find('textarea').val($('Content',Block).text());
										return;	
									}						  
								});
							}
						});
					}
			},
			M:function(){
				var _Panel = YT.Panel.ModalDialog($('.Block').html());
				$('.readyBlock').remove();
				$.ajax({
					url: YT_CMS_XML_URL+YTConfig.Block,
					type: 'GET',
					dataType: 'xml',
					data: { t:Math.random() },
					success: function(xml) {
						$('Block', xml).each(function(i) {
							var Block = $('Block', xml).get(i);
							var r = _Panel.find('tr').eq(1).clone().attr('class','readyBlock');
								r.find('td').eq(0).text($('Name',Block).text())
								r.find('td').eq(1).html('<a href="#">更新</a> <a href="#">删除</a>');
								r.find('td').eq(1).find('a').eq(0).click(function(){
									YT.Panel.Block.C(i);
								});
								r.find('td').eq(1).find('a').eq(1).click(function(){
									$.ajax({
										url: 'YT.Ajax.asp',
										type: 'POST',
										dataType: 'html',
										data: { Action: 'DelBlock' ,Index:i, t:Math.random() },
										success: function() {
											YT.Panel.Block.M();
										}
									});								  
								});
								r.hover(function() {
									$(this).addClass('color1')
								}, function() {
									$(this).removeClass('color1')
								}).insertBefore(_Panel.find('tr').eq(1))
						});
					}
				});	
			}
		},
		TPL:{
			C:function(n){
				var t=YT.Panel.ModalDialog($('#Step2').html());
				var s=document.createElement('input');
					s.value=/b_article-multi-[a-z]+\.html/.test(n.title)?'Multi':'Single';
					s.type='hidden';
					t.find('div').append(s);
					t.find('div').append(n.title);
					t.find('select').eq(0).css({width:'100%',height:'300px'});
					var _Json={};
						_Json.File=n.title;
						_Json.Type=s.value;
						l(_Json);
					if(n.type!=-1){t.find('select').eq(1).val(n.type);};
					t.find('select').eq(1).trigger('change');
					t.find('input[type="button"]').click(function(){
						var j=[];
						t.find('select').eq(0).find('option:selected').each(function(){
							j.push($(this).val());					   
						});
						_Json.Bind=j;
						_Json.Type=s.value;
						$.ajax({
							url: YT_CMS_XML_URL+YTConfig.TPL,
							type: 'GET',
							dataType: 'xml',
							data: { t:Math.random() },
							error: function(x){
								if(x.status==404){
									k({action:'SaveTPL',json:_Json,index:-1});
								}
							},
							success: function(xml) {
								var b=false,index=-1;
								$('TPL',xml).each(function(i){
									var TPL=$('TPL',xml).get(i);
									if($('File',TPL).text()==_Json.File&&$('Type',TPL).text()==_Json.Type){
										b=true;
										index=i;
										return;
									}					   
								});
								if(j.length>0){
									if(b){
										k({action:'UpdateTPL',json:_Json,index:index});
									}else{
										k({action:'SaveTPL',json:_Json,index:-1});	
									}
								}
							}
						});	
					});
					function k(o){
						$.ajax({
							url: 'YT.Ajax.asp',
							type: 'POST',
							dataType: 'json',
							data: { Action: o.action ,Index:o.index,Json:$.toJSONString(o.json), t:Math.random() },
							success: function(result) {
								if(result){
									$(t).find('span').trigger('click');
								}
							}
						});		
					}
					function l(_Json){
						$.ajax({
							url: YT_CMS_XML_URL+YTConfig.TPL,
							type: 'GET',
							dataType: 'xml',
							data: { t:Math.random() },
							success: function(xml) {
								var b=false,index=-1;
								$('TPL',xml).each(function(i){
									var TPL=$('TPL',xml).get(i);
									if($('File',TPL).text()==_Json.File&&$('Type',TPL).text()==_Json.Type){
										b=true;index=i;return;
									}					   
								});
								t.find('select').eq(0).children().each(function(){
									$(this).attr('selected',false);
								});	
								if(b){
									var TPL=$('TPL',xml).get(index);
									var Bind=$('Bind',TPL).text();
									t.find('select').eq(0).children().each(function(){
										var b=Bind.split(',');
										for(var i=0;i<b.length;i++){
											if($(this).val()==b[i]){
												$(this).attr('selected',true);
												break;
											}	
										}
									});
								}
							}
						});	
					}
			},
			M:function(){
				var _Panel = YT.Panel.ModalDialog($('.TPL').html());
				$('.readyTPL').remove();
				$.ajax({
					url: YT_CMS_XML_URL+YTConfig.TPL,
					type: 'GET',
					dataType: 'xml',
					data: { t:Math.random() },
					success: function(xml) {
						$('TPL', xml).each(function(i) {
							var TPL = $('TPL', xml).get(i);
							var r = _Panel.find('tr').eq(1).clone().attr('class','readyTPL');
								r.find('td').eq(0).text($('File',TPL).text());
								r.find('td').eq(1).text('');
							var Bind = $('Bind',TPL).text().split(',');
								var j=[];
								for(var x=0;x<Bind.length;x++){
									$('#Step2').find('select>option').each(function(){
										if(parseInt($(this).val()) == parseInt(Bind[x])){
											j.push('<em>'+$(this).text()+'</em>');
										}												
									});
								}
								r.find('td').eq(1).append(j.join(','));	//绑定
								r.find('td').eq(2).text($('Type',TPL).text()=='Single'?'文章页':'列表页');
								r.find('td').eq(3).html('<a href="#">更新</a> <a href="#">删除</a>');
								r.find('td').eq(3).find('a').eq(0).click(function(){
									YT.Panel.TPL.C({title:r.find('td').eq(0).text(),type:$('Type',TPL).text()});
								});
								r.find('td').eq(3).find('a').eq(1).click(function(){
									$.ajax({
										url: 'YT.Ajax.asp',
										type: 'POST',
										dataType: 'html',
										data: { Action: 'DelTPL' ,Index:i, t:Math.random() },
										success: function() {
											YT.Panel.TPL.M();
										}
									});								  
								});
								r.hover(function() {
									$(this).addClass('color1')
								}, function() {
									$(this).removeClass('color1')
								}).insertBefore(_Panel.find('tr').eq(1))
						});
					}
				});	
			}
		},
		ModalDialog:function(text){
			$('#Panel').find('div').css({
				position:'absolute',
				width:'100%',
				opacity:1
			}).html(text);
			$('#Panel').css({
				height:$(document).height(),
				left:0,
				top:0,
				width:'100%',
				opacity:0.9
			}).fadeIn('slow');
			$('#Panel').find('span').show().click(function(){
				$(this).parent().fadeOut('slow');
			});
			$(document).keypress(function(e){
				if(e.keyCode==27){
					$('#Panel').find('span').trigger('click');	
				}
			});
			return $('#Panel');
		}
	},
	Copyright:function(){
		$.ajax({
			url: ZC_BLOG_HOST+'ZB_USERS/PLUGIN/YTCMS/plugin.xml',
			type: 'GET',
			dataType: 'xml',
			data: { t:Math.random() },
			success: function(xml) {
				var s=' ',h=$('#headerWelcome'),f=$('#footerWelcome');
					h.append('<A href="'+$('url',xml).eq(0).text()+'"></A>');
					h.find('A').append($('note',xml).eq(0).text());
					f.append('邮箱:'+$('author>email',xml).text());
					f.append(s);
					f.append('作者:'+$('author>name',xml).text());
					f.append(s);
					f.append($('name',xml).eq(0).text());
					f.append(s);
					f.append('插件版本:'+$('version',xml).text());
					f.append(s);
					f.append('适用于:'+$('adapted',xml).text());
					f.append(s);
					f.append('开始时间:'+$('pubdate',xml).text());
					f.append(s);
					f.append('结束时间:'+$('modified',xml).text());
			}
		});
	},
	CMS:[],
	InsertText:function(obj,str,bool){  
		if(str == '') return;
		//为了兼容火狐
		var _length = obj.value.length;  
			obj.focus();  
		if(typeof document.selection!='undefined'){  
			if(bool){
				document.selection.createRange().text = str;  
			}else{
				document.selection.createRange().text = str.replace('@T',document.selection.createRange().text); 	
			}
		}else {
			if(!bool){
				str = str.replace('@T',obj.value.substring(obj.selectionStart,obj.selectionEnd)); 	
			}
			obj.value = obj.value.substr(0,obj.selectionStart)+str+obj.value.substring(obj.selectionEnd,_length);  
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
									var _select = $('#Step2').find('select').clone();
										_select.attr('class','Parameters').attr('lang',typeof(_Parameters[_i].Value)).attr('title',_Parameters[_i].Text).attr('size',5).css({width:'100%'});
										$(t).find('li')[1].appendChild(_select[0]);
								}
								if(_Parameters[_i].Text.indexOf('分类')==-1){
									var _input = document.createElement('input');
										_input.className = 'Parameters';
										_input.type = 'text';
										_input.lang = typeof(_Parameters[_i].Value);
										_input.title = _Parameters[_i].Text+',类型'+_input.lang;
										_input.value = _Parameters[_i].Value;
										$(t).find('li')[1].appendChild(_input);
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
									YT.InsertText($(t).find('textarea')[0],this.value,true);
									___s.selectedIndex = 0;
								}
								$(t).find('li')[1].appendChild(___s);
							var _but = document.createElement('input');
								_but.className = 'Save';
								_but.type = 'button';
								_but.value = 'CODE';
								_but.onclick = function(){
									var _t = '<YT:@YT DataSource="@D">@T</YT>';
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
										YT.InsertText($(t).find('textarea')[0],_t,false);
								};
								$(t).find('li')[1].appendChild(_but);
						}
					};
					$(t).find('li')[1].appendChild(__s);
			}
		};
		$(t).find('li')[1].appendChild(_s);
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