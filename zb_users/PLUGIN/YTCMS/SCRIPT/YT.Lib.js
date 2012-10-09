///////////////////////////////////////////////////////////////////////////////
// 作	 者:    	瑜廷
// 技术支持:     33195@qq.com
// 程序名称:    	YT.CMS Script
// 开始时间:    	2011-05-28
// 最后修改:    	2012-08-08
// 备	 注:    	only for YT.CMS
///////////////////////////////////////////////////////////////////////////////
var YT = {
	CMS:[],
	ESC:null,
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
												case 'SQL':
													YT.Panel.SQL();
												break;
												case 'DEMO':
													YT.Demo();
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
								$(t).find('input').eq(0).click(function(){
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
							if(l.toUpperCase()=='YT_Service'.toUpperCase()){$(this).remove();}
							if(e.val()==0){
								if(l.toUpperCase()=='YT_Money'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Type'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Fee'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Payment'.toUpperCase()){$(this).remove();}
							}else{
								if(l.toUpperCase()=='YT_Logistics_Type'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Fee'.toUpperCase()){$(this).remove();}
								if(l.toUpperCase()=='YT_Logistics_Payment'.toUpperCase()){$(this).remove();}
							}
							if(l=='YT_Money'){
								b=true;
								return;	
							}
						});var jsonAry=[],jsonAry2=[];
						if(!b){
							_Panel.find('#Add').trigger('click');
							var p=_Panel.find('#Step1').find('div').find('p:last');
								p.find('input').eq(0).val('YT_Money');
								p.find('input').eq(1).val('金额');
								p.find('input').eq(2).val('100');
								p.find('select').eq(0).val('VARCHAR');
								p.find('select').eq(1).val('text');
						}
						if(e.val()==2||e.val()==3){
							_Panel.find('#Add').trigger('click');
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
						if(e.val()>0){
							_Panel.find('#Add').trigger('click');
							var p=_Panel.find('#Step1').find('div').find('p:last');
								p.find('input').eq(0).val('YT_Service');
								p.find('input').eq(1).val('服务类型');
								jsonAry=[];
								jsonAry.push({t:'即时交易',v:'CREATE_DIRECT_PAY_BY_USER'});
								jsonAry.push({t:'担保交易',v:'CREATE_PARTNER_TRADE_BY_BUYER'});
								jsonAry.push({t:'双功能',v:'TRADE_CREATE_BY_BUYER'});
								for(var j=0;j<jsonAry.length;j++){
									if((e.val()-1)==j){
										jsonAry2.push(jsonAry[j]);	
									}
								}
								for(var j=0;j<jsonAry.length;j++){
									if((e.val()-1)!=j){
										jsonAry2.push(jsonAry[j]);	
									}
								}
								p.find('input').eq(2).val($.toJSONString(jsonAry2));
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
					for(var i=0;i<j.YTARRAY.length;i++){
						if(j.YTARRAY[i]==key){return eval('j.'+key);}	
					}
					return null;
				}
			});
			try{$('#cmbCate').trigger('change');}catch(e){}
			(function(){  
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
			})();
		},
		SQL:function(){
			var t=YT.Panel.ModalDialog('<ul></ul>');
				$.ajax({
					url: 'YT.Ajax.asp',
					type: 'POST',
					dataType: 'html',
					data: { Action: 'ImportList' , t:Math.random() },
					success: function(json) {
						if(json!=''){
							json=json.split(',');
							for(var i=0;i<json.length;i++){
								t.find('ul').append('<div class="SQL"><em>'+json[i]+'</em><font></font></div>');	
							}
							t.find('ul').append('<input type="button" value="全部导入" />');
							t.find('input').click(function(){
								if(confirm('blog_Category,blog_Article表将被清空数据,仍进行此操作?')){
									t.find('div.SQL').each(function(){
										var e=this;
										$.ajax({
											url: 'YT.Ajax.asp',
											type: 'POST',
											dataType: 'html',
											data: { Action: 'Import' ,Name:$(e).find('em').text(), t:Math.random() },
											success: function() {
												$(e).find('font').css('color','blue').text(' √');
											}
										});						
									});
								}
							});
						}
					}
				});
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
		ModalDialog:function(text){
			$('#Panel').find('div').css({
				position:'absolute',
				width:'100%',
				opacity:1,
				background:"#B3C3CD"
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
			YT.ESC=setInterval(function(){
				if(typeof($(document).data('events'))=='undefined'||typeof($(document).data('events')['keypress'])=='undefined'){
					$(document).keypress(function(e){
						if(e.keyCode==27){
							$('#Panel').find('span').trigger('click');	
						}
					});
				}
			},800);
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
	Demo:function(){
		$.get('YT.Ajax.asp', { Action: 'Demo',t:Math.random() },function(txt){
			var t = YT.Panel.ModalDialog(txt);			
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
										YT.InsertText($(t).find('textarea')[0],_t,false);
								};
								$(t).find('li')[1].appendChild(_but);
						}
					};
					$(t).find('li')[1].appendChild(__s);
			}
		};
		$(t).find('li')[1].appendChild(_s);
		if(isAlipay){
			var __but = document.createElement('input');
				__but.type = 'button';
				__but.value = 'ALIPAY';
				__but.onclick = function(){
					$.ajax({
						url: ZC_BLOG_HOST+'ZB_USERS/PLUGIN/YTAlipay/form.html',
						type: 'GET',
						dataType: 'html',
						success: function(s) {
							YT.InsertText($(t).find('textarea')[0],s,false);
						}
					});
				};
				$(t).find('li')[3].appendChild(__but);
		}
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