KindEditor.plugin('zbimages', function(K) {
        var editor = this, name = 'zbimages';
        // 点击图标时执行
        editor.clickToolbar(name, function() {
                //editor.insertHtml('你好');
				this.callbacks = function(obj,win){
				this.value = '';
				for(key in obj) {
					this.value += ueconfig.imagePath.replace('{$ZC_BLOG_HOST}','') + obj[key].url;
				}
				win.close();
			};
			window.showModalDialog('image.html',this,'dialogWidth:635px;dialogHeight:390px;resizable:no;scroll:no;status:no;');
			
			
			var dialog = K.dialog({
					width : 500,
					title : '测试窗口',
					body : '<div style="margin:10px;"><strong>内容</strong></div>',
					closeBtn : {
							name : '关闭',
							click : function(e) {
									dialog.remove();
							}
					},
					yesBtn : {
							name : '确定',
							click : function(e) {
									alert(this.value);
							}
					},
					noBtn : {
							name : '取消',
							click : function(e) {
									dialog.remove();
							}
					}
			});			
			
			
			
			
			
			
			
			
			
			
			
			
			
        });
});


//	$('.upload').each(function(){
//		$(this).click(function(){
//			this.callbacks = function(obj,win){
//				this.value = '';
//				for(key in obj) {
//					this.value += ueconfig.imagePath.replace('{$ZC_BLOG_HOST}','') + obj[key].url;
//				}
//				win.close();
//			};
//			window.showModalDialog('{$ZC_BLOG_HOST}zb_users/plugin/ytcms/image.html',this,'dialogWidth:635px;dialogHeight:390px;resizable:no;scroll:no;status:no;');
//		});
//	});