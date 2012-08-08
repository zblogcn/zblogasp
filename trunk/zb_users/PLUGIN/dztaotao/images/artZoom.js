/*
微博简约贴图缩放插件

Version:    1.0.2 [2010-05-28]
Author:     tangbin
Email:      1987.tangbin@gmail.com
Website:    www.planeArt.cn
*/
(function($){
	$.fn.artZoom = function(){
		$(this).live('click', function(){
			var maxImg = $(this).attr('href'),
			viewImg = $(this).attr('rel').length === '0' ? $(this).attr('rel') : maxImg; // 如果连接含有rel属性，则新窗口打开的原图地址为此rel里面的地址
			
			if ($(this).find('.loading').length == 0) $(this).append('<span class="loading" title="Loading..">Loading..</span>');
			imgTool($(this), maxImg, viewImg);
			return false;
		});
		
		// 图片预先加载
		var loadImg = function (url, fn){
			var img = new Image();
			img.src = url;
			if (img.complete){
				fn.call(img);
			}else{
				img.onload = function(){
					fn.call(img);
				};
			};
		};
		
		// 图片工具条
		var imgTool = function(on, maxImg, viewImg) {
			var width = 0,
			height = 0,
			tool = function(){
				on.find('.loading').remove();
				on.hide();
				
				if (on.next('.artZoomBox').length != 0){
					return on.next('.artZoomBox').show();
				};
				
				// 等比例限制宽度
				var maxWidth = on.parent().innerWidth() - 12; // 获取父级元素宽度
				if (width > maxWidth) {
					height = maxWidth / width * height;
					width = maxWidth;
				};
				
				var html = '<div class="artZoomBox"><div class="tool"><a class="hideImg" href="#" title="收起">收起</a><a class="imgLeft" href="#" title="向左转">向左转</a><a class="imgRight" href="#" title="向右转">向右转</a><a class="viewImg" href="' + viewImg + '" title="查看原图">查看原图</a></div><a href="' + viewImg + '" class="maxImgLink"> <img class="maxImg" width="' + width + '" height="' + height + '" src="' + maxImg + '" /></a></div>';
				on.after(html);
				var box = on.next('.artZoomBox');
				box.hover(function(){
					box.addClass('js_hover');
				}, function(){
					box.removeClass('js_hover');
				});
				box.find('a').bind('click', function(){
					// 收起
					if($(this).hasClass('hideImg') || $(this).hasClass('maxImgLink')) {
						box.hide();
						box.prev().show();
					};
					// 左旋转
					if($(this).hasClass('imgLeft')) {
						box.find('.maxImg').rotate('left')
					};
					// 右旋转
					if($(this).hasClass('imgRight')) {
						box.find('.maxImg').rotate('right')
					};
					// 新窗口打开
					if($(this).hasClass('viewImg')) window.open(viewImg);
					
					return false;
				});
				
			};
			
			loadImg(maxImg, function(){
				width = this.width;
				height= this.height;
				tool();
			});
			
			// 图片旋转
			// 方案修改自：http://byzuo.com/
			$.fn.rotate = function(p){

				var img = $(this)[0],
					n = img.getAttribute('step');

				// 保存图片大小数据
				if (!this.data('width') && !$(this).data('height')) {
					this.data('width', img.width);
					this.data('height', img.height);
				};

				if(n == null) n = 0;
				if(p == 'left'){
					(n == 3) ? n = 0 : n++;
				}else if(p == 'right'){
					(n == 0)? n = 3 : n--;
				};
				img.setAttribute('step', n);

				// IE浏览器使用滤镜旋转
				if(document.all) {
					img.style.filter = 'progid:DXImageTransform.Microsoft.BasicImage(rotation='+ n +')';
					// IE8高度设置
					if ($.browser.version == 8) {
						switch(n){
							case 0:
								this.parent().height('');
								//this.height(this.data('height'));
								break;
							case 1:
								this.parent().height(this.data('width') + 10);
								//this.height(this.data('width'));
								break;
							case 2:
								this.parent().height('');
								//this.height(this.data('height'));
								break;
							case 3:
								this.parent().height(this.data('width') + 10);
								//this.height(this.data('width'));
								break;
						};
					};
				// 对现代浏览器写入HTML5的元素进行旋转： canvas
				}else{
					var c = this.next('canvas')[0];
					if(this.next('canvas').length == 0){
						this.css({'visibility': 'hidden', 'position': 'absolute'});
						c = document.createElement('canvas');
						c.setAttribute('class', 'maxImg canvas');
						img.parentNode.appendChild(c);
					}
					var canvasContext = c.getContext('2d');
					switch(n) {
						default :
						case 0 :
							c.setAttribute('width', img.width);
							c.setAttribute('height', img.height);
							canvasContext.rotate(0 * Math.PI / 180);
							canvasContext.drawImage(img, 0, 0);
							break;
						case 1 :
							c.setAttribute('width', img.height);
							c.setAttribute('height', img.width);
							canvasContext.rotate(90 * Math.PI / 180);
							canvasContext.drawImage(img, 0, -img.height);
							break;
						case 2 :
							c.setAttribute('width', img.width);
							c.setAttribute('height', img.height);
							canvasContext.rotate(180 * Math.PI / 180);
							canvasContext.drawImage(img, -img.width, -img.height);
							break;
						case 3 :
							c.setAttribute('width', img.height);
							c.setAttribute('height', img.width);
							canvasContext.rotate(270 * Math.PI / 180);
							canvasContext.drawImage(img, -img.width, 0);
							break;
					};
				};
			};
			
		};
	};
	
	// 给所有含'artZoom'类的链接绑定本效果
	$('a.artZoom').artZoom();

})(jQuery);