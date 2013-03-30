/**
 *	 Fool.js (by @idiot)
 *	 I pity the user.
 */
 
(function($) {

	$.fn.prefixCSS = function(prop, val) {
		var prefixes = ['Webkit', 'Moz', 'Ms', 'O'],
			obj = {};
		
		for(var i in prefixes) {
			obj[prefixes[i] + (prop.charAt(0).toUpperCase() + prop.substr(1))] = val;
		}
		
		return this.css(obj);
	};
	
	//  Don't be a dick.
	$.fool = function(options) {
		var iframe = '<embed src="http://player.youku.com/player.php/sid/*/v.swf" width="700" height="550" type="application/x-shockwave-flash" wmode="opaque" flashvars="isAutoPlay=true" allowscriptaccess="sameDomain" quality="high" style="position:absolute;left:-999em;top:-999em;-webkit-user-select:none;-webkit-user-drag:none;" frameborder="0" allowfullscreen></embed>',
		
			//  Our good king, Rick Astley
			rick = 'oHg5SJYRHA0',
			
			//  A list of the annoying videos
			videos = ["XNDYwOTM2MDQw","XMjMzNjc1OTc2","XNDI3OTkyMTI4","XMTM4MzIzNDIw","XNDIwMjA5NjYw","XNTExNjgyNTk2","XMzQ2NzA3MTE2","XNzcyNzc1NzY=","XMjQyOTU1NDgw","XMTE3MDk5Mzk2","XNTMwMzUwMDY0"],
			
			//最炫民族风、初音圆周率、月亮之上、江南style、忐忑、爱的供养、猪八戒背媳妇、敢问路在何方、爱我中华、在希望的田野上、金箍棒
			//没记错我放了这几首
			
			//  Baby, let's make our move
			moves = {
			
				//  Show a random youtube video and hide it
				hiddenVideos: function(url) {
					//  Grab a random video
					var video = url ? url : videos[Math.round((Math.random() * (videos.length - 1)))];
					
					return this.append(iframe.replace('*', video));
				},
				
				//  I've dropped a lot of pranks, but I'm never going to give you up.
				rick: function() {
					return moves.hiddenVideos.call(body, rick);
				},
				
				//  Hide random elements on hover
				vanishingElements: function() {
					return $('h1,h2,h3,p,div:not(.timber),input,header,footer,section').hover(function() {
						if(Math.random() > .75) {
							$(this).css('opacity', $(this).css('opacity') == 0 ? 1 : 0);
						}
					});
				},
				
				fallingScrollbar: function() {
				
					$('.timber').fadeOut(200, function() {
						$(this).remove();
					});
				
					var h = $(window).height() + 30,
						html = '<div class="timber" style="-webkit-transform-origin:50% 100%;-moz-transform-origin:50% 100%;-ms-transform-origin:50% 100%;-o-transform-origin:50% 100%;transform-origin:50% 100%;-webkit-transition:-webkit-transform .8s;-moz-transition:-moz-transform .8s;-ms-transition:-ms-transform .8s;-o-transition:-o-transform .8s;transition:transform .8s;position:fixed;right:0;bottom:0;overflow:scroll;width:14px;height:' + h + 'px">' + new Array(80).join('<br>') + '</div>',
						
						me = this.css('overflow', 'hidden').append(html),
						rot = 'rotate(-100deg)';
					
					setTimeout(function() {
						me.children('.timber').prefixCSS('transition', '.8s').css({
							right: -23,
							bottom: 7,
						});
					}, 250);
				},
				
				questionTime: function() {
					var q = ['准备好了吗，孩子们？', '太小声咯。', '哦哦哦~~~，是谁住在深海的大凤梨里？', '方方黄黄伸缩自如！', '如果四处探险是你的愿望！', '那就敲敲甲板让大鱼开路！'],
						a = '海绵宝宝！海绵宝宝！海绵宝宝！';
						
					for(var i in q) {
						window.prompt(q[i],"海绵宝宝！");
					}
					
					for(var t = 0; t < 4; t++) {
						alert(a);
					}
					
					alert("哈哈哈！海绵~~~宝宝！！！")
				},
				
				//  I can hack a site!
				h4xx0r: function() {
					this[0].contentEditable = true;
					return document.designMode = 'on';
				},
				
				//  This probably won't work with the falling scrollbar.
				upsideDown: function() {
					body.attr('style', '-webkit-transform: rotate(10deg); -moz-transform: rotate(10deg); -ms-transform: rotate(10deg); -o-transform: rotate(10deg); transform: rotate(10deg); filter: progid:DXImageTransform.Microsoft.Matrix(M11=-1, M12=-1.2246063538223773e-16, M21=1.2246063538223773e-16, M22=-1, sizingMethod=\'auto expand\'); zoom: 1;');
				},
				
				//  A bit crooked, but a lot of fun
				wonky: function() {
					body.attr('style', '-webkit-transform: rotate(-.7deg);-moz-transform: rotate(-.7deg); -ms-transform:rotate(-.7deg); -o-transform:rotate(-.7deg); transform: rotate(-.7deg); filter: progid:DXImageTransform.Microsoft.Matrix(M11=0.999925369660452,M12=0.012217000835247169,M21=-0.012217000835247169,M22=0.999925369660452,sizingMethod=\'auto expand\'); zoom: 1;');
				},
				
				//  flashes the screen on and off
				flash: function() {
					var fade = function() {
							body.delay(250).animate({opacity: 0}, 1).delay(250).animate({opacity: 1}, 1, fade);
						};
						
					fade();
				},
				
				//  might not work
				crashAndBurn: function() {
					for(i = 0; i <= 0e10; i++) {
						$.fool('crashAndBurn');
					}
				},
				
				//  make a shutter descend unto the screen
				shutter: function() {
					var shutter = body.append('<div id="shutter" />').children('#shutter');
					
					shutter.css({
						position: 'fixed',
						left: 0,
						top: 0,
						right: 0,
						bottom: '100%',
						
						background: '#000'
					}).animate({
						bottom: 99999
					}, 100000);
				},
				
				//  simple
				unclickable: function() {
					//  twitter.com/#!/idiot/status/180261881460690945
					body.attr('style', 'pointer-events: none; -webkit-user-select: none; -moz-user-select: none; cursor: wait;');
				}
			},
			
			body = $('body');
		
		//  Check we've got options	
		if(options) {
			//  Are we calling multiple options
			if(typeof options == 'object') {
				for(i in options) {
					if(options[i] != false && moves[i]) {
						moves[i].call(body);
					}
				}
			} else {
				//  Assume string
				if(moves[options]) {
					moves[options].call(body);
				}
			}
		} else { //  If not, call in Mr. Astley
			return moves['rick'].call(body);
		}
	};

})(jQuery);

