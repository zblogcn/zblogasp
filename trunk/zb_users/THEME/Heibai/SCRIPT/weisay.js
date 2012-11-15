//顶部导航下拉菜单
jQuery(document).ready(function(){
jQuery(".topnav ul li").hover(function(){
	jQuery(this).children("ul").show();
	jQuery(this).addClass("li01");
	},function(){
	jQuery(this).children("ul").hide();
	jQuery(this).removeClass("li01");
});
});

//侧边栏TAB效果
jQuery(document).ready(function(){
jQuery('#tab-title span').click(function(){
	jQuery(this).addClass("selected").siblings().removeClass();
	jQuery("#tab-content > ul").slideUp('1500').eq(jQuery('#tab-title span').index(this)).slideDown('1500');
});
});

//图片渐隐
jQuery(function () {
jQuery('.thumbnail img,.thumbnail_a img,.thumbnail_t img,.thumbnail_b img').hover(
function() {jQuery(this).fadeTo("fast", 0.5);},
function() {jQuery(this).fadeTo("fast", 1);
});
});

//新窗口打开
jQuery(document).ready(function(){
	jQuery("a[rel='external'],a[rel='external nofollow']").click(
	function(){window.open(this.href);return false})
});

//顶部微博等图标渐隐
jQuery(document).ready(function(jQuery){
			jQuery('.icon1,.icon2,.icon3,.icon4').wrapInner('<span class="hover"></span>').css('textIndent','0').each(function () {
				jQuery('span.hover').css('opacity', 0).hover(function () {
					jQuery(this).stop().fadeTo(350, 1);
				}, function () {
					jQuery(this).stop().fadeTo(350, 0);
				});
			});
});