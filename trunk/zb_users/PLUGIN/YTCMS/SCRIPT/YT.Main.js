$(document).ready(function() {
	try{
		var ary = [];
		ary.push('YT.Panel.Block.M()');
		ary.push('YT.Panel.Model.M()');
		ary.push('YT.Panel.TPL.M()');
		ary.push('YT.Panel.SQL()');
		ary.push('YT.Demo()');
		$('.d_mainbox:eq(0)').show();
		$('.d_tab a').each(function(i) {
			$(this).click(function(){
				eval(ary[i]);
				$(this).addClass('d_tab_on').siblings().removeClass('d_tab_on');
				$($('.d_mainbox')[i]).show().siblings('.d_mainbox').hide();
			})
		});
		eval(ary[0]);
		//加载视图
		//YT.Panel.Load();
		//加载插件信息
		YT.Copyright();
		//加载模块创建
		YT.Panel.Block.C();
		//加载模型创建
		YT.Panel.Model.C();
		//加载模板创建
		YT.Panel.TPL.C();
	}catch(e){}
});