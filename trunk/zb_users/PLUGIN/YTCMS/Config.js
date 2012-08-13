function jsonToObject(json){
	return eval('('+json+')');	
}
function jsEscape(Text){
	return escape(Text);
}
function jsUnEscape(Text){
	return unescape(Text);
}
var YTConfig = {
	Global:{
		Url:[
			'HTTP://SHOP.PUGUWANG.COM/',
			'HTTP://SHOP.PUGUWANG.COM/NOTICE/',
			'HTTP://SHOP.PUGUWANG.COM/SHOP/',
			'HTTP://SHOP.PUGUWANG.COM/HELP/'
		]
	},
	TPL:'TPL.XML',
	Block:'BLOCK.XML',
	Model:'MODEL.XML',
	Table:['YT_Alipay'],
	Multi:[],
	Single:[]
}