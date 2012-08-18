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
	Block:'BLOCK.XML',
	Model:'MODEL.XML',
	Table:['YT_Alipay']
}