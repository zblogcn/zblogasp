<%
''*****************************************************
'   ZSXSOFT JSON操作处理类
'   主要功能：（oAuth1）排序、添加删除数据避免&=&=麻烦，直接addobj即可
''*****************************************************
%>
<%
'合并JSON，将oauth基本参数与api参数合并
Function ZBQQConnect_jsonExtendBasic(jsonobj1,ByVal strjson2)
	dim j,k,l,n,m
	set j=	jsonobj1
	set k=ZBQQConnect_json.toobject(strjson2)
	set m=ZBQQConnect_json.e(j,k)
	
	ZBQQConnect_jsonExtendBasic=ZBQQConnect_json.TOJSONwithEncode(m)
End Function
%>

<script language="javascript" runat="server">
/*
JS部分核心组件之一——JSON读取、删除、修改、添加
部分代码来源于网络，作者未知
第一次感觉到js在vbs里的好用，嗯。。
*/
function ZBQQConnect_json(){return "请使用不同对象，嗯。"}
ZBQQConnect_json.e=function (a,b){return this.extend({}, [a,b]);}
ZBQQConnect_json.extend=function (des, src, override){if(src instanceof Array){for(var i = 0, len = src.length; i < len; i++)this.extend(des, src[i], override);}  for( var i in src){if(override || !(i in des)){des[i] = src[i];}} delete des["0"];delete des["1"];return des;	}
ZBQQConnect_json.toObject=function(json) {var o = eval('('+json+')');return o;}
ZBQQConnect_json.addObj=function(o,attr,str){o[attr] = str;}
ZBQQConnect_json.delObj=function(o,attr) {delete o[attr];}
ZBQQConnect_json.isObj=function(o,attr){if(typeof(o[attr])!="undefined") return true;}
ZBQQConnect_json.noObj=function(o,attr){if(typeof(o[attr])=="undefined") return true;}
ZBQQConnect_json.toJSON=function(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + String(o[attr]).replace(/(,)/g,"x@._a") + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
ZBQQConnect_json.toJSONwithEncode=function(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + encodeURIComponent(String(o[attr]).replace(/(,)/g,"x@._a")) + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
ZBQQConnect_json.toStr=function(o){var json = this.toJSON(o);var o = this.toObject(json);var str = "";for(attr in o) {str = str == "" ?  attr + "="+ o[attr] : str + "&" + attr + "="+ o[attr];}return str;}
ZBQQConnect_json.toObject2=function(o) {var str = "";for(attr in o) {str = str == "" ? "name:'" + attr + "',value:'"+ o[attr] + "'" : str + "},{name:'" + attr + "',value:'" + o[attr] + "'";}str = "[{" + str + "}]";return this.toObject(str);}
ZBQQConnect_json.getItem=function(obj,Num,Name){return obj[Num][Name];}
ZBQQConnect_json.viewObject=function(obj){  var msg;for(var a in obj){  msg += ' ' + a;for(var x in obj[a]){msg += ' ' + x;msg += ' ' + obj[a][x];}}return msg;}
</script>
