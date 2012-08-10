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
	set k=ZBQQConnect_toobject(strjson2)
	set m=ZBQQConnect_JSONE(j,k)
	ZBQQConnect_jsonExtendBasic=ZBQQConnect_TOJSONwithEncode(m)
End Function
%>
<script language="javascript" type="text/javascript" runat="server">
/*
JS部分核心组件之一——JSON读取、删除、修改、添加
部分代码来源于网络，作者未知
*/
function ZBQQConnect_JSONE(a,b){return ZBQQConnect_JSONExtend({}, [a,b]);}
function ZBQQConnect_JSONExtend(des, src, override){if(src instanceof Array){for(var i = 0, len = src.length; i < len; i++)ZBQQConnect_JSONExtend(des, src[i], override);}  for( var i in src){if(override || !(i in des)){des[i] = src[i];}} delete des["0"];delete des["1"];return des;	}
function ZBQQConnect_toObject(json) {var o = eval('('+json+')');return o;}
function ZBQQConnect_addObj(o,attr,str){o[attr] = str;}
function ZBQQConnect_delObj(o,attr) {delete o[attr];}
function ZBQQConnect_isObj(o,attr){if(typeof(o[attr])!="undefined") return true;}
function ZBQQConnect_noObj(o,attr){if(typeof(o[attr])=="undefined") return true;}
function ZBQQConnect_toJSON(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + String(o[attr]).replace(/(,)/g,"x@._a") + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
function ZBQQConnect_toJSONwithEncode(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + encodeURIComponent(String(o[attr]).replace(/(,)/g,"x@._a")) + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
function ZBQQConnect_toStr(o){var json = ZBQQConnect_toJSON(o);var o = ZBQQConnect_toObject(json);var str = "";for(attr in o) {str = str == "" ?  attr + "="+ o[attr] : str + "&" + attr + "="+ o[attr];}return str;}
function ZBQQConnect_toObject2(o) {var str = "";for(attr in o) {str = str == "" ? "name:'" + attr + "',value:'"+ o[attr] + "'" : str + "},{name:'" + attr + "',value:'" + o[attr] + "'";}str = "[{" + str + "}]";return ZBQQConnect_toObject(str);}
function ZBQQConnect_getItem(obj,Num,Name){return obj[Num][Name];}
function ZBQQConnect_viewObject(obj){  var msg;for(var a in obj){  msg += ' ' + a;for(var x in obj[a]){msg += ' ' + x;msg += ' ' + obj[a][x];}}return msg;}
</script>