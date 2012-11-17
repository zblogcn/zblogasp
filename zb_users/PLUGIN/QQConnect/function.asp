<!-- #include file="functions/microblog.asp"-->
<!-- #include file="functions/qqconnect.asp"-->
<!-- #include file="functions/network.asp"-->
<!-- #include file="functions/hmac_sha1.asp"-->
<%
'我擦，还是得用VBS写。
'VBS丫调用不了JS的伪类啊戳



%>
<script language="javascript" runat="server">
var qqconnect={}
function init_qqconnect(){
	qqconnect["tconfig"]=newClass("TConfig");
	qqconnect.tconfig.Load("QQConnect");
	if(qqconnect.tconfig.exists("version")==false) checkconfig_qqconnect();
	qqconnect["config"]={
		"weibo":{
			"appkey":"2e21c7b056f341b080d4d3691f3d50fb"
			,"appsecret":"1b84a3016c132a6839d082605b854bbe"
			,"token":qqconnect.tconfig.Read("Weibo_Token")
			,"secret":qqconnect.tconfig.Read("Weibo_Secret")
		}
		,"qqconnect":{
			"appid":qqconnect.tconfig.Read("appid")
			,"appsecret":qqconnect.tconfig.Read("key")
			,"openid":""
			,"accesstoken":""
			,"admin":{
				"openid":qqconnect.tconfig.Read("Connect_OpenID")
				,"accesstoken":qqconnect.tconfig.Read("Connect_AccessToken")
			}
		}
	}

	
	qqconnect["t"]=newClass("qqconnect_weibo");
	qqconnect.t.callbackurl=BlogHost+"zb_users/plugin/qqconnect/main.asp?act=callback&type=weibo";
	qqconnect["c"]=newClass("qqconnect_connect");
	qqconnect.c.callbackurl=BlogHost+"zb_users/plugin/qqconnect/main.asp?act=callback&type=connect"
	qqconnect["n"]=newClass("qqconnect_network");
	//qqconnect.c.debug=true;
	//qqconnect.t.debug=true;
}
function qqconnect_encodeurl(url){
	var s=url;
	s=Server.URLEncode(s);
	s=s.replace(/%5F/g,"_").replace(/%2E/g,".").replace(/%2D/g,"-").replace(/%7E/g,"~").replace(/\+/g,"%20");
	return s
}
function qqconnect_getip(){
	return (Request.ServerVariables("REMOTE_ADDR").Item=="")?Request.ServerVariables("HTTP_X_FORWARDED_FOR").Item:Request.ServerVariables("REMOTE_ADDR").Item

}
function qqconnect_navbar(id){
		var json={
			name:["main.asp","setting.asp"]
			,cls:["m-left","m-left"]
			,text:["首页","设置"]
			,level:[5,4]
			};
		var str="";
		for(var i=0;i<json.name.length;i++){
			if(BlogUser.Level<=json.level[i]){
			str+=MakeSubMenu(json.text[i],json.name[i],json.cls[i]+(id==i?" m-now ":""),false)}
		}
		return str
}

	
	
function checkconfig_qqconnect(){
		qqconnect.tconfig.Write("version","1.0");
		for(var i=97;i<=105;i++){
			qqconnect.tconfig.Write(String.fromCharCode(i),(String.fromCharCode(i)!="g"?true:false))
		}
		qqconnect.tconfig.Write("a1","0");
		qqconnect.tconfig.Write("content","更新了文章：《%t》，%u");
		qqconnect.tconfig.Write("pl","@%a 评论 %c");
		qqconnect.tconfig.Save();
}

//没法子，只好这样
function qqconnect_json(){return "请使用不同对象，嗯。"}
qqconnect_json.e=function (a,b){return this.toObject(this.toJSONwithEncode(this.extend({}, [a,b])));}
qqconnect_json.extend=function (des, src, override){if(src instanceof Array){for(var i = 0, len = src.length; i < len; i++)this.extend(des, src[i], override);}  for( var i in src){if(override || !(i in des)){des[i] = src[i];}} delete des["0"];delete des["1"];return des;	}
qqconnect_json.toObject=function(json) {var o = eval('('+json+')');return o;}
qqconnect_json.addObj=function(o,attr,str){o[attr] = str;}
qqconnect_json.delObj=function(o,attr) {delete o[attr];}
qqconnect_json.toJSON=function(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + String(o[attr]).replace(/(,)/g,"x@._a") + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
qqconnect_json.toJSONwithEncode=function(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + qqconnect_encodeurl(String(o[attr]).replace(/(,)/g,"x@._a")) + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
qqconnect_json.toStr=function(o){var json = this.toJSON(o);var o = this.toObject(json);var str = "";for(attr in o) {str = str == "" ?  attr + "="+ o[attr] : str + "&" + attr + "="+ o[attr];}return str;}
qqconnect_json.toObject2=function(o) {var str = "";for(attr in o) {str = str == "" ? "name:'" + attr + "',value:'"+ o[attr] + "'" : str + "},{name:'" + attr + "',value:'" + o[attr] + "'";}str = "[{" + str + "}]";return this.toObject(str);}
</script>