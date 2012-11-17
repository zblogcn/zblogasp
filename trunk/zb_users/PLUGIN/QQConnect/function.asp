<!-- #include file="functions/microblog.asp"-->
<!-- #include file="functions/qqconnect.asp"-->
<!-- #include file="functions/network.asp"-->
<!-- #include file="functions/database.asp"-->
<!-- #include file="functions/hmac_sha1.asp"-->

<script language="javascript" runat="server">
var qqconnect={}
function init_qqconnect(){
	if(qqconnect["init"]) return ""
	qqconnect["init"]=true;
	qqconnect["tconfig"]=newClass("TConfig");
	qqconnect.tconfig.Load("QQConnect");
	if(qqconnect.tconfig.exists("version")==false) qqconnect.functions.checkconfig();
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
	qqconnect["d"]=newClass("qqconnect_db")
	//qqconnect.c.debug=true;
	//qqconnect.t.debug=true;
}

qqconnect["functions"]={
	"encodeurl":function(url){
		var s=url;
		s=Server.URLEncode(s);
		s=s.replace(/%5F/g,"_").replace(/%2E/g,".").replace(/%2D/g,"-").replace(/%7E/g,"~").replace(/\+/g,"%20");
		return s
	}
	,"getip":function(){
		return (Request.ServerVariables("REMOTE_ADDR").Item=="")?Request.ServerVariables("HTTP_X_FORWARDED_FOR").Item:Request.ServerVariables("REMOTE_ADDR").Item
	}
	,"navbar":function(id){
		var json={
			name:["main.asp","setting.asp","m.asp"]
			,cls:["m-left","m-left","m-left"]
			,text:["首页","设置","绑定"]
			,level:[5,4,4]
			};
		var str="";
		for(var i=0;i<json.name.length;i++){
			if(BlogUser.Level<=json.level[i]){
			str+=MakeSubMenu(json.text[i],json.name[i],json.cls[i]+(id==i?" m-now ":""),false)}
		}
		return str
	}
	,"checkconfig":function(){
		qqconnect.tconfig.Write("version","1.0");
		for(var i=97;i<=105;i++){
			qqconnect.tconfig.Write(String.fromCharCode(i),(String.fromCharCode(i)!="g"?true:false))
		}
		qqconnect.tconfig.Write("a1","0");
		qqconnect.tconfig.Write("content","更新了文章：《%t》，%u");
		qqconnect.tconfig.Write("pl","@%a 评论 %c");
		qqconnect.tconfig.Save();
	}
	,"formatstring":function(data){
		var s=data;
		s=TransferHTML(UBBCode(s,"[link][email][font][code][face][image][flash][typeset][media]"),"[nohtml]");
		s=this.toHTML(s)
		return s
	}
	,"json":{
		"e":function (a,b){return this.toObject(this.toJSONwithEncode(this.extend({}, [a,b])));}
		,"extend":function (des, src, override){if(src instanceof Array){for(var i = 0, len = src.length; i < len; i++)this.extend(des, src[i], override);}  for( var i in src){if(override || !(i in des)){des[i] = src[i];}} delete des["0"];delete des["1"];return des;	}
		,"toObject":function(json) {var o = eval('('+json+')');return o;}
		,"addObj":function(o,attr,str){o[attr] = str;}
		,"delObj":function(o,attr) {delete o[attr];}
		,"toJSON":function(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + String(o[attr]).replace(/(,)/g,"x@._a") + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
		,"toJSONwithEncode":function(o){var json = "";for(attr in o) {json = json == "" ?  "'" + attr + "':'"+ String(o[attr]).replace(/(,)/g,"x@._a") + "'" : json + ",'" + attr + "':'" + qqconnect.functions.encodeurl(String(o[attr]).replace(/(,)/g,"x@._a")) + "'";}json = "{" + json + "}";json = "{" + json.match(/[^,\{]+(?=\}|,)/g).sort().join(",") + "}";json = json.replace(/(x@._a)/g,",");return  json ;}
		,"toStr":function(o){var json = this.toJSON(o);var o = this.toObject(json);var str = "";for(attr in o) {str = str == "" ?  attr + "="+ o[attr] : str + "&" + attr + "="+ o[attr];}return str;}
		,"toObject2":function(o) {var str = "";for(attr in o) {str = str == "" ? "name:'" + attr + "',value:'"+ o[attr] + "'" : str + "},{name:'" + attr + "',value:'" + o[attr] + "'";}str = "[{" + str + "}]";return this.toObject(str);}
	}
	,"toHTML":function(data){
		var a,b,d
		d=data;
		a=new Array("lt","gt","amp","quot","copy","reg","times","divide","nbsp","yen","ordf","macr","acute","sup1","frac34","atilde","egrave","iacute","ograve","times","uuml","aacute","aelig","euml","eth","otilde","uacute","yuml","iexcl","brvbar","laquo","deg","micro","ordm","iquest","auml","eacute","icirc","oacute","oslash","yacute","acirc","ccedil","igrave","ntilde","ouml","ucirc","curren","copy","reg","sup3","cedil","frac12","acirc","ccedil","igrave","ntilde","ouml","ucirc","agrave","aring","ecirc","iuml","ocirc","ugrave","thorn","cent","sect","not","plusmn","para","raquo","agrave","aring","ecirc","iuml","ocirc","ugrave","thorn","atilde","egrave","iacute","ograve","divide","uuml","pound","uml","shy","sup2","middot","frac14","aacute","aelig","euml","eth","otilde","uacute","szlig","auml","eacute","icirc","oacute","oslash","yacute","oelig","oelig","tilde","zwj","lsquo","bdquo","rsaquo","scaron","lrm","rsquo","dagger","euro","scaron","rlm","sbquo","dagger","yuml","thinsp","ndash","ldquo","permil","circ","zwnj","mdash","rdquo","lsaquo","fnof","epsilon","kappa","omicron","upsilon","alpha","zeta","lambda","pi","upsilon","thetasym","prime","image","uarr","larr","forall","isin","minus","ang","int","ne","sup","otimes","lfloor","spades","alpha","zeta","lambda","pi","phi","beta","eta","mu","rho","phi","upsih","prime","real","rarr","uarr","part","notin","lowast","and","there4","equiv","nsub","perp","rfloor","clubs","beta","eta","mu","rho","chi","gamma","theta","nu","sigmaf","chi","piv","oline","trade","darr","rarr","exist","ni","radic","or","sim","le","sube","sdot","lang","hearts","gamma","theta","nu","sigma","psi","delta","iota","xi","sigma","psi","bull","frasl","alefsym","harr","darr","empty","prod","prop","cap","cong","ge","supe","lceil","rang","diams","delta","iota","xi","tau","omega","epsilon","kappa","omicron","tau","omega","hellip","weierp","larr","crarr","harr","nabla","sum","infin","cup","asymp","sub","oplus","rceil","loz");
		b=new Array("<",">","&","\"","©","®","×","÷","\r\n","¥","ª","¯","´","¹","¾","Ã","È","Í","Ò","×","Ü","á","æ","ë","ð","õ","ú","ÿ","¡","¦","«","°","µ","º","¿","Ä","É","Î","Ó","Ø","Ý","â","ç","ì","ñ","ö","û","¤","©","®","³","¸","½","Â","Ç","Ì","Ñ","Ö","Û","à","å","ê","ï","ô","ù","þ","¢","§","¬","±","¶","»","À","Å","Ê","Ï","Ô","Ù","Þ","ã","è","í","ò","÷","ü","£","¨","\r\n","²","·","¼","Á","Æ","Ë","Ð","Õ","Ú","ß","ä","é","î","ó","ø","ý","Œ","œ","˜","‍","‘","„","›","Š","‎","’","†","€","š","‏","‚","‡","Ÿ"," ","–","“","‰","ˆ"," ","—","”","‹","ƒ","Ε","Κ","Ο","Υ","α","ζ","λ","π","υ","?","′","ℑ","↑","⇐","∀","∈","−","∠","∫","≠","⊃","⊗","?","♠","Α","Ζ","Λ","Π","Φ","β","η","μ","ρ","φ","?","″","ℜ","→","⇑","∂","∉","∗","∧","∴","≡","⊄","⊥","?","♣","Β","Η","Μ","Ρ","Χ","γ","θ","ν","ς","χ","?","‾","™","↓","⇒","∃","∋","√","∨","∼","≤","⊆","⋅","?","♥","Γ","Θ","Ν","Σ","Ψ","δ","ι","ξ","σ","ψ","•","⁄","ℵ","↔","⇓","∅","∏","∝","∩","∝","≥","⊇","?","?","♦","Δ","Ι","Ξ","Τ","Ω","ε","κ","ο","τ","ω","…","℘","←","↵","⇔","∇","∑","∞","∪","≈","⊂","⊕","?","◊")
		for(var c=0;c<=a.length;c++){
			d=d.replace("&"+a[c]+";",b[c])
		}
		b=d.match(/&#(\d+?);/g);
		var e;
		if(b!=null){
			for(c=0;c<=b.length;c++){
				e = b[c];
				if(e==null) break
				e = e.substr(2,e.length-3)
				if(e-65536>0) e=e-65536
				d = d.replace(b[c], String.fromCharCode(e))
			}
		}
	
	return d
	}
	,"getpicture":function(s){
		var temp="";
		var r=/<img.*src\s*=\s*[\""|\']?\s*([^>\""\'\s]*)/i;
		if(r.test(s)) temp=r.exec(s)[1]
		if(temp.indexOf("http")<0&&temp!=""){temp=BlogHost + temp}
		return temp
	}
	,"savereg":function(uid,openid,accesstoken){
		if(!(typeof(openid)==undefined)){
			init_qqconnect();
			qqconnect.d.OpenID=openid;
			qqconnect.d.AccessToken=accesstoken;
			qqconnect.d.objUser.LoadInfoById(uid);
			qqconnect.d.Bind();
			return true;
		}
		return false;
	}
	,"getstate":function(){
		return "zsxsoft_"+MD5("zsxsoft_"+this.getip()+Request.ServerVariables("HTTP_USER_AGENT")).substr(0,6).toLowerCase();
	}
}


</script>