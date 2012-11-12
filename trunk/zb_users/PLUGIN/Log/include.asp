<!-- #include file="config.asp"-->
<%
Call RegisterPlugin("Log","ActivePlugin_Log")

Function ActivePlugin_Log()
	Call Add_Action_Plugin("Action_Plugin_System_Initialize","LogWithoutInitialize")
	Call Add_Action_Plugin("Action_Plugin_Default_Begin","If Log_Default Then System_Initialize:LogWithInitialize")
End Function



%>
<script language="javascript" runat="server">
function LogUa(s){
	var json={"type":"browser","data":""};
	s=s.toLowerCase().toString();
	//Robots
	if(s.indexOf("baiduspider")>0){json["type"]="spider";json["data"]="baidu"} //百度
	else if(s.indexOf("googlebot")>0){json["type"]="spider";json["data"]="google"} //谷歌
	else if(s.indexOf("sogou")>0&&s.indexOf("spider")>0){json["type"]="spider";json["data"]="sogou"}//搜狗
	else if(s.indexOf("sosospider")>0){json["type"]="spider";json["data"]="soso"} //搜搜
	else if(s.indexOf("youdao")>0){json["type"]="spider";json["data"]="youdao"} //有道
	else if(s.indexOf("bingbot")>0){json["type"]="spider";json["data"]="bing"} //必应
	else if(s.indexOf("jikespider")>0){json["type"]="spider";json["data"]="jike"} //即刻
	else if(s.indexOf("360spider")>0){json["type"]="spider";json["data"]="360"} //流氓
	
	//Mobile
	else if(CheckMobile()){
		json["type"]="mobile";
		json["data"]="mobile"
	}
	//Browsers
	
	else if(s.indexOf("lbbrowser")>0){json["data"]="liebao"}   //猎豹
	else if(s.indexOf("maxthon")>0){json["data"]="maxthon"} //傲游
	else if(s.indexOf("theworld")>0){json["data"]="theworld"} //世界之窗
	else if(s.indexOf("tencenttraveler")>0){json["data"]="tt"} //TencentTraveler
	else if(s.indexOf("qqbrowser")>0){json["data"]="qb"} //QQ浏览器
	else if(s.indexOf("bidubrowser")>0){json["data"]="baidu"} //百度浏览器
	else if(s.indexOf("saayaa")>0){json["data"]="saayaa"} //闪游浏览器
	else if(s.indexOf("ylmfbr")>0){json["data"]="114"} //114浏览器
	else if(s.indexOf("lunaspace")>0){json["data"]="lunaspace"} //LunaSpace
	else if(/se .+?metasr/gi.test(s)>0){json["data"]="sogou"}//搜狗浏览器，由于要判断正则所以放后
	
	 /*
	中国使用人数较少的浏览器
	hotbrowser|	mozilla suite|	camino|	3b|	ndsbrowser|	wii internet channle|	konqueror|	greenbrowser|	amaya|	aweb|	arachne|	aol|	aphrodite|	beonex communicator|	camino|	compuserve|	doczilla|	epiphany|	galeon|	ibm web browser|	kmeleon|	kazehakase|	manyone|	minimo|	salamander|	seamonkey|	skipstone|	flock|	activestate komodo|	liferea|	kazehakase|	avant|	flock|	epiphany|	galeon|	iceweasel|	dillo|	lynx|	minefield|	shiretoko|	elinks|*/
					
	else if(s.indexOf("360se")>0||s.indexOf("360ee")>0){json["data"]="360"} //傻B
	
	//谷歌、Safari、火狐、IE必须放最后判断，前面那些套壳浏览器你们好意思么
	else if(s.indexOf("opera")>0){json["data"]="Opera"} //Opera
	else if(s.indexOf("firefox")>0){json["data"]="FireFox"} //谷歌
	else if(s.indexOf("chrome")>0){json["data"]="Chrome"} //谷歌
	else if(s.indexOf("ie")>0){
		if(s.indexOf("ie 6")>0){json["data"]="ie 6"}//比360傻逼好一点，不过也是无可救药，开发人员除外。
		else if(s.indexOf("ie 7")>0){json["data"]="ie 7"}//市场占有率好低
		else if(s.indexOf("ie 8")>0){json["data"]="ie 8"}//求求你了XP快死吧，IE9我来了
		else if(s.indexOf("ie 9")>0){json["data"]="ie 9"}//第一个支持HTML5的IE啊。。
		else if(s.indexOf("ie 10")>0){json["data"]="ie 10"}//微软来吧~
		else{json["data"]="ie"}//这种还用着Win95\98\Me\2000的...
	} //IE	
	else if(s.indexOf("safari")>0){json["data"]="Safari"} //Safari
	return json
}
function LogWithoutInitialize(){
	var s=BlogHost;
	s=s.replace(/\//g,"\\/");
	//Response.Write(s+"default\.asp");
	//Response.End();
	var zsx=NewClass("TCounter");
	var temp=zsx.GetUrl();
	var ua=LogUa(Request.ServerVariables("HTTP_USER_AGENT").item),matchcmd=false;
	ua=ua.type+"-"+ua.data;
	if(temp.indexOf("cmd.asp")>0){
		/*switch(Request.QueryString("act")){
			case "cmt":zsx.Add(ua,"添加评论",false);matchcmd=true;break;
			
		}*/
		var ary="Root|login|verify|logout|admin|cmt|vrs|rss|batch|BlogReBuild|FileReBuild|ArticleMng|ArticleEdt|ArticlePst|ArticleDel|CategoryMng|CategoryPst|CategoryDel|CommentMng|CommentDel|UserMng|UserEdt|UserMod|UserCrt|UserDel|FileMng|FileUpload|FileDel|Search|TagMng|TagEdt|TagPst|TagDel|SettingMng|SettingSav|PlugInMng|FunctionMng".split("|");
		var s=ZC_MSG019.split("%s");

		for(var i=0;i<=ary.length;i++){
			if(Request.QueryString("act")==ary[i]){matchcmd=true;zsx.Add(ua,s[i],false)}
		}
	}
	if(matchcmd==false){
		if(/(zb_users\/(plugin|theme))|zb_system/i.test(temp)){
			zsx.Add(ua,"后台访问",false);
		}
		else if(new RegExp(s+"default\.asp","i").test(temp)){
			zsx.Add(ua,"首页",false);
		}
		else if(new RegExp(s+"catalog\.asp","i").test(temp)){
			zsx.Add(ua,"分类页",false);
		}
		else if(new RegExp(s+"tags\.asp","i").test(temp)){
			zsx.Add(ua,"Tags",false);
		}
		else if(new RegExp(s+"view\.asp","i").test(temp)){
			zsx.Add(ua,"文章页",false);
		}
		else if(new RegExp(s+"search\.asp","i").test(temp)){
			zsx.Add(ua,"搜索",false);
		}
		else{zsx.Add("",ua,false)}
	}
}
</script>