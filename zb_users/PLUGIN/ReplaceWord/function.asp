
<script language="javascript" runat="server">
var replaceword={};
replaceword["init"]=function(){
	replaceword["xmldom"]=Server.CreateObject("Microsoft.XMLDOM");
	replaceword.xmldom.load(Server.MapPath("config.asp"));
	replaceword["words"]=undefined;
	if(replaceword.xmldom.readyState==4){
		replaceword.words=replaceword.xmldom.documentElement.selectNodes("word");
	}
	replaceword["orig"]="";
	replaceword["string"]="";
	replaceword["defaultstring"]="**"
	
}

replaceword["user"]=function(id){return replaceword.words[id].attributes.getNamedItem("user").value}
replaceword["regex"]=function(id){
	var m=replaceword.words[id].attributes.getNamedItem("regexp").value;
	return (m=="False"?false:true)
}
replaceword["str"]=function(id){
	return replaceword.words[id].selectSingleNode("str").text
}
replaceword["des"]=function(id){
	return replaceword.words[id].selectSingleNode("description").text
}
replaceword["rep"]=function(id){
	var s2=replaceword.words[id].selectSingleNode("replace").text;
	s2=(s2=="" ? replaceword.defaultstring : s2);
	return s2
}
replaceword["replace"]=function(){
	var str,re;
	str=replaceword.orig;
	var s1,s2;
	for(var id=0; id<=replaceword.words.length-1; id++){
		s2=replaceword.rep(id);
		if(replaceword.regex(id)){
			s1=new RegExp(replaceword.str(id),"ig");
			str=str.replace(s1,s2);
		}else{
			s1=replaceword.str(id);
			str=str.vbsreplace(s1,s2);
		}
	}
	replaceword.string=str;
	return str;
}

replaceword["submenu"]=function(id){
	var json={
		name:["main.asp","import.asp","import.asp?act=export"]
		,cls:["m-left","m-left","m-left"]
		,text:["设置","导入","导出"]
		,level:[1,1,1]
		};
	var str="";
	for(var i=0;i<json.name.length;i++){
		if(BlogUser.Level<=json.level[i]){
		str+=MakeSubMenu(json.text[i],json.name[i],json.cls[i]+(id==i?" m-now ":""),false)}
	}
	return str
}
</script>