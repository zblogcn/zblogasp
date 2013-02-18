<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("MenuManage","ActivePlugin_MenuManage")
'挂口部分
Function ActivePlugin_MenuManage()

	MenuManage.init()
	
	Call Add_Action_Plugin("Action_Plugin_System_Initialize_Succeed","MenuManage.left.export()")

End Function


Function MakeLeftMenu(requireLevel,strName,strUrl,strLiId,strAId,strImgUrl)

	'If BlogUser.Level>requireLevel Then Exit Function
	
	Call MenuManage.left.push(strName,strUrl,strLiId,strAId,strImgUrl,requireLevel,False)
	
	'MakeLeftMenu=""

End Function


%>
<script language="javascript" runat="server" >
var MenuManage={};
MenuManage["init"]=function(){
	MenuManage["c"]=newClass("TConfig");
	MenuManage.c.Load("MenuManage");
	if(!MenuManage.c.Exists("major")){
		MenuManage.c.Write("major","1");
		MenuManage.c.Write("default","aArticleEdt|aArticleMng|aPageMng|hr|aCategoryMng|aTagMng|aCommentMng|aFileMng|aUserMng|hr|aThemeMng|aPlugInMng|aFunctionMng");
		MenuManage.c.Write("config","default");	
		MenuManage.c.Save()	
	}
	
};

MenuManage["left"]={
	"cfgl":{},
	"push":function(strName,strUrl,strLiId,strAId,strImgUrl,requireLevel,custom){
		
		MenuManage.left.cfgl[strAId]=function(){
			var sdata="";
			sdata+="{'custom':"+custom+",";
			sdata+="'liid':'"+strLiId.replace(/'/g,"\\'")+"',";
			sdata+="'name':'"+strName.replace(/'/g,"\\'")+"',";
			sdata+="'aid':'"+strAId.replace(/'/g,"\\'")+"',";
			sdata+="'url':'"+strUrl.replace(/'/g,"\\'")+"',";
			sdata+="'img':'"+strImgUrl.replace(/'/g,"\\'")+"',"
			sdata+="'level':"+requireLevel.toString().replace(/'/g,"\\'")+"}";
			return sdata
		}()
		
		

	}
	,
	"export":function(){
		
		for(var j in MenuManage.left.cfgl){
			MenuManage.c.Write(j,MenuManage.left.cfgl[j]);
		}
		if(MenuManage.c.Read("config")=="default"){MenuManage.c.Write("config",MenuManage.c.Read("default"))}
		MenuManage.c.Save();
		
		var s=MenuManage.c.Read("config").split("|"),d="";
		for(j in s){
			if(MenuManage.left.cfgl[s[j]]!=null){
				var k=eval('('+MenuManage.left.cfgl[s[j]]+')');
				if(BlogUser.Level<=parseInt(k.level)){
					d+=function(strLiId,strAId,strUrl,strImgUrl,strName){
						if(strAId!="hr"){
							var s="<li id=\""+strLiId+"\"><a id=\""+strAId+"\" href=\""+strUrl+"\"><span";
							if(strImgUrl!=""){
								s+=" style=\"background-image:url('"+strImgUrl+"')\""
							}
							s+=">"+strName+"</a></li>"
						}
						else{
							var s="<li class='split'><hr/></li>"
						}
						return s;
					}(k.liid,k.aid,k.url,k.img,k.name)
				}
			}
		}
		Response_Plugin_Admin_Left=d;
		//var s=newClass()
	}
}
</script>