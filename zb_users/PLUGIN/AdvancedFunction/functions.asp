<script language="javascript" runat="server">
var advancedfunction={
	
	config:{}
	,cls:{}
	,init:function(){
		this.cls["config"]=newclass("TConfig");
		this.cls.config.Load("AdvancedFunction");
		if(this.cls.config.Exists("版本")==false){
			this.cls.config.Write("版本","1.0");
			this.cls.config.Write("最新文章","0");
			this.cls.config.Write("访问最多文章","0");
			this.cls.config.Write("本月最热文章","0");
			this.cls.config.Write("本年最热文章","0");
			this.cls.config.Write("分类最热文章","0");
			this.cls.config.Write("评论最多文章","0");
			this.cls.config.Write("本月评论最多","0");
			this.cls.config.Write("本年评论最多","0");
			this.cls.config.Write("分类评论最多","0");
			this.cls.config.Write("随机文章","0");
			this.cls.config.Save();
		}
			this.config["版本"]=this.cls.config.Read("版本");
			this.config["访问最多文章"]=this.cls.config.Read("访问最多文章");
			this.config["本月最热文章"]=this.cls.config.Read("本月最热文章");
			this.config["本年最热文章"]=this.cls.config.Read("本年最热文章");
			this.config["分类最热文章"]=this.cls.config.Read("分类最热文章");
			
			this.config["评论最多文章"]=this.cls.config.Read("评论最多文章");
			this.config["本月评论最多"]=this.cls.config.Read("本月评论最多");
			this.config["本年评论最多"]=this.cls.config.Read("本年评论最多");
			this.config["分类评论最多"]=this.cls.config.Read("分类评论最多");
			
			this.config["随机文章"]=this.cls.config.Read("随机文章");
	}
	,functions:{
		readconfig:function(s){var m;eval("m=advancedfunction.config."+s);return m;}
		,savefunction:function(id,name,htmlid,content){
			var objfunc;
			GetFunction();
			//Response.Write(htmlid+"|"+(FunctionMetas.GetValue(id.toLowerCase())));
			//return false
			if(FunctionMetas.GetValue(id)==jempty){
				objfunc=newClass("TFunction");
				objfunc.ID=0;
				objfunc.Name=name;
				objfunc.FileName=id;
				objfunc.HtmlID=htmlid;
				objfunc.Ftype="ul";
				objfunc.Order=20;
				objfunc.MaxLi=0;
				objfunc.SidebarID=10000;
				objfunc.isSystem=false;
				}
			else{
				objfunc=Functions(FunctionMetas.GetValue(id))
			}
			objfunc.Content=content;
			return objfunc.Save();
		}
		,makeview:function(sql,Max){
			var subtemplate=new Array(Max);
			var strsql;
			strsql="SELECT TOP "+Max+" [log_ID],[log_CateID],[log_Title],[log_Content],[log_Level],[log_AuthorID],";
			strsql+="[log_PostTime],[log_Url],[log_FullUrl],[log_Type],[log_ViewNums] FROM [blog_Article] WHERE [log_Level]=4 AND [log_Type]=0 ";
			strsql+=(sql==""?"":"AND "+sql);
			strsql+=" ORDER BY [log_ViewNums] DESC";
			var s=NewClass("TArticle");
			var template="<li><a href=\"$url$\" title=\"$title$\">$title_sort$($viewnums$)</a></li>"
			var objrs=objconn.Execute(strsql)
			for(var i=0;i<=Max;i++){
				if(objrs.EOF){break;}
				var time=jsTimetovbs_vbs(objrs("log_PostTime"));
				s.loadinfobyarray(jsarraytovbs_js(new Array(objrs("log_ID"),"",objrs("log_CateID"),objrs("log_Title"),"","",objrs("log_Level"),objrs("log_AuthorID"),time,"",objrs("log_ViewNums"),"",objrs("log_Url"),"","",objrs("log_FullUrl"),objrs("log_Type"),"")));
				subtemplate[i]=template.replace("$url$",s.fullurl);
				subtemplate[i]=subtemplate[i].replace("$title$",s.title);
				subtemplate[i]=subtemplate[i].replace("$title_sort$",s.title.substr(0,20));
				subtemplate[i]=subtemplate[i].replace("$viewnums$",s.viewnums);
				objrs.MoveNext;
			}
			if(i<=0){return ""};
			return subtemplate.join("")
		}
		,makecomm:function(sql,Max){
			
			var subtemplate=new Array(Max);
			var strsql;
			strsql="SELECT TOP "+Max+" [log_ID],[log_CateID],[log_Title],[log_Content],[log_Level],[log_AuthorID],";
			strsql+="[log_PostTime],[log_Url],[log_FullUrl],[log_Type],[log_CommNums] FROM [blog_Article] WHERE [log_Level]=4 AND [log_Type]=0 ";
			strsql+=(sql==""?"":"AND "+sql);
			strsql+=" ORDER BY [log_CommNums] DESC";

			var template="<li><a href=\"$url$\" title=\"$title$\">$title_sort$($commnums$)</a></li>"
			var s=NewClass("TArticle");
			var objrs=objconn.Execute(strsql);
			for(var i=0;i<=Max;i++){
				if(objrs.EOF){break;}
				var time=jsTimetovbs_vbs(objrs("log_PostTime"));
				s.loadinfobyarray(jsarraytovbs_js(new Array(objrs("log_ID"),"",objrs("log_CateID"),objrs("log_Title"),"","",objrs("log_Level"),objrs("log_AuthorID"),time,objrs("log_CommNums"),"","",objrs("log_Url"),"","",objrs("log_FullUrl"),objrs("log_Type"),"")));
				subtemplate[i]=template.replace("$url$",s.fullurl);
				subtemplate[i]=subtemplate[i].replace("$title$",s.title);
				subtemplate[i]=subtemplate[i].replace("$title_sort$",s.title.substr(0,20));
				subtemplate[i]=subtemplate[i].replace("$commnums$",s.commnums);
				objrs.MoveNext;
			}
			if(i<=0){return ""};
			return subtemplate.join("")

		}
	}		
		,访问最多文章:function(save){
			if(this.config.访问最多文章==0){return false;}
			var strContent=this.functions.makeview("",this.config.访问最多文章);
			var id="HottestArticle";
			var title="最热文章";
			if(save){return this.functions.savefunction(id,title,"div"+id,strContent);}
			return strContent;
		}
		
		,本月最热文章:function(save){
			if(this.config.本月最热文章==0){return false;}
			var strContent=this.functions.makeview("([log_PostTime]>"+(ZC_MSSQL_ENABLE?"getdate()":"Now()")+"-30)",this.config.本月最热文章);
			var id="MonthHottestArticle";
			var title="本月最热文章";
			if(save){return this.functions.savefunction(id,title,"div"+id,strContent);}
			return strContent;
		}
		,本年最热文章:function(save){
			if(this.config.本年最热文章==0){return false;}
			var strContent=this.functions.makeview("([log_PostTime]>"+(ZC_MSSQL_ENABLE?"getdate()":"Now()")+"-365)",this.config.本年最热文章);
			var id="YearHottestArticle";
			var title="本年最热文章";
			if(save){return this.functions.savefunction(id,title,"div"+id,strContent);}
			return strContent;
		}
		,分类最热文章:function(save){
			if(this.config.分类最热文章==0){return false;}
			var jsCate;
			jsCate=new VBArray(Categorys).toArray();
			for(var i=0;i<jsCate.length;i++){
				var strContent=this.functions.makeview("[log_cateid]="+jsCate[i].ID,this.config.分类最热文章);
				var id="Cate"+jsCate[i].ID+"HottestArticle";
				var title=jsCate[i].Name+"最热文章";
				if(save){this.functions.savefunction(id,title,"div"+id,strContent);}
			//return strContent;
			
			}
			return true;
		}
		,评论最多文章:function(save){
			if(this.config.评论最多文章==0){return false;}
			var strContent=this.functions.makecomm("",this.config.评论最多文章);
			var id="MostCommentedArticle";
			var title="最被吐槽";
			if(save){return this.functions.savefunction(id,title,"div"+id,strContent);}
			return strContent;
		}
		
		,本月评论最多:function(save){
			if(this.config.本月评论最多==0){return false;}
			var strContent=this.functions.makecomm("([log_PostTime]>"+(ZC_MSSQL_ENABLE?"getdate()":"Now()")+"-30)",this.config.本月评论最多);
			var id="MonthMostCommentedArticle";
			var title="本月评论最多";
			if(save){return this.functions.savefunction(id,title,"div"+id,strContent);}
			return strContent;
		}
		,本年评论最多:function(save){
			if(this.config.本年评论最多==0){return false;}
			var strContent=this.functions.makecomm("([log_PostTime]>"+(ZC_MSSQL_ENABLE?"getdate()":"Now()")+"-365)",this.config.本年评论最多);
			var id="YearMostCommentedArticle";
			var title="本年评论最多";
			if(save){return this.functions.savefunction(id,title,"div"+id,strContent);}
			return strContent;
		}
		,分类评论最多:function(save){
			if(this.config.分类评论最多==0){return false;}
			var Category,jsCate;
			jsCate=new VBArray(Categorys).toArray();
			for(var i=0;i<jsCate.length;i++){
				var strContent=this.functions.makecomm("[log_cateid]="+jsCate[i].ID,this.config.分类评论最多);
				var id="Cate"+jsCate[i].ID+"MostCommentedArticle";
				var title=jsCate[i].Name+"评论最多";
				if(save){this.functions.savefunction(id,title,"div"+id,strContent);}
			//return strContent;
			}
			return true;
		}
		,随机文章:function(save){
			if(this.config.随机文章==0){return false}
			var subtemplate=new Array(this.config.随机文章);
			var template="<li><a href=\"$url$\" title=\"$title$\">$title_sort$</a></li>"
			var s=NewClass("TArticle");
			var objrs=objconn.Execute("SELECT TOP "+this.config.随机文章+" [log_ID],[log_CateID],[log_Title],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_Url],[log_FullUrl],[log_Type],[log_CommNums] FROM [blog_Article] WHERE [log_Level]=4 AND [log_Type]=0 ORDER BY " + (ZC_MSSQL_ENABLE?"newid()":"rnd(log_id)"));
			for(var i=0;i<=this.config.随机文章;i++){
				if(objrs.EOF){break;}
				var time=jsTimetovbs_vbs(objrs("log_PostTime"));
				s.loadinfobyarray(jsarraytovbs_js(new Array(objrs("log_ID"),"",objrs("log_CateID"),objrs("log_Title"),"","",objrs("log_Level"),objrs("log_AuthorID"),time,"","","",objrs("log_Url"),"","",objrs("log_FullUrl"),objrs("log_Type"),"")));
				subtemplate[i]=template.replace("$url$",s.fullurl);
				subtemplate[i]=subtemplate[i].replace("$title$",s.title);
				subtemplate[i]=subtemplate[i].replace("$title_sort$",s.title.substr(0,20));
				objrs.MoveNext;
			}
			if(save){return this.functions.savefunction("RandomArticle","随机文章","divRandomArticle",subtemplate.join(""));}
			return subtemplate.join("");

		}
		,分类:function(){
			if(this.config.分类最热文章==0){return false}
			var template="<li><a href=\"$url$\" title=\"$title$\">$title_sort$</a></li>"
			var s=NewClass("TArticle");
			var Category,jsCate;
			jsCate=new VBArray(Categorys).toArray();
			for(var i=0;i<jsCate.length;i++){
				if(advancedfunction.cls.config.Read("分类_"+jsCate[i].ID)!=""){
					var subtemplate=new Array(this.config.分类最热文章);
					var objrs=objconn.Execute("SELECT TOP "+this.config.分类最热文章+" [log_ID],[log_CateID],[log_Title],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_Url],[log_FullUrl],[log_Type],[log_CommNums] FROM [blog_Article] WHERE [log_Level]=4 AND [log_Type]=0 AND [log_CateID]="+jsCate[i].ID+" ORDER BY [log_PostTime] DESC");
					for(var sm=0;sm<=this.config.分类最热文章;sm++){
						if(objrs.EOF){break;}
						var time=jsTimetovbs_vbs(objrs("log_PostTime"));
						s.loadinfobyarray(jsarraytovbs_js(new Array(objrs("log_ID"),"",objrs("log_CateID"),objrs("log_Title"),"","",objrs("log_Level"),objrs("log_AuthorID"),time,"","","",objrs("log_Url"),"","",objrs("log_FullUrl"),objrs("log_Type"),"")));
						subtemplate[sm]=template.replace("$url$",s.fullurl);
						subtemplate[sm]=subtemplate[sm].replace("$title$",s.title);
						subtemplate[sm]=subtemplate[sm].replace("$title_sort$",s.title.substr(0,20));
						objrs.MoveNext;
					}
					
					this.functions.savefunction("Cate"+jsCate[i].ID+"Article",jsCate[i].Name,"divCate"+jsCate[i].ID,subtemplate.join(""));

				}
			}
		}
		
		,run:function(strs){
			
			this.init();
			var spt=strs.split(",");
			var attr;
			for(attr in spt){
				eval("advancedfunction."+spt[attr]+"(true)");
			}
			
		}
		
		
		

}
function jsarraytovbs_js(jsarray){return jsArrayTovbs_vbs(jsarray.join(String.fromCharCode(1)))}
function IIIf(ex,ex2,t1,t2,t3){
	if(ex){return t1}
	if(ex2){return t2}
	return t3
}
</script>
<%
Function jsArrayTovbs_vbs(jsarraystr)
	jsArrayTovbs_vbs=Split(jsarraystr,Chr(1))
End Function
Function jsTimetovbs_vbs(s)
	jsTimetovbs_vbs=CStr(CDate(s))
End Function
%>