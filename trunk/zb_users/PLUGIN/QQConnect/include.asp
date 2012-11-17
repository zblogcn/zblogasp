<!-- #include file="function.asp"-->
<%
Call RegisterPlugin("QQConnect","qqconnect_include")
%>

<script language="javascript" runat="server">
function qqconnect_include(){
//创建左侧菜单
	init_qqconnect_include()
	Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(5,"QQ互联",BlogHost+"zb_users/plugin/qqconnect/main.asp","newQQConnect","anewQQConnect",BlogHost+"zb_users/plugin/qqconnect/resources/Connect_logo_1.png"));
	
	//文章发布
	Add_Action_Plugin("Action_Plugin_Edit_Form","qqconnect.include.action.edit_form(EditArticle)");
	Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","Call qqconnect.include.action.articlepst(Request.Form(\"syn_qq\"),Request.Form(\"syn_tqq\"))");
	
	//注册用户
	if(CheckPluginState("RegPage")){
		if(typeof(Request.QueryString("openid").Item)!=undefined) {
			var strQQ=TransferHTML(FilterSQL(Request.QueryString("openid")),"[nohtml]").replace("\"","\"\"");
			var strAcc=TransferHTML(FilterSQL(Request.QueryString("accesstoken")),"[nohtml]").replace("\"","\"\"");
			if((strQQ.length==32)&&(strAcc.length=32)){
				Add_Response_Plugin("Response_Plugin_RegPage_End","<input type=\"hidden\" value=\""+strQQ+"\" name=\"OpenID\"/>")
				Add_Response_Plugin("Response_Plugin_RegPage_End","<input type=\"hidden\" value=\""+strAcc+"\" name=\"AccessToken\"/>")
			}
		}
		Add_Action_Plugin("Action_Plugin_RegSave_End","If qqconnect.functions.savereg(RegUser.ID,Request(\"openid\"),Request(\"accesstoken\"))=True Then strResponse=\"<scri"+"pt language='javascript' type='text/javascript'>alert('恭喜，注册成功。\\n欢迎您成为本站一员。\\n\\n单击确定登陆本站。');location.href=\"\""+BlogHost+"zb_users/plugin/qqconnect/main.asp\"\"</scri"+"pt>\"")
		
	}
	//评论同步
	

}
function InstallPlugin_QQConnect(){
	init_qqconnect();
	qqconnect.d.CreateDB()
}
function init_qqconnect_include(){

	qqconnect["include"]={
		"action":{
			"edit_form":function(obj){
					//qqconnect["object"]=obj;
					var html;
					init_qqconnect();
					html="\r\n";
					html+="<style type=\"text/css\">.syn_qq, .syn_tqq, .syn_qq_check, .syn_tqq_check"+
						"{display:inline-block;margin-top:3px;width:19px;height:19px;line-height:64px;overflow:hidden;vertical-align:top;cursor:pointer;"+
						"background: transparent url(../../zb_users/plugin/qqconnect/resources/connect_post_syn.png) no-repeat 0 0;}"+
						".syn_tqq{background-position: 0 -22px;margin-left: 5px;}.syn_qq_check{background-position: -22px 0;}"+
						".syn_tqq_check{background-position: -22px -22px;margin-left: 5px;}"+
						"</style>";
					html+="\r\n";
					html+=(qqconnect.tconfig.read("a")=="True"?
							"<a title='已设置同步至QQ空间' class='syn_qq_check' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a>"+
							"<input type='hidden' name='syn_qq' id='syn_qq' value='1'/>"
							:
							"<a title='未设置同步至QQ空间' class='syn_qq' href='javascript:void(0);' id='connectPost_synQQ'>QQ空间</a>"+
							"<input type='hidden' name='syn_qq' id='syn_qq' value='0'/>");
					html+="\r\n";
					html+=(qqconnect.tconfig.read("b")=="True"?
							"<a title='已设置同步至腾讯微博' class='syn_tqq_check' href='javascript:void(0);' id='connectPost_syntQQ'>腾讯微博</a>"+
							"<input type='hidden' name='syn_tqq' id='syn_tqq' value='1'/>"
							:
							"<a title='未设置同步至腾讯微博' class='syn_tqq' href='javascript:void(0);' id='connectPost_syntQQ'>腾讯微博</a>"+
							"<input type='hidden' name='syn_tqq' id='syn_tqq' value='0'/>");
					html+="\r\n";
					html+="<scr"+"ipt type='text/javascript'>"+
						"var qqconnect_synQQState = "+(qqconnect.tconfig.read("a")=="True"?"false":"true")+","+
						"	qqconnect_synTState = "+(qqconnect.tconfig.read("b")=="True"?"false":"true")+";"+
						"var qqconnect_synQQ = $('#connectPost_synQQ');var qqconnect_syntQQ = $('#connectPost_syntQQ');"+
						"function qqconnect_changestate0() {if (qqconnect_synQQState) {"+
						"		qqconnect_synQQ.removeClass('syn_qq_check');qqconnect_synQQ.addClass('syn_qq');qqconnect_synQQ.attr('title', '未设置同步至QQ空间');"+
						"		$('#syn_qq').attr('value', '0');qqconnect_synQQState = false;"+
						"	} else {"+
						"		qqconnect_synQQ.removeClass('syn_qq');qqconnect_synQQ.addClass('syn_qq_check');qqconnect_synQQ.attr('title', '已设置同步至QQ空间');"+
						"		$('#syn_qq').attr('value', '1');qqconnect_synQQState = true"+
						"	}};"+
						"function qqconnect_changestate1() {"+
						"	if (qqconnect_synTState) {"+
						"		qqconnect_syntQQ.removeClass('syn_tqq_check');qqconnect_syntQQ.addClass('syn_tqq');"+
						"		qqconnect_syntQQ.attr('title','未设置同步至腾讯微博');$('#syn_tqq').attr('value', '0');qqconnect_synTState = false;"+
						"	} else {"+
						"		qqconnect_syntQQ.removeClass('syn_tqq');qqconnect_syntQQ.addClass('syn_tqq_check');"+
						"		qqconnect_syntQQ.attr('title', '已设置同步至腾讯微博');$('#syn_tqq').attr('value', '1');qqconnect_synTState = true;"+
						"	}};"+
						"$(document).ready(function() {"+
						"	qqconnect_changestate0();qqconnect_changestate1();"+
						"	qqconnect_synQQ.click(function(){qqconnect_changestate0()});qqconnect_syntQQ.click(function(){qqconnect_changestate1()});});"+
						"</scrip"+"t>";
					html+="\r\n";

					if(obj.id==0) Add_Response_Plugin("Response_Plugin_Edit_Form3",html);
				}
			,"articlepst":function(sync_zone,sync_weibo){
				var sync={"zone":(sync_zone=="0"?false:true),"weibo":(sync_weibo=="0"?false:true)}
				qqconnect["temp"]=sync;
				Add_Filter_Plugin("Filter_Plugin_PostArticle_Core","qqconnect.include.filter.postarticle")
				
			
			}
		}
		,"filter":{
			"postarticle":function(object){
				if(object.ID==0) Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","qqconnect.include.filter.postarticle_succeed")
			}
			,"postarticle_succeed":function(object){
				if(object.ID==0||object.FType==1||object.Level<=2){return false}
				var strT ,bolN,objTemp,strTemp;
				init_qqconnect();
				//Set ZBQQConnect_DB.objUser=BlogUser
				//ZBQQConnect_DB.LoadInfo 2
				//ZBQQConnect_Class.OpenID=ZBQQConnect_DB.OpenID
				//ZBQQConnect_Class.AccessToken=ZBQQConnect_DB.AccessToken
				strTemp=object.Intro.replace("<#ZC_BLOG_HOST#>",blogHost);
				strTemp=qqconnect.functions.formatstring(strTemp);
				var t_add,tupian;
				if(qqconnect.tconfig.read("c")=="True") {
					tupian=UBBCode(object.Content,"[image]");
					tupian=tupian.replace("<#ZC_BLOG_HOST#>",blogHost);
					tupian=qqconnect.functions.getpicture(tupian);
					if(tupian=="") tupian="~";//qqconnect.tconfig.read("");
					tupian=tupian.replace("\\","/").replace("'","\'");
				}else{
					tupian="~"
				}
				if (qqconnect.temp.zone){
					qqconnect.config.qqconnect.openid=qqconnect.config.qqconnect.admin.openid
					qqconnect.config.qqconnect.accesstoken=qqconnect.config.qqconnect.admin.accesstoken
					t_add = qqconnect.c.Share(object.Title,object.Url.toString().replace("<#ZC_BLOG_HOST#>",BlogHost),"",strTemp,tupian,1);
					t_add = qqconnect.functions.json.toObject(t_add);
					if(t_add.ret==0){
						SetBlogHint_Custom("同步到QQ空间成功！")
					}
					else{
						SetBlogHint_Custom("同步到QQ空间出现问题" + t_add.ret)
						//Response.Write "调试信息：<br/>"&ZBQQConnect_class.debugMsg
					}
					qqconnect.c.debugMsg="";
				}
				
				if (qqconnect.temp.weibo){
					var str=qqconnect.tconfig.Read("content");
					str=str.replace("%t",qqconnect.functions.formatstring(object.Title));
					str=str.replace("%u",encodeURI(object.Url));
					str=str.replace("%b",qqconnect.functions.formatstring(BlogTitle));
					str=str.replace("%i",strTemp);
					str=str.replace("%3C#ZC_BLOG_HOST#%3E",BlogHost);
					t_add = qqconnect.t.t(str,tupian)
					t_add= qqconnect.functions.json.toObject(t_add)
					if(t_add.ret==0){
						SetBlogHint_Custom("恭喜，同步到腾讯微博成功")
						object.Meta.SetValue("ZBQQConnect_WBID",t_add.data.id);
						object.Post()
					}
					else{
						SetBlogHint_Custom("同步到腾讯微博出现问题" + t_add.ret)
					}
				}
			}
		}
	}
}

</script>