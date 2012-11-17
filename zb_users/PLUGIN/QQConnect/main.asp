<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
init_qqconnect()

If CheckPluginState("QQConnect")=False Then Call ShowError(48)
BlogTitle="QQ互联"

Select Case Request.QueryString("act")
	Case "callback"
		Select Case Request.QueryString("type")
			Case "connect"
					Call qqconnect.c.GetOpenId(qqconnect.c.CallBack)
					If BlogUser.Level=1 Then
						Call qqconnect.tconfig.write("Connect_OpenID",qqconnect.config.qqconnect.openid)
						Call qqconnect.tconfig.write("Connect_AccessToken",qqconnect.config.qqconnect.accesstoken)
						Call qqconnect.tconfig.Save
					End If
					SetBlogHint True,Empty,Empty
					Response.Redirect "main.asp"
			Case "weibo"
				If BlogUser.Level=1 Then
					Call qqconnect.t.Run(11,"","","","")
					Call qqconnect.tconfig.write("weibo_token",qqconnect.config.weibo.token)
					Call qqconnect.tconfig.write("weibo_secret",qqconnect.config.weibo.secret)
					Call qqconnect.tconfig.Save
					SetBlogHint True,Empty,Empty
					Response.Redirect "main.asp"
				End If
		End Select
		Response.End
	Case "logout"
		Select Case Request.QueryString("type")
			Case "connect"
				If BlogUser.Level=1 Then
					Call qqconnect.tconfig.write("Connect_OpenID","")
					Call qqconnect.tconfig.write("Connect_AccessToken","")
					Call qqconnect.tconfig.Save
				End If
				SetBlogHint True,Empty,Empty
				Response.Redirect "main.asp"
			Case "weibo"
				If BlogUser.Level=1 Then
					Call qqconnect.tconfig.write("weibo_token","")
					Call qqconnect.tconfig.write("weibo_secret","")
					Call qqconnect.tconfig.Save
					SetBlogHint True,Empty,Empty
					Response.Redirect "main.asp"
				End If
		End Select
	Case "weibo"
		Select Case Request.QueryString("type")
			Case "latest"
				If BlogUser.Level=1 Then
					Response.Write qqconnect.t.api("http://open.t.qq.com/api/statuses/home_timeline","{}","GET")
				End If
			Case "mention"
				If BlogUser.Level=1 Then
					Response.Write qqconnect.t.api("http://open.t.qq.com/api/statuses/mentions_timeline","{}","GET")
				End If
			Case "new"	
				If BlogUser.Level=1 Then
					Response.Write qqconnect.t.t(Request.Form("data"),"")
				End If
		End Select
		Response.End
End Select

Call CheckReference("")
%>
    
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">QQ互联</div>
<div class="SubMenu"><%=qqconnect.functions.navbar(0)%></div>
<div id="divMain2">
<div id="ShowBlogHint"><%=GetBlogHint%></div>
<table width="100%" border="1">
    <%
Response.Write "<tr height='32'><td>"
Dim tmpObject
If qqconnect.config.qqconnect.appid<>"" Then
	If BlogUser.Level=1 Then
		qqconnect.config.qqconnect.openid=qqconnect.config.qqconnect.admin.openid
		qqconnect.config.qqconnect.accesstoken=qqconnect.config.qqconnect.admin.accesstoken
		If qqconnect.config.qqconnect.openid="" Then
			Response.Write "<a href='" & qqconnect.c.Authorize() & "'><img src='resources/logo_170_32.png'/></a>"
		Else
			Set tmpObject=qqconnect.functions.json.toobject(qqconnect.c.api("https://graph.qq.com/user/get_user_info","{}","GET"))
			Response.Write "欢迎回来，QQ空间用户" & tmpObject.nickname & "<a href='main.asp?act=logout&type=connect'>点击这里注销</a>"
		End If
	Else
		'xxxxx
	End If
Else
	Response.Write "未配置QQ互联APPID，无法使用本功能。"
End If
Response.Write "</td></tr>"
%>    

<%
Response.Write "<tr height='32'><td>"
If BlogUser.Level=1 Then
	If qqconnect.config.weibo.token="" Then
		Response.Write "<a href='" & qqconnect.t.Run(1,"","","","") & "'><img src='resources/wb_170_32.png'/></a>"
	Else
		Set tmpObject=qqconnect.functions.json.toobject(qqconnect.t.api("http://open.t.qq.com/api/user/info","{}","GET"))
		Response.Write "欢迎回来，腾讯微博用户" & tmpObject.data.nick & "(" & tmpObject.data.name & ") <a href='main.asp?act=logout&type=weibo'>点击这里注销</a>"
		Response.Write "<p>&nbsp;</p><p><input type='text' style='width:50%' value='' id='zsx'/><input type='submit' id='ok' value='发微博'/></p>"
		Response.Write "<p>&nbsp;</p><p><a href='javascript:;' id='href1'>得到最新微博</a>&nbsp;<a href='javascript:;' id='href2'>得到提到我的</a></p><p>&nbsp;</p>"
	End If
End If
Response.Write "</td></tr></table><p>&nbsp;</p><div id='latest'>Loading</div>"
%>

<!--</table>-->
<script type="text/javascript">
var j=$("#latest");
if(j){
	$("#ok").click(function(){
		$.post("main.asp?act=weibo&type=new",{"data":$("#zsx").val()},function(data){
			var json=eval("("+data+")").data;
			var d="";
			if(json==null){d="发送失败"}else{d="发送成功，<a href='http://t.qq.com/p/t/"+json.id+"' target='_blank'>点击这里查看</a>"}
			//var d=data;
			j.html(d);
		}
		)
	});
	$("#href1").click(function(){
		$.get("main.asp?act=weibo&type=latest",{},function(data){
			exportjson(data,j);
		})
		})
	$("#href2").click(function(){
		$.get("main.asp?act=weibo&type=mention",{},function(data){
			exportjson(data,j);
		})
		});
	$("#href1").click()
	};

function exportjson(data,obj){
	var json=eval("("+data+")").data.info;
	str="<table><tr height='32'><th width='300px'>发布者</th><th>内容</th></tr>"
	for(var i=0;i<json.length;i++){
		str+="<tr height='32'><td><a href='http://t.qq.com/p/t/"+json[i].id+"' target='_blank'>"+json[i].nick+"("+
		json[i].name+")</a></td><td>"+json[i].text+"</td></tr>"
	}
	str+="</table>"
	obj.html(str);
	bmx2table()
	}
ActiveLeftMenu("anewQQConnect");
</script>


</div>
</div>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
