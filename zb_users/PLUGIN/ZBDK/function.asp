<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->

<script language="javascript" runat="server">
var title="Z-Blog Plugin Development Kit"
function ZBDK(){return false}
ZBDK.main=function(){
	Response.Write("ZBDK，全称Z-Blog Plugin Development Kit，是为插件开发人员开发的一套工具包。它集合了许多插件开发中常用的工具，可以帮助插件开发者更好地进行插件开发。"+(Request.ServerVariables("HTTP_USER_AGENT").Item.toLowerCase().indexOf("ie 6")>0?"<font color='red'>但不支持IE6.</font>":"")+"<br/><br/>");
	Response.Write("该版本ZBDK最后更新时间：2012-11-24<br/><br/>");
	Response.Write("该插件有一定的危险性，一旦进行了误操作可能导致博客崩溃，请谨慎使用。<br/><br/>");
	Response.Write("工具列表：\n\n<br/><br/><table width='100%'>");
	Response.Write("<tr height='40'><td width='50'>ID</td><td width='120'>工具名</td><td>信息</td></tr>");
	Response.Write("<tr height='40'><td>1</td><td><a href='BlogConfig/main.asp'>BlogConfig</a></td><td>可以对blog_Config里的数据进行管理，用于调试TConfig类。</td></tr>");
	Response.Write("<tr height='40'><td>2</td><td><a href='LoadCounter/main.asp'>LoadCounter</a></td><td>日志读取器，对于调试一些无法直接在前台显示内容的插件有很大用处。</td></tr>");
	Response.Write("<tr height='40'><td>3</td><td><a href='RunSQL/main.asp'>RunSQL</a></td><td>SQL语句在线运行器，可以查看SQL语句的运行情况。</td></tr>");
	Response.Write("<tr height='40'><td>4</td><td><a href='PluginInterface/main.asp'>PluginInterface</a></td><td>可以查看某个接口被哪些插件挂上和挂上接口的顺序。</td></tr>");
	Response.Write("<tr height='40'><td>5</td><td><a href='OnlinePlugin/main.asp'>OnlinePlugin</a></td><td>不需要再创建或编辑现有插件就可以挂接口的工具（呃。。）</td></tr>");



}
ZBDK.submenu=function(j){
	var aryname=new Array("首页","BlogConfig","LoadCounter","RunSQL","PluginInterface","OnlinePlugin");
	var aryurl=new Array("main.asp","BlogConfig/main.asp","LoadCounter/main.asp","RunSQL/main.asp","PluginInterface/main.asp","OnlinePlugin/main.asp");
	var arycss=new Array("m-left","m-left","m-left","m-left");
	for(var i=0;i<=aryname.length;i++){
		Response.Write(MakeSubMenu(aryname[i],BlogHost+"zb_users/plugin/zbdk/"+aryurl[i],((j==i||j==aryname[i])?arycss[i]+" m-now":arycss[i]),false));
	}
}


</script>