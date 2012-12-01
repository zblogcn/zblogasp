<script language="javascript" src="../YTCMS/Config.js" runat="server" type="text/javascript"></script>
<!--#include file="YT.Lib.asp" -->
<%
Call RegisterPlugin("YTAlipay","ActivePlugin_YT_Alipay")
Sub ActivePlugin_YT_Alipay()
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"订单管理",GetCurrentHost&"zb_users/plugin/YTAlipay/YT.Panel.asp","nav_quoted","aYTAlipayMng",""))
End Sub
'卸载插件
Function UnInstallPlugin_YTAlipay()
	Call new YT_Alipay.UnInstall()
End Function
'安装插件
Function InstallPlugin_YTAlipay()
	Call new YT_Alipay.Install()
End Function
%>