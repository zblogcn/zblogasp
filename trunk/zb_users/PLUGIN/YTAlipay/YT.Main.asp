<script language="javascript" src="../YTCMS/Config.js" runat="server" type="text/javascript"></script>
<!--#include file="YT.Lib.asp" -->
<%
Call RegisterPlugin("YTAlipay","ActivePlugin_YT_Alipay")
Sub ActivePlugin_YT_Alipay()

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