<%
'///////////////////////////////////////////////////////////////////////////////
'// 月上之木 2012.8.25
'///////////////////////////////////////////////////////////////////////////////


Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/SweetTitles/jquery.sweetTitles.js\""></script>"");")
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<link rel=\""stylesheet\"" href=\""" & BlogHost & "zb_users/PLUGIN/SweetTitles/sweetTitles.css\"" />"");")



'注册插件
Call RegisterPlugin("SweetTitles","ActivePlugin_SweetTitles")

Function ActivePlugin_SweetTitles()


End Function


'安装插件
Function InstallPlugin_SweetTitles


End Function

'卸载插件
Function UnInstallPlugin_SweetTitles


End Function
%>