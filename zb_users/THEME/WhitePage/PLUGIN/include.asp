<%
Call RegisterPlugin("WhitePage","ActivePlugin_WhitePage")

Function ActivePlugin_WhitePage()
    Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"主题配置",BlogHost & "zb_users/theme/whitepage/plugin/main.asp","aWhitePageManage",""))
End Function
%>