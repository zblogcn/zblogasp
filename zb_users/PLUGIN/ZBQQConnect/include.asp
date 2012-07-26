<!-- #include file="function\ZBConnectQQ_Public.asp"-->
<!-- #include file="function\ZBConnectQQ_JSON.asp"-->
<%
dim ZBQQConnect_class
set ZBQQConnect_class=new ZBQQConnect
ZBQQConnect_class.app_key="100291142"    '设置appkey
ZBQQConnect_class.app_secret="6e39bee95a58a8c99dce88ad5169a50e"  '设置app_secret
ZBQQConnect_class.callbackurl="http://www.zsxsoft.com/zblog-1-9/ZB_USERS/PLUGIN/ZBQQConnect/callback.asp"  '设置回调地址
ZBQQConnect_class.debug=false 'Debug模式设置

%>