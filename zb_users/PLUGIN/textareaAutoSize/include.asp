<%
'==============Z-Blog===============
'=============未寒=============
'=====http://imzhou.com/=======


Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/textareaAutoSize/jquery.autosize.js\""></script>"");")
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/textareaAutoSize/autosize.js\""></script>"");")


'注册插件
Call RegisterPlugin("textareaAutoSize","ActivePlugin_textareaAutoSize")


%>