<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	未寒(involvements)
'// 技术支持:      im@imzhou.com
'// 程序名称:     	ColorBoxImage
'// 开始时间:    	2012.10.11
'// 最后修改:       2012.10.11
'///////////////////////////////////////////////////////////////////////////////

Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<link rel=\""stylesheet\"" href=\""" & BlogHost & "zb_users/PLUGIN/ColorBoxImage/source/colorbox.css\"" type=\""text/css\"" media=\""screen\""/>"");")
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/ColorBoxImage/source/jquery.colorbox-min.js\""></script>"");")
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/ColorBoxImage/source/colorbox.js\""></script>"");")
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/PLUGIN/ColorBoxImage/source/app.js\""></script>"");")



'注册插件
Call RegisterPlugin("ColorBoxImage","ActivePlugin_ColorBoxImage")


%>