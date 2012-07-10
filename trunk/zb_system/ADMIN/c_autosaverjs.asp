<%@CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    sipo
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_autosaverjs.asp
'// 开始时间:    2006-1-19
'// 最后修改:    2006-7-27
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_function_md5.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%
Response.ContentType="application/x-javascript"
Call System_Initialize()

Public ZC_AUTOSAVE_FILENAME
ZC_AUTOSAVE_FILENAME="autosave"&"_"&MD5(ZC_BLOG_HOST & ZC_BLOG_CLSID & BlogUser.Name)&".txt"  


IF IsEmpty(ReQuest.QueryString("act")) Then
	SaveContent()
ElseIf Request.QueryString("act")="edit" then
	ExportAutoSaveJS()
End IF

'*********************************************************
' 目的：    Convert Bytes To Str
'*********************************************************
Function BytesToBstr(body,Cset)
		On Error Resume Next
		Dim objstream
		Set objstream = Server.CreateObject("adodb.stream")
		objstream.Type = 1
		objstream.Mode =3
		objstream.Open
		objstream.Write body
		objstream.Position = 0
		objstream.Type = 2
		objstream.Charset = Cset
		BytesToBstr = objstream.ReadText 
		objstream.Close
		Set objstream = Nothing
End Function

'*********************************************************
' 目的：    Save Draft And DisPlay
'*********************************************************
Function SaveContent()
		If BlogUser.Level>3 Then
		Response.Write ZC_MSG259
		Response.End 
		End If
		On Error Resume Next
		Dim objStream
		Set objStream = Server.CreateObject("ADODB.Stream")
		With objStream
		.Type = 2
		.Mode = 3
		.Open
		.Charset = "utf-8"
		.Position = objStream.Size
		.WriteText=BytesToBstr(Request.BinaryRead(Request.TotalBytes),"UTF-8")
		.SaveToFile Server.MapPath("../../ZB_USERS/CACHE/"&ZC_AUTOSAVE_FILENAME),2
		.Close
		End With
		Set objStream = NoThing
		If Err.Number=0 then
		Response.Write "<span style="""">&nbsp;"&formatdatetime(now,4)&":"&Right("0"&second(now),2)&"<a href="""&ZC_BLOG_HOST&"zb_users/CACHE/"&ZC_AUTOSAVE_FILENAME&""" target=""_blank"" style=""text-decoration: none;"">"&ZC_MSG258&"</a>&nbsp;</span>"
		Else
		Response.Write "<span style="""">&nbsp;"&formatdatetime(now,4)&""&ZC_MSG257&"&nbsp;"&Err.Number&Err.description&"</span>"
		End If
		Response.End
End Function


'*********************************************************
' 目的：   输出自动保存脚本
'*********************************************************
Function ExportAutoSaveJS()
	Response.Clear
	'//////////////
	Response.Write "  function init(){"
	If Request.QueryString("type")="normal" Then Response.Write "init_edit();return postForm.value;"
	If Request.QueryString("type")="ueditor" Then Response.Write "init_ueditor();return editor.getContent();"
	Response.Write "  }"
	Response.Write "  function restore(obj){"
	If Request.QueryString("type")="normal" Then Response.Write "init_edit();postForm.value=obj;"
	If Request.QueryString("type")="ueditor" Then Response.Write "init_ueditor();return editor.setContent(obj);"

	Response.Write "  }"
	'/////////////
	Response.Write "  var AutoSaveTime=60;"
	Response.Write "  var FileName="""&ZC_BLOG_HOST&"zb_users/CACHE/"&ZC_AUTOSAVE_FILENAME&""";"
	Response.Write "  var postForm = null; "
	Response.Write "  var msg = null; "
	Response.Write "  function init_edit(){"
	Response.Write "  postForm = document.edit.txaContent;"
	Response.Write "  msg = document.getElementById(""msg"");"
	Response.Write "  }"
	Response.Write "  function init_ueditor(){"
	Response.Write "  msg = document.getElementById(""msg"");"
	Response.Write "  postForm = document.edit.ueditor;"
	Response.Write "  }"
	'/////////////
	'/////////////
	Response.Write "var ti=AutoSaveTime;"
	Response.Write "function savedraft()"
	Response.Write "{	 init();"
	Response.Write "	if (postForm!=null&&typeof(postForm)!=undefined){"
	Response.Write "		var url = ""c_autosaverjs.asp"";"
	Response.Write "		var postStr = init();"
	Response.Write "		if (postStr){"
	Response.Write "		var ajax = getHTTPObject();"
	Response.Write "		ajax.open('POST', url, true); "
	Response.Write "		ajax.setRequestHeader(""Content-Type"",""application/x-www-form-urlencoded""); "
	Response.Write "		ajax.onreadystatechange = function(){if (ajax.readyState == 4 && ajax.status == 200) msg.innerHTML = ajax.responseText;};"
	Response.Write "		ajax.send(postStr);"
	Response.Write "		ti=-1000;"
	Response.Write "		}else{"
	Response.Write "		msg.innerHTML = """&ZC_MSG256&""";"
	Response.Write "		ti=-1000;}"
	Response.Write "	}else{msg.innerHTML = """&ZC_MSG255&""";ti=-1000;}"
	Response.Write "}"
	Response.Write "function restoredraft()"
	Response.Write "{ init();"
	Response.Write "if (window.confirm('"&ZC_MSG254&"'))"
	Response.Write "{"
	Response.Write "	if (postForm!=null&&typeof(postForm)!=undefined){"
	Response.Write "		var url = FileName;"
	Response.Write "		var ajax = getHTTPObject();"
	Response.Write "		ajax.open(""GET"", url+'?random='+Math.random(), true); "
	Response.Write "		ajax.onreadystatechange = function() { "
	Response.Write "		if (ajax.readyState == 4 && ajax.status == 200) { "
	Response.Write "		restore(ajax.responseText);"
	Response.Write "		msg.innerHTML ="""&ZC_MSG253&"""; } } ;"
	Response.Write "		ajax.send(null); "
	Response.Write "		ti=-1000;"
	Response.Write "	}else{msg.innerHTML = """&ZC_MSG255&""";ti=-1000;}"
	Response.Write ""
	Response.Write "}"
	Response.Write "}"
	Response.Write "function Viewdraft()"
	Response.Write "{ "
	Response.Write "window.open(FileName,'','');"
	Response.Write "}"
	Response.Write "document.getElementById(""msg2"").innerHTML =""&nbsp;<a href='javascript:try{Viewdraft()}catch(e){};' style='cursor:hand;'>["&ZC_MSG015&"]</a>&nbsp;<a href='javascript:try{restoredraft()}catch(e){};' style='cursor:hand;'>["&ZC_MSG252&"]</a>&nbsp;<a href='javascript:try{savedraft()}catch(e){};' style='cursor:hand;'>["&ZC_MSG004&"]</a>"";"
	Response.Write "function timer() { "
	Response.Write "ti=ti-1;"
	Response.Write "var timemsg=document.getElementById(""timemsg"");timemsg.innerHTML = ti+"""&ZC_MSG251&""";"
	Response.Write "if (ti>=0){window.setTimeout(""timer()"", 1000);}else{if (ti<=-1000)"
	Response.Write "{ti=AutoSaveTime;timer();}else{timemsg.innerHTML = """&ZC_MSG250&"..."";savedraft"
	Response.Write "();ti=AutoSaveTime;timer();}} }"
	Response.Write "window.setTimeout(""timer()"", 0);"
	Response.Write "    function getHTTPObject() {"
	Response.Write "	var xmlhttprequest=false; "
	Response.Write "    try {"
	Response.Write "	  xmlhttprequest = new XMLHttpRequest();"
	Response.Write "	} catch (trymicrosoft) {"
	Response.Write "	  try {"
	Response.Write "		xmlhttprequest = new ActiveXObject(""Msxml2.XMLHTTP"");"
	Response.Write "	  } catch (othermicrosoft) {"
	Response.Write "		try {"
	Response.Write "		  xmlhttprequest = new ActiveXObject(""Microsoft.XMLHTTP"");"
	Response.Write "		} catch (failed) {"
	Response.Write "		  xmlhttprequest = false;"
	Response.Write "		}"
	Response.Write "	  }"
	Response.Write "	}"
	Response.Write "	return xmlhttprequest;"
	Response.Write "    }"
End Function

Call System_Terminate()
%>
