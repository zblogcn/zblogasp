<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_html_js_add.asp
'// 开始时间:    2009.12.01
'// 最后修改:    
'// 备    注:    html模板脚本辅助 ADD
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.ContentType="application/x-javascript" %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->

<% Response.Clear %>

<%=LoadFromFile(Server.MapPath("../admin/ueditor/third-party/SyntaxHighlighter/shCore.js"),"utf-8")%>
SyntaxHighlighter.defaults['gutter'] = false;
SyntaxHighlighter.defaults['smart-tabs']=true;
//SyntaxHighlighter.defaults['html-script']=true;
//SyntaxHighlighter.defaults['auto-links']=false;
SyntaxHighlighter.all();
//SyntaxHighlighter.highlight(document.getElementsByTagName("pre"),null,true);



var str00="<%=GetCurrentHost()%>";
var str01="<%=ZC_MSG033%>";
var str02="<%=ZC_MSG034%>";
var str03="<%=ZC_MSG035%>";
var str06="<%=ZC_MSG057%>";
var intMaxLen="<%=ZC_CONTENT_MAX%>";
var strFaceName="<%=ZC_EMOTICONS_FILENAME%>";
var strFaceSize="<%=ZC_EMOTICONS_FILESIZE%>";
var strFaceType="<%=ZC_EMOTICONS_FILETYPE%>";
var strBatchView="";
var strBatchInculde="";
var strBatchCount="";

$(document).ready(function(){ 

	try{

		$.getScript("<%=GetCurrentHost()%>zb_system/function/c_html_js.asp?act=batch"+unescape("%26")+"view=" + escape(strBatchView)+unescape("%26")+"inculde=" + escape(strBatchInculde)+unescape("%26")+"count=" + escape(strBatchCount));

		var objImageValid=$("img[src*='zb_system/function/c_validcode.asp?name=commentvalid']");
		if(objImageValid.size()>0){
			objImageValid.css("cursor","pointer");
			objImageValid.click( function() {
					objImageValid.attr("src","<%=GetCurrentHost()%>zb_system/function/c_validcode.asp?name=commentvalid"+"&amp;random="+Math.random());
			} );
		};



	}catch(e){};

});

$("head").append("<link rel='stylesheet' type='text/css' href='<%=GetCurrentHost()%>/zb_system/ADMIN/ueditor/third-party/SyntaxHighlighter/shCoreDefault.css'/>");

//$.getScript("<%=GetCurrentHost()%>zb_system/admin/ueditor/third-party/SyntaxHighlighter/shCore.js",function(){SyntaxHighlighter.all();});