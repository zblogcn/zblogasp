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

//$.getScript("<%=BlogHost%>zb_system/admin/ueditor/third-party/SyntaxHighlighter/shCore.js",function(){SyntaxHighlighter.all();});
$("head").append("<link rel='stylesheet' type='text/css' href='<%=BlogHost%>zb_system/admin/ueditor/third-party/SyntaxHighlighter/shCoreDefault.css'/>");


var str00="<%=BlogHost%>";
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

	var objImageValid=$("img[src*='zb_system/function/c_validcode.asp?name=commentvalid']");
	if(objImageValid.size()>0){
		objImageValid.css("cursor","pointer");
		objImageValid.click( function(){objImageValid.attr("src","<%=BlogHost%>zb_system/function/c_validcode.asp?name=commentvalid"+"&amp;random="+Math.random());});
	};

	if(!$("li.msgarticle").html()){$("ul.mutuality").hide()}
	if($("ul.msghead ~ ul.msg").length==0){$("ul.msghead").hide()}
	$(".post-tags").each(function(){if(!$(this).find(".tags").html()){$(this).hide()}});

	$.getScript("<%=BlogHost%>zb_system/function/c_html_js.asp?act=batch"+unescape("%26")+"view=" + escape(strBatchView)+unescape("%26")+"inculde=" + escape(strBatchInculde)+unescape("%26")+"count=" + escape(strBatchCount));

		//为了在编辑器之外能展示高亮代码
    	 SyntaxHighlighter.highlight();

    	 //调整左右对齐
    	 for(var i=0,di;di=SyntaxHighlighter.highlightContainers[i++];){
         	   var tds = di.getElementsByTagName('td');
            	for(var j=0,li,ri;li=tds[0].childNodes[j];j++){
                	ri = tds[1].firstChild.childNodes[j];
                	ri.style.height = li.style.height = ri.offsetHeight + 'px';
            	}
    	 }

	if(GetCookie("username")!=""&&GetCookie("password")!=""){$.getScript("<%=BlogHost%>zb_system/function/c_html_js.asp?act=autoinfo")}

	var s=document.location;
	$("#divNavBar").find("a").each(function(){if($(this).attr("href")==s.toString().split("#")[0]){$(this).addClass("on");return false;}});


	//日历当天高亮显示
	try{eval("document.getElementById(\"pCalendar_\"+(new Date()).getFullYear()+\"_\"+((new Date()).getMonth()+1)+\"_\"+(new Date()).getDate()).className+=\" cd\"");}catch(e){};

});

<%=Response_Plugin_Html_Js_Add%>