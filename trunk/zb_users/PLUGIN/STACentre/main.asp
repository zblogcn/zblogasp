<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接 
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)
BlogTitle="静态管理中心" 
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
#headWrap {
	width: 100%;
	height: 35px;
	position: fixed;
	left: 0;
	z-index: 299;
	-moz-box-shadow: 0 2px 3px rgba(0,0,0,.12);
	-webkit-box-shadow: 0 2px 3px rgba(0,0,0,.12);
	box-shadow: 0 2px 3px rgba(0,0,0,.12);
	background: #333;
	opacity: 0.8;
	filter: alpha(opacity=80);
	color: white
}
.headInside {
	width: 1024px;
	margin: 0 auto;
	position: relative;
	z-index: 104;
	height: 35px;
}
.headInside h1 {
	position: absolute;
	left: 0;
	top: 0;
	color: white;
	margin: .20em 0;
}
.headInside h1 a {
	color: white
}
.topNav {
	position: absolute;
	left: 130px;
	top: 0;
	height: 35px;
}
.topNav .topNavItem {
	float: left;
	position: relative;
	margin-right: 1px;
}
.topNav a {
	color: white;
	font-size: 14px;
	display: block;
	height: 12px;
	line-height: 10px;
	line-height: 13px\9;
	overflow: hidden;
	padding: 11px 0 12px 0;
	_float: left;
	_cursor: pointer;
	_position: relative;
}
.topNav .topNavItem a i {
	position: absolute;
	display: block;
	height: 35px;
	width: 100%;
	top: 0;
	left: 0;
}
.topNav .topNavItem u {
	position: relative;
	z-index: 105;
	padding: 0 10px;
	text-decoration: none;
}
.disable{
	display:none
}
.topNav .topNavItem a:hover i, .topNav .topNavItem .active i {
background-color: #666;
opacity: .35;
filter: alpha(opacity=35);
}
</style>
<script type="text/javascript" src="function.js"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="headWrap" style="display:none">
          <div class="headInside">
            <h1><a title="快速插入" href="javascript:;">快速插入</a></h1>
            <ul class="topNav navMenu">
              <li class="topNavItem" _enable="{%host%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%host%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%post%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%post%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%category%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%category%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%author%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%author%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%year%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%year%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%month%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%month%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%day%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%day%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%date%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%date%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%id%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%id%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%alias%}"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>{%alias%}</u><i></i></a></li>
              <li class="topNavItem" _enable="POST_FOLTER"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>POST</u><i></i></a></li>
			  <li class="topNavItem" _enable="POST_FOLTER"><a href="javascript:;" onclick="InsertText(_focus,$(this).text(),false)"><u>ARCHIVES</u><i></i></a></li>
               
            </ul>
          </div>
        </div>
        <script type="text/javascript">
		$(function(){
			var _h=$("#headWrap").html();
			$("#headWrap").remove();
			$("#header").before('<div id="headWrap" style="display:none">'+_h+'</div>');
			$("#headWrap").show()
		}())
		</script>
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">配置页面</span></a><a href="list.asp"><span class="m-left">ReWrite规则</span></a><a href="help.asp"><span class="m-right">帮助</span></a> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <form id="form" name="form" method="post" action="save.asp">
              <input type="hidden" name="edtZC_POST_STATIC_MODE" id="edtZC_POST_STATIC_MODE" value="<%=ZC_POST_STATIC_MODE%>">
              <table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>
                <tr>
                  <td width='30%'><p align='left'><b>·全局静态化选项</b><br>
                      <span class='note'>&nbsp;&nbsp;使用伪静态前必须确认主机是否支持</span></p></td>
                  <td><p>
                      <label>
                        <input type="radio" value="STATIC" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="STATIC","checked='checked'","")%>>
                        &nbsp;&nbsp;1.文章静态</label>
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <label>
                        <input type="radio" value="ACTIVE" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="ACTIVE","checked='checked'","")%>>
                        &nbsp;&nbsp;2.全局动态</label>
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <label>
                        <input type="radio" value="REWRITE" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="REWRITE","checked='checked'","")%>>
                        &nbsp;&nbsp;3.全局伪静态</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·文章存放目录</b><br>
                      <span class='note'>&nbsp;&nbsp;静态生成文章的目录,也就是{%post%}参数的值</span></p></td>
                  <td><p>
                      <input id='edtZC_STATIC_DIRECTORY' _enblist="POST_FOLTER" name='edtZC_STATIC_DIRECTORY' style='width:500px;' type='text' value='<%=ZC_STATIC_DIRECTORY%>'>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·文章的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_ARTICLE_REGEXT' _enblist="{%category%}{%author%}{%year%}{%month%}{%day%}{%id%}{%alias%}" name='edtZC_ARTICLE_REGEX' style='width:500px;' type='text' value='<%=ZC_ARTICLE_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a></p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(1,1)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;配置1:文章名型(默认) http://www.yourblog/post/articlename.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,2)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;配置2:日期+文章名型 http://www.yourblog/2012/08/articlename.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,3)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;配置3:分类别名+文章名型 http://www.yourblog/categroyname/articlename.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,4)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;配置4:文章名目录型 http://www.yourblog/post/articlename/</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,5)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;配置5:分类别名+文章ID目录型 http://www.yourblog/categroyname/123/</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·页面的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_PAGE_REGEX' _enblist="{%category%}{%author%}{%year%}{%month%}{%day%}{%id%}{%alias%}" name='edtZC_PAGE_REGEX' style='width:500px;' type='text' value='<%=ZC_PAGE_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a> </p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(2,1)">
                        <input type="radio" name="radio2">
                        &nbsp;&nbsp;配置1:页面名型(默认) http://www.yourblog/pagename.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(2,2)">
                        <input type="radio" name="radio2">
                        &nbsp;&nbsp;配置2:页面名目录型 http://www.yourblog/pagename/</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·首页分页的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_DEFAULT_REGEX' _enblist="" name='edtZC_DEFAULT_REGEX' style='width:500px;' type='text' value='<%=ZC_DEFAULT_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a></p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(6,1)">
                        <input type="radio" name="radio6">
                        &nbsp;&nbsp;配置1:首页分页(默认) http://www.yourblog/default_2.html</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·分类页的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_CATEGORY_REGEX' _enblist="{%id%}{%alias%}"  name='edtZC_CATEGORY_REGEX' style='width:500px;' type='text' value='<%=ZC_CATEGORY_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                      </label>
                    </p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(3,1)">
                        <input type="radio" name="radio3">
                        &nbsp;&nbsp;配置1:分类ID型(默认) http://www.yourblog/category-id.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(3,2)">
                        <input type="radio" name="radio3">
                        &nbsp;&nbsp;配置2:分类ID目录型 http://www.yourblog/category/id/</label>
                    </p>
                    <p>
                      <label onclick="changeval(3,3)">
                        <input type="radio" name="radio3">
                        &nbsp;&nbsp;配置3:分类别名目录 http://www.yourblog/categroy/alias/</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·作者页的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_USER_REGEX' _enblist="{%id%}{%alias%}"  name='edtZC_USER_REGEX' style='width:500px;' type='text' value='<%=ZC_USER_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                      </label>
                    </p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(7,1)">
                        <input type="radio" name="radio7">
                        &nbsp;&nbsp;配置1:作者ID型(默认) http://www.yourblog/author-1.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(7,2)">
                        <input type="radio" name="radio7">
                        &nbsp;&nbsp;配置1:作者别名型 http://www.yourblog/author-alias.html</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·TAGS页的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_TAGS_REGEX'  _enblist="{%id%}{%alias%}"  name='edtZC_TAGS_REGEX' style='width:500px;' type='text' value='<%=ZC_TAGS_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                      </label>
                    </p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(4,1)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;配置1:Tags ID型(默认) http://www.yourblog/tags-id.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(4,2)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;配置2:Tags 名称型 http://www.yourblog/tags-name.html</label>
                    </p></td>
                </tr>
                <tr>
                  <td width='30%'><p align='left'><b>·日期页的URL配置</b><br>
                      <span class='note'></span></p></td>
                  <td><p>
                      <input id='edtZC_DATE_REGEX' _enblist="{%date%}"  name='edtZC_DATE_REGEX' style='width:500px;' type='text' value='<%=ZC_DATE_REGEX%>'>
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                      </label>
                    </p></td>
                </tr>
                <tr style="display:none">
                  <td width='30%'><p></p></td>
                  <td><p>
                      <label onclick="changeval(5,1)">
                        <input type="radio" name="radio5">
                        &nbsp;&nbsp;配置1:日期型(默认) http://www.yourblog/date-2012-12.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(5,2)">
                        <input type="radio" name="radio5">
                        &nbsp;&nbsp;配置2:日期型2 http://www.yourblog/archives/2012-12.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(5,3)">
                        <input type="radio" name="radio5">
                        &nbsp;&nbsp;配置3:日期目录型 http://www.yourblog/archives/2012-12/</label>
                    </p></td>
                </tr>
              </table>
              <p><span class="note">分类,作者,TAGS可用{%id%},{%name%}或{%alias%},分类的{%alias%}为空用name代替,作者的{%alias%}为空用name代替,TAGS的{%alias%}为URLEncode编码处理过的name,日期页可用{%date%}.</span></p>
              <p><span class="note">您可以自定义静态配置,配置必须是{%host%}打头,".html"结尾,中间可以用{%post%},{%category%},{%author%},{%year%},{%month%},{%day%},{%id%},{%alias%}组合.</span></p>
              <p><span class="note">{%post%}为文章发布目录,可以在网站设置里修改,{%category%}为文章的别名,{%user%}是用户别名,{%alias%}是文章别名,没有设置文章别名系统会自动采用ID填充.</span></p>
              <!--<p><span class="star">注意:开启文章及页面和分类页的Rewrite支持选项后,请查看"ReWrite规则"并应用在主机上方能生效.</span></p>-->
              <input type="hidden" name="edtZC_STATIC_MODE" id="edtZC_STATIC_MODE" value="<%=ZC_STATIC_MODE%>">
              <input name="" type="submit" class="button" value="保存">
            </form>
          </div>
        </div>
        <script type="text/javascript">
			var _focus=document;
			$(document).ready(function(){ 
				enable("none");
				$("input[type='text']").focus(function(){_focus=this.id;enable($(this).attr("_enblist"))})
				flashradio();
				 
			});
		
		
		</script>

        
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->