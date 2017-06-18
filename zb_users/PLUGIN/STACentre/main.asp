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
	height: 25px;
	position: absolute;
	left: 0;
	z-index: 299;
/*	-moz-box-shadow: 0 2px 3px rgba(0,0,0,.12);*/
/*	-webkit-box-shadow: 0 2px 3px rgba(0,0,0,.12);*/
/*	box-shadow: 0 2px 3px rgba(0,0,0,.12);*/
/*	background: #333;*/
	opacity: 0.8;
	filter: alpha(opacity=80);
	color: white
}
.headInside {
/*	width: 890px;*/
	margin: 0 auto;
	position: relative;
	z-index: 104;
	height: 25px;
	-moz-box-shadow: 0 2px 3px rgba(0,0,0,.12);
	-webkit-box-shadow: 0 2px 3px rgba(0,0,0,.12);
	box-shadow: 0 2px 3px rgba(0,0,0,.12);
	background: #333;
}
.headInside h1 {
/*	position: absolute;*/
	left: 0;
	top: 0;
	color: white;
	margin: .20em .50em;
	font-size: 1em;
	width: 80px;
	float: left;
}
.headInside h1 a {
	color: white
}
.topNav {
/*	position: absolute;*/
	left: 130px;
	top: 0;
	height: 25px;
	float: left;
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
	padding: 7px 0 8px 0;
	_cursor: pointer;
	_float: left;
	_position: relative;
}
.topNav .topNavItem a i {
	position: absolute;
	display: block;
	height: 25px;
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
            <h1><a title="快速插入">快速插入</a></h1>
            <ul class="topNav navMenu">
              <li class="topNavItem" _enable="{%host%}"><a title="博客地址"><u>{%host%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%post%}"><a title="文章存放目录"><u>{%post%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%category%}"><a title="文章分类"><u>{%category%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%author%}"><a title="文章作者"><u>{%author%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%year%}"><a title="文章发表年份"><u>{%year%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%month%}"><a title="文章发表月份"><u>{%month%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%day%}"><a title="文章发表日"><u>{%day%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%date%}"><a title="日期"><u>{%date%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%id%}"><a title="ID"><u>{%id%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%alias%}"><a title="别名"><u>{%alias%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%name%}"><a title="名称"><u>{%name%}</u><i></i></a></li>
              <li class="topNavItem" _enable="{%page%}"><a title="分页"><u>{%page%}</u><i></i></a></li>
              <li class="topNavItem" _enable="POST_FOLTER"><a><u>post</u><i></i></a></li>
			  <li class="topNavItem" _enable="POST_FOLTER"><a><u>archives</u><i></i></a></li>
            </ul>
          </div>
        </div>
        <script type="text/javascript">
		$(function(){
			var _h=$("#headWrap").html();
			$("#headWrap").remove();
			$("#header").before('<div id="headWrap" style="display:none">'+_h+'</div>');
			//$("#headWrap").show();
			$(".topNavItem a").attr("href","javascript:;").click(function(event){event.stopPropagation(); InsertText(_focus,$(this).text(),false)})
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
              <table width="100%" style="padding:0px;margin:0px;" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="30%"><p align="left"><b>·全局静态化选项</b><br>
                      <span class="note">&nbsp;&nbsp;使用伪静态前必须确认主机是否支持</span></p></td>
                  <td><p>
                      <label>
                        <input type="radio" value="STATIC" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="STATIC","checked=""checked""","")%>>
                        &nbsp;&nbsp;1.文章静态</label>
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <label>
                        <input type="radio" value="ACTIVE" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="ACTIVE","checked=""checked""","")%>>
                        &nbsp;&nbsp;2.全局动态</label>
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <label>
                        <input type="radio" value="REWRITE" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="REWRITE","checked=""checked""","")%>>
                        &nbsp;&nbsp;3.全局伪静态</label>
                    </p></td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·文章存放目录</b><br>
                      <span class="note">&nbsp;&nbsp;静态生成文章的目录,也就是{%post%}参数的值</span></p></td>
                  <td><p>
                      <input id="edtZC_STATIC_DIRECTORY" _enblist="POST_FOLTER" name="edtZC_STATIC_DIRECTORY" style="width:500px;" type="text" value="<%=ZC_STATIC_DIRECTORY%>">
                    </p></td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·文章的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_ARTICLE_REGEXT" _enblist="{%category%}{%author%}{%year%}{%month%}{%day%}{%id%}{%alias%}" name="edtZC_ARTICLE_REGEX" style="width:500px;" type="text" value="<%=ZC_ARTICLE_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a></p></td>
                </tr>
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td><p>
                      <label onclick="changeval(1,1)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;<%=BlogHost%><%=ZC_STATIC_DIRECTORY%>/alias.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,2)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;<%=BlogHost%><%=Year(Now)%>/<%=Month(Now)%>/alias.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,3)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;<%=BlogHost%>分类/alias.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,4)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;<%=BlogHost%><%=ZC_STATIC_DIRECTORY%>/alias/</label>
                    </p>
                    <p>
                      <label onclick="changeval(1,5)">
                        <input type="radio" name="radio">
                        &nbsp;&nbsp;<%=BlogHost%>分类/123/</label>
                    </p></td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·页面的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_PAGE_REGEX" _enblist="{%category%}{%author%}{%year%}{%month%}{%day%}{%id%}{%alias%}" name="edtZC_PAGE_REGEX" style="width:500px;" type="text" value="<%=ZC_PAGE_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a> </p></td>
                </tr>
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td>
					<p>
                      <label onclick="changeval(2,1)">
                        <input type="radio" name="radio2">
                        &nbsp;&nbsp;<%=BlogHost%>alias.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(2,2)">
                        <input type="radio" name="radio2">
                        &nbsp;&nbsp;<%=BlogHost%>alias/</label>
                    </p>
					<p>
                      <label onclick="changeval(2,3)">
                        <input type="radio" name="radio2">
                        &nbsp;&nbsp;<%=BlogHost%>3.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(2,4)">
                        <input type="radio" name="radio2">
                        &nbsp;&nbsp;<%=BlogHost%>3/</label>
                    </p>
				  </td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·首页分页的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_DEFAULT_REGEX" _enblist="{%page%}" name="edtZC_DEFAULT_REGEX" style="width:500px;" type="text" value="<%=ZC_DEFAULT_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a></p></td>
                </tr>
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td>
                    <%If BlogVersion >= 140808 Then%>
                    <p>
                      <label onclick="changeval(6,1)">
                      <input type="radio" name="radio6">
                      &nbsp;&nbsp;新版分页：<%=BlogHost%>page/2/</label>
                    </p>
                    <%End If%>
                    <p>
                      <label onclick="changeval(6,2)">
                      <input type="radio" name="radio6">
                      &nbsp;&nbsp;旧版分页：<%=BlogHost%>default_2.html</label>
                    </p>
                  </td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·分类页的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_CATEGORY_REGEX" _enblist="{%id%}{%alias%}"  name="edtZC_CATEGORY_REGEX" style="width:500px;" type="text" value="<%=ZC_CATEGORY_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                    </p></td>
                </tr>
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td>
                    <p>
                      <label onclick="changeval(3,1)">
                      <input type="radio" name="radio3">
                      &nbsp;&nbsp;<%=BlogHost%>categroy/alias/</label>
                    </p>
                    <p>
                      <label onclick="changeval(3,2)">
                      <input type="radio" name="radio3">
                      &nbsp;&nbsp;<%=BlogHost%>category/2/</label>
                    </p>
                    <p>
                      <label onclick="changeval(3,3)">
                      <input type="radio" name="radio3">
                      &nbsp;&nbsp;<%=BlogHost%>category-alias.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(3,3)">
                      <input type="radio" name="radio3">
                      &nbsp;&nbsp;<%=BlogHost%>category-2.html</label>
                    </p>
                  </td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·作者页的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_USER_REGEX" _enblist="{%id%}{%alias%}"  name="edtZC_USER_REGEX" style="width:500px;" type="text" value="<%=ZC_USER_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                    </p></td>
                </tr>
                <!--<tr style="display:none">-->
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td>
                  <p>
                      <label onclick="changeval(7,1)">
                        <input type="radio" name="radio7">
                        &nbsp;&nbsp;<%=BlogHost%>author/<%=BlogUser.FirstName%>/</label>
                    </p>
                    <p>
                      <label onclick="changeval(7,2)">
                        <input type="radio" name="radio7">
                        &nbsp;&nbsp;<%=BlogHost%>author/1/</label>
                    </p>
                    <p>
                      <label onclick="changeval(7,3)">
                        <input type="radio" name="radio7">
                        &nbsp;&nbsp;<%=BlogHost%>author-alias.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(7,4)">
                        <input type="radio" name="radio7">
                        &nbsp;&nbsp;<%=BlogHost%>author-1.html</label>
                    </p>
                    </td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·TAGS页的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_TAGS_REGEX"  _enblist="{%id%}{%alias%}{%name%}"  name="edtZC_TAGS_REGEX" style="width:500px;" type="text" value="<%=ZC_TAGS_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                      </label>
                    </p></td>
                </tr>
                <!--<tr style="display:none">-->
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td>
                    <p>
                      <label onclick="changeval(4,1)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;<%=BlogHost%>tags/name/</label>
                    </p>
                    <p>
                      <label onclick="changeval(4,2)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;<%=BlogHost%>tags/intro/</label>
                    </p>
                    <p>
                      <label onclick="changeval(4,3)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;<%=BlogHost%>tags/id/</label>
                    </p>
                    <p>
                      <label onclick="changeval(4,4)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;<%=BlogHost%>tags-name.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(4,5)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;<%=BlogHost%>tags-intro.html</label>
                    </p>
                    <p>
                      <label onclick="changeval(4,6)">
                        <input type="radio" name="radio4">
                        &nbsp;&nbsp;<%=BlogHost%>tags-id.html</label>
                    </p>
                  </td>
                </tr>
                <tr>
                  <td width="30%"><p align="left"><b>·日期页的URL配置</b><br>
                      <span class="note"></span></p></td>
                  <td><p>
                      <input id="edtZC_DATE_REGEX" _enblist="{%date%}"  name="edtZC_DATE_REGEX" style="width:500px;" type="text" value="<%=ZC_DATE_REGEX%>">
                      &nbsp;&nbsp;<a href="javascript:;" onClick="$(this).hide().parents('tr').next('tr').show();bmx2table()">显示系统预设..</a>
                      </label>
                    </p></td>
                </tr>
                <!--<tr style="display:none">-->
                <tr style="display:none">
                  <td width="30%"><p></p></td>
                  <td>
                    <p>
                      <label onclick="changeval(5,1)">
                        <input type="radio" name="radio5">
                        &nbsp;&nbsp;<%=BlogHost%>date/<%=Year(Now)%>-<%=Month(Now)%>/</label>
                    </p>
                    <p>
                      <label onclick="changeval(5,2)">
                        <input type="radio" name="radio5">
                        &nbsp;&nbsp;<%=BlogHost%>date-<%=Year(Now)%>-<%=Month(Now)%>.html</label>
                    </p>
                  </td>
                </tr>
              </table>
              <!--<p><span class="star">注意:开启文章及页面和分类页的Rewrite支持选项后,请查看"ReWrite规则"并应用在主机上方能生效.</span></p>-->
			  <hr/>
              <input type="hidden" name="edtZC_STATIC_MODE" id="edtZC_STATIC_MODE" value="<%=ZC_STATIC_MODE%>">
              <input name="" type="submit" class="button" value="保存">
            </form>
          </div>
        </div>
        <script type="text/javascript">
			var _focus=document;
			$(document).not($("#headWrap")).click(function (event){$('#headWrap').slideUp("fast");});
			$(document).ready(function(){
				enable("none");
				$("input[type='text']").click(function(event){
					w=$(document).outerWidth();
					l=$(this).offset().left-100;
					m=w-890>l?l:w-890;
					_focus=this.id;enable($(this).attr("_enblist"));
					$("#headWrap").css({'top':$(this).offset().top-25,'left':m});
					$('#headWrap').slideDown("fast");
					event.stopPropagation();
					})
				flashradio();
			});
			$(":radio[name='POST_STATIC']").live("click",function(){
				$("#edtZC_POST_STATIC_MODE").val($(this).val());

				$("#edtZC_STATIC_MODE").val($(this).val()=="STATIC"?"ACTIVE":$(this).val());
				if($(this).val()=="STATIC"){
					$("#edtZC_ARTICLE_REGEXT").val("{%host%}/{%post%}/{%alias%}.html");
					$("#edtZC_PAGE_REGEX").val("{%host%}/{%alias%}.html");
					$("#edtZC_DEFAULT_REGEX").val("{%host%}/catalog.asp");
					$("#edtZC_CATEGORY_REGEX").val("{%host%}/catalog.asp?cate={%id%}");
					$("#edtZC_USER_REGEX").val("{%host%}/catalog.asp?auth={%id%}");
					$("#edtZC_TAGS_REGEX").val("{%host%}/catalog.asp?tags={%alias%}");
					$("#edtZC_DATE_REGEX").val("{%host%}/catalog.asp?date={%date%}");
				};
				if($(this).val()=="ACTIVE"){
					$("#edtZC_ARTICLE_REGEXT").val("{%host%}/view.asp?id={%id%}");
					$("#edtZC_PAGE_REGEX").val("{%host%}/view.asp?id={%id%}");
					$("#edtZC_DEFAULT_REGEX").val("{%host%}/catalog.asp");
					$("#edtZC_CATEGORY_REGEX").val("{%host%}/catalog.asp?cate={%id%}");
					$("#edtZC_USER_REGEX").val("{%host%}/catalog.asp?auth={%id%}");
					$("#edtZC_TAGS_REGEX").val("{%host%}/catalog.asp?tags={%alias%}");
					$("#edtZC_DATE_REGEX").val("{%host%}/catalog.asp?date={%date%}");
				};
				if($(this).val()=="REWRITE"){
					changeval(1,1);
					changeval(2,1);
					changeval(3,1);
					changeval(4,1);
					changeval(5,1);
					if (blogversion>=140808)
						changeval(6,1);
					else
						changeval(6,2);
					changeval(7,1);
				};
				flashradio();
			});



      function bmx2table(){
          var class_=new Array("color2","color3","color4");
          var j=$("table tr:has(th):visible").addClass("color1");
            $("table").each(function(){
            if(j.length==0){class_[1]="color2";class_[0]="color3";}
            $(this).find("tr:not(:has(th)):visible:even").removeClass(class_[0]).addClass(class_[1]);
            $(this).find("tr:not(:has(th)):visible:odd").removeClass(class_[1]).addClass(class_[0]);
          })
          $("tr:not(:has(th))").mouseover(function(){$(this).addClass(class_[2])}).mouseout(function(){$(this).removeClass(class_[2])});
        };

		</script>


        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->