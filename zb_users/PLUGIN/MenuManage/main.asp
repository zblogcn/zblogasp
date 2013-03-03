<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Dim l

l=FilterSQL(Request.QueryString("name"))

Call System_Initialize()

Call CheckReference("")

If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("MenuManage")=False Then Call ShowError(48)

BlogTitle="菜单管理器"


Call MenuManage.page.top()
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="MenuManage.css" type="text/css" media="screen"/>
<script type="text/javascript" src="MenuManage.js"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=MenuManage.page.submenu(0)%> </div>
          <div id="divMain2"> <script type="text/javascript">ActiveTopMenu("zbdk");</script>
            <div class="divMenuManage">
            <form method="post" action="?act=save" enctype="application/x-www-form-urlencoded">
              <div class="divMenuManagenav" name="tree" id="tree">
                <ul>
                  <%=MenuManage.page.sidebar.all()%>
                </ul>
              </div>
              <div id="content" class="divMenuManageContent">
                <div class="divMenuManageContentBody">请选择</div>
              </div>
              <div class="clear"></div>
              </div>
              <input type="hidden" id="sort" value="sort" />
            </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<script type="text/javascript" language-"javscript">
$(document).ready(function() {
	/*$("#tree ul li");*/
	$("#leftmenu").sortable({stop:function(){
		var s="";
		$("#leftmenu li a").each(function(index,element){
			s+=$(element).attr("id")+"|"
		});
		$.post("main.asp?act=savebar",{"bar":s.substr(0,s.length-1)});
	}}).disableSelection();
	$("#leftmenu li a").click(function(e,o){
	});
	

});
</script>
<script language="javascript" runat="server" >
MenuManage["page"]={
	"top":function(){
		MenuManage.page.content();
	},
	"sidebar":{
		"used":function(){
			var k="",o="",d="",s={},p=MenuManage.c.Read("config").split("|");
			for(var i=0;i<p.length;i++){
				s=eval("("+MenuManage.c.Read(p[i])+")");
				d+="<li id='"+s.liid+"' class='_li'><a id='"+s.aid+"' class='li_a' href='javascript:;'>"+s.name+(s.custom?"":"(无法删除)")+"</a></li>"
			}
			return d
		},
		"all":function(){
			var n=new VBArray(MenuManage.c.Meta.Names).toArray(),v=new VBArray(MenuManage.c.Meta.Values).toArray();
			var k="",o="",d="",s={};
			for(var i=1;i<n.length;i++){
				k=unescape(n[i]),o=unescape(v[i]);
				if(!/^(major|config|default)$/.test(k)){
					s=eval("("+o+")");
					d+="<li id='"+s.liid+"' class='_li'><a id='"+s.aid+"' class='li_a' href='javascript:;'>"+s.name+(s.custom?"":"(无法删除)")+"</a></li>"
				}
			}
			return d
		}
	},
	"content":function(){
		switch(Request.QueryString("act").Item){
			case "sidebar":
			break;
			
			case "content":
			break;
			
			case "edit":
			break;
			
			case "delete":
			break;
			
			case "savebar":
				MenuManage.c.Write("config",Request.Form("bar").Item);
				MenuManage.c.Save()
				Response.Write("{'success':'ok'}")
				Response.End();
			break;
		}
	},
	"submenu":function(i){
		var data=[
			{"data":"首页","url":"main.asp","css":""},
			{"data":"新建","url":"javascript:create()","css":""},
			{"data":"刷新","url":"javascript:location.reload()","css":""}
		]
		var st="";
		for(var lst in data){
			st+=MakeSubMenu(data[lst].data,data[lst].url,data[lst].css+((lst==i)?" m-now":""),false);
		}
		return st;

	}
}
</script>