<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    33195@qq.com
'// 程序名称:    YT.Build
'// 开始时间:    	2011.03.26
'// 最后修改:    2012-08-08
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #INCLUDE FILE="../../C_OPTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_FUNCTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_LIB.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_BASE.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_EVENT.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_PLUGIN.ASP" -->
<!-- #INCLUDE FILE="../../PLUGIN/P_CONFIG.ASP" -->
<script language="javascript" type="text/javascript" runat="server">
	var Cmd={
		exec:function(s){return eval('('+s+')');}
	};
</script>
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("YTBuild")=False Then Call ShowError(48)
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script language="javascript" src="Script/YT.Build.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
$(document).ready(function(){
	YT.Default();
	YT.Catalog('bc','null',null);
	YT.View('bcv','c','Cate');
	YT.Catalog('bcc','c','Cate');
	YT.View('buv','u','Auth');
	YT.Catalog('buc','u','Auth');
	YT.View('btv','t','Tags');
	YT.Catalog('btc','t','Tags');
	YT.View('bdv','d','Date');
	YT.Catalog('bdc','d','Date');
	$('#Close').click(function(){YT.Close();});	
	$('input[name="buildPanel"]').click(function(){
		$('#Close').click();
		$('input[name="buildPanel"]').each(function(){$('#'+$(this).val()).hide();});
		$('#'+$(this).val()).show();	
	})
	$('input[name="buildPanel"]').eq(0).click();
/*	var st,i=0;
	$('#post').click(function(){
		var obj=[];	  
		$('#noobPanel').find('input[type="checkbox"]').each(function(){
			var s=$(this).parent().parent().find('td').eq(0).text('');
			if(this.checked){
				obj.push({
					event:$(this).val(),
					text:$(this).parent().text(),
					load:setInterval(function(){
						s.text().length>3?s.text(''):s[0].innerHTML+=".";	
					},100),
					object:s
				});
			}												   
		});
		i=0;
		YT.Thread[0]=-1;
		st=setInterval(function(){
			sThread(obj);
		},1000);
	});
	function sThread(a){
		if(i<a.length){
			if(YT.CompleteThread()){
				eval(a[i].event);
				clearInterval(a[i].load);
				a[i].object.text(' √');
				i++;
			}
		}else{
			clearInterval(st);
		}
	}*/
});
</script>
<style type="text/css">
	#Panel {display:none; position:absolute; z-index:9999; top:0; width:100%; left:0; background:#fff;}
	#Template,#Te {text-align:left; margin:1px; padding:5px; display: none; background:#000; clear:both;}
	#Status {position:relative; text-align:left;}
	#Status font {color:#FFF; position:relative; z-index:99999;}
	#Status div {width:0%; color:#FFF; background:#666; position:absolute; height:1.8em; left:0;}
	#UserShow,.UserShow {line-height:22px; color:#FFF;}
	select#u,select#c,select#t,select#d {width:98%;height:110px;}
	.but input {width:49%;height:100px;float:left;}
	.but2 input {width:20%;height:50px;}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader">YT.Build</div>
  <div class="SubMenu"> <a href="YT.Panel.asp"><span class="m-left m-now">控制面板</span></a><a href="YT.Config.asp"><span class="m-left">系统配置</span></a>
  </div>
  <div id="divMain2">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20%" align="right">界面</td>
    <td width="80%"> <label> <input name="buildPanel" checked="checked" value="noobPanel" type="radio" />菜鸟</label>
 <label>  <input name="buildPanel" checked="checked" value="expertPanel" type="radio" />专家</label></td>
    </tr>
    </table><br />
  <div id="noobPanel" style="display:none">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="but2">
    <tr>
    <td width="20%" align="right"></td>
    <td width="80%"><input type="button" value="首页" onclick="$('#default').click()" /></td>
    </tr>
    <tr>
    <td align="right"></td>
    <td><input type="button" value="文章页" onclick="$('#bcv').click()" /></td>
    </tr>
    <tr>
    <td align="right"></td>
    <td><input type="button" value="分类列表" onclick="$('#bcc').click()" /></td>
    </tr>
    <tr>
    <td align="right"></td>
    <td><input type="button" value="作者列表" onclick="$('#buc').click()" /></td>
    </tr>
    <tr>
    <td align="right"></td>
    <td><input type="button" value="TAG列表" onclick="$('#btc').click()" /></td>
    </tr>
    <tr>
    <td align="right"></td>
    <td><input type="button" value="文章归档列表" onclick="$('#bdc').click()" /></td>
    </tr>
    <tr>
    <td align="right"></td>
    <td><input type="button" value="首页日志列表" onclick="$('#bc').click()" /></td>
    </tr>
<!--    <tr>
    <td align="right"></td>
    <td><input type="button" value="发布" class="button" id="post" /></td>
    </tr>-->
    </table>
  </div>
  <div id="expertPanel" style="display:none">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
    <td width="20%" class="but" align="right">
    <input type="button" id="bcc" value="分类列表" />
    <input type="button" id="bcv" value="分类文章" />
    </td>
    <td width="80%">
    <select id="c" multiple="multiple">
    <%
            Call GetCategory()
            Dim Category
            For Each Category in Categorys
                If IsObject(Category) And Category.ID <> 0 Then
                %>
                <option value="<%= Category.ID %>"><%= TransferHTML(Category.Name,"[html-format]") %></option>
                <%
                End If
            Next
    %>
    </select>
    <a href="#" id="default" style="display:none">首页</a>
    <a href="#" id="bc" style="display:none">BLOG</a>
    <select id="null" multiple="multiple" style="display:none;"><option value="0">BLOG首页日志列表</option></select>
    </td>
    </tr>
    <tr>
    	<td class="but" align="right">
        <input type="button" id="buc" value="作者列表" />
    	<input type="button" id="buv" value="作者文章" />
        </td>
        <td>
        <select id="u" multiple="multiple">
        <%
                Call GetUser()
                Dim User
                For Each User in Users
                    If IsObject(User) Then
                    %>
                    <option value="<%= User.ID %>"><%= TransferHTML(User.Name,"[html-format]") %></option>
                    <%
                    End If
                Next
        %>
        </select>
        </td>
    </tr>
    <tr>
    	<td class="but" align="right">
        <input type="button" id="btc" value="TAG列表" />
    	<input type="button" id="btv" value="TAG文章" />
        </td>
        <td>
        <select id="t" multiple="multiple">
        <%
        Call GetTags()
            Dim objRS
            Set objRS=objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] ORDER BY [tag_Name] ASC")
            If (Not objRS.bof) And (Not objRS.eof) Then
                Do While Not objRS.eof
                    %>
                    <option value="<%= Tags(objRS("tag_ID")).ID %>"><%= TransferHTML(Tags(objRS("tag_ID")).Name,"[html-format]") %></option>
                    <%
                    objRS.MoveNext
                Loop
            End If
            objRS.Close
            Set objRS=Nothing
        %>
        </select>
        </td>
    </tr>
    <tr>
    	<td class="but" align="right">
        <input type="button" id="bdc" value="归档列表" />
    	<input type="button" id="bdv" value="归档文章" />
        </td>
        <td>
        <select id="d" multiple="multiple">
        <%
        Dim j
        For Each j In new YTStatic.GetdtmYM
            If IsDate(j) Then
                %>
                <option value="<%= j %>"><%= j %></option>
                <%
            End If
        Next
        %>
        </select>
        </td>
    </tr>
    </table>
  </div>
    <div id="Panel">
    <table width="500" align="center" class="tableBorder" border="0" cellspacing="0" cellpadding="0">
        <tr class="color1">
            <td colspan="3" class="color1"><span id="title"></span><a style="float:right; cursor:pointer;" id="Close">关闭</a></td>
        </tr>
        <tr>
            <td colspan="3" id="content"></td>
        </tr>
    </table>
    </div>
    <div id="Template">
        <div id="Status">
            <div></div>
            <font></font>
        </div>
        <div id="UserShow" class="UserShow"></div>
    </div>
</div>
</div>
<script type="text/javascript">ActiveLeftMenu("aYTBuildMng");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->