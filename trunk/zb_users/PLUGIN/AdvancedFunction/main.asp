<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%'UI设计部分有参考coolmud的列表插件%>
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
If CheckPluginState("AdvancedFunction")=False Then Call ShowError(48)
BlogTitle="增强侧栏"
Dim subCate
%>
<%
init()%>
<script language="javascript" runat="server">
	
	function init(){
		advancedfunction.init();
		advancedfunction.run("随机文章,访问最多文章,本月最热文章,本年最热文章,分类最热文章,评论最多文章,本月评论最多,本年评论最多,分类评论最多,分类");
		
	}
	</script>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <!--<div class="SubMenu"></div>-->
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <form name="form1" id="form1" action="save.asp?act=save" method="post" enctype="application/x-www-form-urlencoded">
            <p>这，是一款侧栏增强插件。走过路过请不要错过，这里的侧栏很精彩。</p>
            <p>强烈建议不用的侧栏条数设置为0，这样可以避免浪费资源。</p>
            <p>插件只负责为您生成侧边栏，剩余的管理还需要您手动通过“侧栏管理”管理。停用插件后，原有侧栏依然可用，但是不再自动更新。</p>
            <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
              <tr>
                <th width='20%'><p align="center">侧栏选择</p></th>
                <th width='5%'><p align="center">设置条数</p></th>
                <th><p align="center">说明</p></th>
              </tr>

              <tr>
                <td><b>
                  <label for="m_访问最多文章">访问最多文章</label>
                  </b></td>
                <td><p>
                    <input name="m_访问最多文章" type="text" id="m_访问最多文章" size="5" value="<%=advancedfunction.functions.readconfig("访问最多文章")%>" />
                  </p></td>
                <td>显示访问人数最多的文章。</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_本月最热文章">本月最热文章</label>
                  </b></td>
                <td><p>
                    <input name="m_本月最热文章" type="text" id="m_本月最热文章" size="5" value="<%=advancedfunction.functions.readconfig("本月最热文章")%>" />
                  </p></td>
                <td>显示30天内访问人数最多的文章。</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_本年最热文章">本年最热文章</label>
                  </b></td>
                <td><p>
                    <input name="m_本年最热文章" type="text" id="m_本年最热文章" size="5" value="<%=advancedfunction.functions.readconfig("本年最热文章")%>" />
                  </p></td>
                <td>显示365天内访问人数最多的文章。</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_分类最热文章">分类最热文章</label>
                  </b></td>
                <td><p>
                    <input name="m_分类最热文章" type="text" id="m_分类最热文章" size="5" value="<%=advancedfunction.functions.readconfig("分类最热文章")%>" />
                  </p></td>
                <td>显示某个分类访问人数最多的文章。同时该配置项也控制某个分类的最新文章显示数量</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_评论最多文章">评论最多文章</label>
                  </b></td>
                <td><p>
                    <input name="m_评论最多文章" type="text" id="m_评论最多文章" size="5" value="<%=advancedfunction.functions.readconfig("评论最多文章")%>" />
                  </p></td>
                <td>显示评论最多的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_本月评论最多">本月评论最多</label>
                  </b></td>
                <td><p>
                    <input name="m_本月评论最多" type="text" id="m_本月评论最多" size="5" value="<%=advancedfunction.functions.readconfig("本月评论最多")%>" />
                  </p></td>
                <td>显示30天内评论最多的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_本年评论最多">本年评论最多</label>
                  </b></td>
                <td><p>
                    <input name="m_本年评论最多" type="text" id="m_本年评论最多" size="5" value="<%=advancedfunction.functions.readconfig("本年评论最多")%>" />
                  </p></td>
                <td>显示365天内评论最多的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_分类评论最多">分类评论最多</label>
                  </b></td>
                <td><p>
                    <input name="m_分类评论最多" type="text" id="m_分类评论最多" size="5" value="<%=advancedfunction.functions.readconfig("分类评论最多")%>" />
                  </p></td>
                <td>显示30天内评论最多的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_随机文章">随机文章</label>
                  </b></td>
                <td><p>
                    <input name="m_随机文章" type="text" id="m_随机文章" size="5" value="<%=advancedfunction.functions.readconfig("随机文章")%>" />
                  </p></td>
                <td>随机文章（以JS方式调用，但有HTML代码存在。）</td>
              </tr>

              <%
				For Each subCate In Categorys
					If IsObject(subCate) Then
						if advancedfunction.cls.config.Read("分类_"&subCate.ID) <> "" then
						  Response.Write "<tr>"
						  Response.Write "<td>[分类]<b>"&subCate.Name&"</b></td>"
						  Response.Write "<td><p><input class='text' name='m_分类_"&subCate.ID&"' type='text' size='5' value='"&advancedfunction.cls.config.Read("分类_"&SubCate.ID&"")&"' /></p></td>"
						  Response.Write "<td><input type=""button"" class=""button"" value=""删除"" onclick='location.href=""save.asp?act=del&id="&subCate.ID&"""'/></td>"
						  Response.Write "</tr>"
						  Response.Write vbCrlf
						end if
					End If
				Next
			  %>
              <tr>
                <td><b>
                  <label for="添加新分类列表">添加新分类列表</label>
                  </b></td>
                <td></td>
                <td>
				<%
				For Each subCate In Categorys
					If IsObject(subCate) Then
						Response.Write "<p><label><input type=""radio"" name=""newCategory"" value="""&subCate.ID&"""/>"&subCate.Name&"</label></p>"
						Response.Write vbCrlf&"				"
					End If
				Next
				%>
                </td>
              </tr>
            </table>
            <input type="submit" value="保存" class="button"/>
            </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->