<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    33195@qq.com
'// 程序名称:    	Content Manage System
'// 开始时间:    	2011.03.26
'// 最后修改:    2012-08-08
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #INCLUDE FILE="../../C_OPTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_FUNCTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_LIB.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_BASE.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_EVENT.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_PLUGIN.ASP" -->
<!-- #INCLUDE FILE="../../PLUGIN/P_CONFIG.ASP" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
If CheckPluginState("YTCMS") = False Then Call ShowError(48)
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" rev="stylesheet" href="STYLE/YT.Style.css" type="text/css" media="screen">
<script language="javascript" type="text/javascript">
	var ZC_BLOG_HOST='<%=ZC_BLOG_HOST%>';
	var YT_CMS_XML_URL='<%=ZC_BLOG_HOST&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"%>';
	var isAlipay = <%=LCase(CheckPluginState("YTAlipay"))%>
</script>
<script language="javascript" src="Config.js" type="text/javascript"></script>
<script language="JavaScript" src="../../../ZB_SYSTEM/script/common.js" type="text/javascript"></script>
<script language="javascript" src="SCRIPT/YT.Lib.js" type="text/javascript"></script>
<script language="javascript" src="SCRIPT/YT.Interface.js" type="text/javascript"></script>
<script language="javascript" src="SCRIPT/YT.Main.js" type="text/javascript"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="box">
	<div id="headerWelcome"></div>
    <div id="tplList">
    	<div class="row"><img src="<%=ZC_BLOG_HOST%>ZB_USERS/PLUGIN/YTCMS/STYLE/IMAGES/loading.gif" class="loading" /></div>
    </div>
    <div id="footerWelcome"></div>
</div>
<div id="Panel"><span>x</span><div></div></div>
<div id="Template">
    <ul class="p20">
    <li></li>
    <li></li>
    <li><textarea></textarea></li>
    <li><input type="button" value="保存设置" /></li>
    </ul>
</div>
<div class="Block">
    <table width="100%" style="margin-top:0;" cellspacing="0" cellpadding="0" border="0">
        <tr class="color2">
            <td height="25" align="center">模块名称</td>
            <td align="center" width="20%">操作</td>
        </tr>
        <tr>
            <td></td>
            <td align="center"></td>
        </tr>
    </table>
</div>
<div class="TPL">
    <table width="100%" style="margin-top:0;" cellspacing="0" cellpadding="0" border="0">
        <tr class="color2">
            <td height="25" align="center">模板</td>
            <td align="center">绑定</td>
            <td align="center" width="10%">类型</td>
            <td align="center" width="20%">操作</td>
        </tr>
        <tr>
            <td></td>
            <td></td>
            <td align="center"></td>
            <td align="center"></td>
        </tr>
    </table>
</div>
<div class="Model">
	<div id="Step1">
        表名 <input type="text" />
        描述 <input type="text" />
        <input type="button" value="新增字段" id="Add" /><input type="button" id="Next" value="下一步" />
        <label><select>
        <option value="0">默认</option>
        <option value="1">支付宝[即时交易]</option>
        <option value="2">支付宝[担保交易]</option>
        </select></label>
	<div>
        字段 <input type="text" />
        描述 <input type="text" />
        默认值 <input type="text" />
        属性 <select>
        <option value="INT">数字</option>
        <option value="VARCHAR">文本</option>
        <option value="TEXT">备注</option>
        <option value="DATETIME">时间</option>
        <option value="COUNTER(1,1)">自动编号</option>
        </select>
        类型 <select>
        <option value="text">单行文本框</option>
        <option value="checkbox">多选</option>
        <option value="select">下拉框</option>
        <option value="textarea">多行文本框</option>
        </select>
    </div>
    </div>
    <div id="Step2">
    	<%GetCategory()%>
        <select multiple="multiple" size="20">
		<%
        Dim Category
        For Each Category in Categorys
            If IsObject(Category) And Not isEmpty(Category.Name) Then
            %>
            <option value="<%= Category.ID %>"><%= TransferHTML(Category.Name,"[html-format]") %></option>
            <%
            End If
        Next
        %>
        </select>
        <input type="button" value="保存设置" />
    </div>
</div>
<div class="Model">
      <table width="100%" style="margin-top:0;" cellspacing="0" cellpadding="0" border="0">
        <tr class="color2">
          <td height="25" align="center">表</td>
          <td align="center">描述</td>
          <td align="center">绑定</td>
          <td align="center">所属</td>
          <td align="center">操作</td>
        </tr>
        <tr>
          <td align="center"></td>
          <td align="center"></td>
          <td align="center"></td>
          <td align="center"></td>
          <td align="center"></td>
        </tr>
      </table>
</div>
<div id="currPanel">
	<ul>
        <li class="tree"><a href="#">模块</a>
            <div>
                <ul>
                    <li><a href="#CBLOCK">新建</a></li>
                    <li><a href="#MBLOCK">管理</a></li>
                </ul>
            </div>
        </li>
        <li class="tree"><a href="#">模型</a>
            <div>
                <ul>
                    <li><a href="#CMODEL">新建</a></li>
                    <li><a href="#MMODEL">管理</a></li>
                </ul>
            </div>
        </li>
        <li class="tree"><a href="#">子块</a>
            <div>
                <ul>
                    <li class="diyBlock"><a href="#CTPL">绑定</a></li>
                    <li><a href="#MTPL">管理</a></li>
                </ul>
            </div>
        </li>
    </ul>
</div>
<script type="text/javascript">ActiveLeftMenu("aYTCMSMng");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>