<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%
'On Error Resume Next
 %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->

<% 
Call System_Initialize() 

'检查非法链接
Call CheckReference("") 

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("dztaotao")=False Then Call ShowError(48)

BlogTitle="dztaotao - 查看/操作淘淘" 
%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

	<div id="divMain">
		<div class="divHeader"><%=BlogTitle%></div>
        <div id="ShowBlogHint"><%Call GetBlogHint()%></div>
			<div class="SubMenu">
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin.asp?a=list"><span class="m-left">淘淘管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin_cmt.asp?a=list&page=1"><span class="m-left">评论管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/setting.asp"><span class="m-left m-now">配置管理</span></a>
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/help.asp"><span class="m-left">帮助说明</span></a>
			</div>
	<div id="divMain2">
    

<form method="post" id="form2" name="form2">
<table border="1" width="98%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc">
<tr>
  <td colspan="4"><strong>基本设置</strong></td>
  </tr>
<%

	Call dztaotao_Initialize
	
	Dim strZC_DZTAOTAO_TITLE_VALUE
	strZC_DZTAOTAO_TITLE_VALUE=dztaotao_Config.Read("DZTAOTAO_TITLE_VALUE")
	strZC_DZTAOTAO_TITLE_VALUE=TransferHTML(strZC_DZTAOTAO_TITLE_VALUE,"[html-format]")
%>	
<tr>
<td width="10%">标题：</td>
<td colspan="3" width="90%"><input name="strZC_DZTAOTAO_TITLE_VALUE" type="text" id="is_release" value="<%=strZC_DZTAOTAO_TITLE_VALUE%>" />

</td>
</tr>	
<%
	
	Dim strZC_DZTAOTAO_RELEASE_VALUE
	strZC_DZTAOTAO_RELEASE_VALUE=dztaotao_Config.Read("DZTAOTAO_RELEASE_VALUE")
	strZC_DZTAOTAO_RELEASE_VALUE=TransferHTML(strZC_DZTAOTAO_RELEASE_VALUE,"[html-format]")
%>	
<tr>
<td width="10%">前台发布：</td>
<td colspan="3" width="90%"><input name="strZC_DZTAOTAO_RELEASE_VALUE" type="radio" id="is_release" value="5"<%if clng(strZC_DZTAOTAO_RELEASE_VALUE)=5 then response.write " checked"%> />游客
<input name="strZC_DZTAOTAO_RELEASE_VALUE" type="radio" id="no_release" value="1"<%if clng(strZC_DZTAOTAO_RELEASE_VALUE)=1 then response.write " checked"%> />管理员
</td>
</tr>	
<%	
	Dim strDZTAOTAO_PAGECOUNT_VALUE
	strDZTAOTAO_PAGECOUNT_VALUE=dztaotao_Config.Read("DZTAOTAO_PAGECOUNT_VALUE")
	strDZTAOTAO_PAGECOUNT_VALUE=TransferHTML(strDZTAOTAO_PAGECOUNT_VALUE,"[html-format]")
%>
<tr>
  <td width="10%">每页显示：</td>
  <td width="90%">
	<input name="strDZTAOTAO_PAGECOUNT_VALUE" type="text" id="page_count" value="<%=strDZTAOTAO_PAGECOUNT_VALUE%>" size="4" />条
  </td>
</tr>
<%
	
	Dim strDZTAOTAO_PAGEWIDTH_VALUE
	strDZTAOTAO_PAGEWIDTH_VALUE=dztaotao_Config.Read("DZTAOTAO_PAGEWIDTH_VALUE")
	strDZTAOTAO_PAGEWIDTH_VALUE=TransferHTML(strDZTAOTAO_PAGEWIDTH_VALUE,"[html-format]")
%>
<tr>
  <td>内容宽度：</td>
  <td><input name="strDZTAOTAO_PAGEWIDTH_VALUE" type="text" id="page_width" value="<%=strDZTAOTAO_PAGEWIDTH_VALUE%>" size="4" />
    px</td>
</tr>
<%
	Dim strDZTAOTAO_CHK_VALUE
	strDZTAOTAO_CHK_VALUE=dztaotao_Config.Read("DZTAOTAO_CHK_VALUE")
	strDZTAOTAO_CHK_VALUE=TransferHTML(strDZTAOTAO_CHK_VALUE,"[html-format]")
%>
<tr>
  <td colspan="2"><strong>内容审核</strong></td>
  </tr>
<tr>
  <td>发布内容：</td>
  <td><input name="strDZTAOTAO_CHK_VALUE" type="radio" id="taotao_chk1" value="4"<%if clng(strDZTAOTAO_CHK_VALUE)=4 then response.write " checked"%> />
    需要审核
      <input name="strDZTAOTAO_CHK_VALUE" type="radio" id="taotao_chk2" value="0"<%if clng(strDZTAOTAO_CHK_VALUE)=0 then response.write " checked"%> />
      直接发布</td>
</tr>
<%
	Dim strDZTAOTAO_CMTCHK_VALUE
	strDZTAOTAO_CMTCHK_VALUE=dztaotao_Config.Read("DZTAOTAO_CMTCHK_VALUE")
	strDZTAOTAO_CMTCHK_VALUE=TransferHTML(strDZTAOTAO_CMTCHK_VALUE,"[html-format]")
%>
<tr>
  <td>评论审核：</td>
  <td><input name="strDZTAOTAO_CMTCHK_VALUE" type="radio" id="cmt_chk1" value="4"<%if clng(strDZTAOTAO_CMTCHK_VALUE)=4 then response.write " checked"%> />
需要审核
  <input name="strDZTAOTAO_CMTCHK_VALUE" type="radio" id="cmt_chk2" value="0"<%if clng(strDZTAOTAO_CMTCHK_VALUE)=0 then response.write " checked"%> />
直接发布</td>
</tr>
<%	
	Dim strDZTAOTAO_CMTLIMIT_VALUE
	strDZTAOTAO_CMTLIMIT_VALUE=dztaotao_Config.Read("DZTAOTAO_CMTLIMIT_VALUE")
	strDZTAOTAO_CMTLIMIT_VALUE=TransferHTML(strDZTAOTAO_CMTLIMIT_VALUE,"[html-format]")
%>
<tr>
  <td>评论限制：</td>
  <td><input name="strDZTAOTAO_CMTLIMIT_VALUE" type="radio" id="cmt_limt1" value="1"<%if clng(strDZTAOTAO_CMTLIMIT_VALUE)=1 then response.write " checked"%> />
    评论一次
    <input name="strDZTAOTAO_CMTLIMIT_VALUE" type="radio" id="cmt_limt2" value="999"<%if clng(strDZTAOTAO_CMTLIMIT_VALUE)=999 then response.write " checked"%> />
    评论多次</td>
</tr>
<%
	Dim strZC_DZTAOTAO_ISIMG_VALUE
	strZC_DZTAOTAO_ISIMG_VALUE=dztaotao_Config.Read("DZTAOTAO_ISIMG_VALUE")
	strZC_DZTAOTAO_ISIMG_VALUE=TransferHTML(strZC_DZTAOTAO_ISIMG_VALUE,"[html-format]")


%>

<tr>
  <td colspan="2"><strong>扩展设置</strong></td>
  </tr>
<tr>
  <td>上传图片：</td>
  <td>
  <select name="strZC_DZTAOTAO_ISIMG_VALUE" id="is_img">
  <option value="0"<%if clng(strZC_DZTAOTAO_ISIMG_VALUE)=0 then response.write " selected"%>>不支持</option>
  <option value="1"<%if clng(strZC_DZTAOTAO_ISIMG_VALUE)=1 then response.write " selected"%>>支持</option>
  
<%
'dim is_i
'on error resume next
'set jpeg1=server.createobject("persits.jpeg")
'if err.number<>0 then
'Response.write "<option value=""0"" selected>不支持</option><option value=""1"">支持</option>"
'is_i = "<span style=""color:#f00;"">当空间不支持aspJPEG的时候就不能在前台发布图片</span>"
'else
'Response.write "<option value=""0"">不支持</option><option value=""1"" selected>支持</option>"
'End if
%>
    
    </select></td>
</tr>
<tr>
  <td>&nbsp;</td>
  <td><input type="submit" class="button" value=" 保存 " id="btnPost" onclick='document.getElementById("form2").action="savesetting.asp";' /></td>
</tr>

</table>
</form>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
