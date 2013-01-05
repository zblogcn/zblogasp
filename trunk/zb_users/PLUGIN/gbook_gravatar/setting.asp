<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    大猪
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

If CheckPluginState("gbook_gravatar")=False Then Call ShowError(48)

BlogTitle="最新评论 - 查看/操作" 
%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

	<div id="divMain">
		<div class="divHeader"><%=BlogTitle%></div>
        <div id="ShowBlogHint"><%Call GetBlogHint()%></div>
			
	<div id="divMain2">
    

<form method="post" id="form2" name="form2">
<table border="1" width="98%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc">
<tr>
  <td colspan="4" height="35"><strong>基本设置</strong></td>
  </tr>
<%

	Call gbook_gravatar_Initialize
	
'	Dim str_DZ_IDS_VALUE
'	str_DZ_IDS_VALUE=gbook_gravatar_Config.Read("DZ_IDS_VALUE")
'	str_DZ_IDS_VALUE=TransferHTML(str_DZ_IDS_VALUE,"[html-format]")
%>	
<tr style="display:none">
<td width="10%">文章ID：</td>
<td colspan="3" width="90%"><input name="str_DZ_IDS_VALUE" type="text" id="str_DZ_IDS_VALUE" value="<%'=str_DZ_IDS_VALUE%>" />

</td>
</tr>	
<%
	
	Dim str_DZ_AVATAR_VALUE
	str_DZ_AVATAR_VALUE=gbook_gravatar_Config.Read("DZ_AVATAR_VALUE")
	str_DZ_AVATAR_VALUE=TransferHTML(str_DZ_AVATAR_VALUE,"[html-format]")
%>	
<tr>
<td width="10%">默认头像：</td>
<td colspan="3" width="90%"><input name="str_DZ_AVATAR_VALUE" type="text" id="str_DZ_AVATAR_VALUE" value="<%=str_DZ_AVATAR_VALUE%>" />
</td>
</tr>	
<%	
	Dim str_DZ_WH_VALUE
	str_DZ_WH_VALUE=gbook_gravatar_Config.Read("DZ_WH_VALUE")
	str_DZ_WH_VALUE=TransferHTML(str_DZ_WH_VALUE,"[html-format]")
%>
<tr>
  <td width="10%">头像宽高：</td>
  <td width="90%">
	<input name="str_DZ_WH_VALUE" type="text" id="str_DZ_WH_VALUE" value="<%=str_DZ_WH_VALUE%>" size="4" />PX
  </td>
</tr>
<%
	
	Dim str_DZ_TITLE_VALUE
	str_DZ_TITLE_VALUE=gbook_gravatar_Config.Read("DZ_TITLE_VALUE")
	str_DZ_TITLE_VALUE=TransferHTML(str_DZ_TITLE_VALUE,"[html-format]")
%>
<tr>
  <td>标题长度：</td>
  <td><input name="str_DZ_TITLE_VALUE" type="text" id="str_DZ_TITLE_VALUE" value="<%=str_DZ_TITLE_VALUE%>" size="4" />
    </td>
</tr>

<%
	
	Dim str_DZ_COUNT_VALUE
	str_DZ_COUNT_VALUE=gbook_gravatar_Config.Read("DZ_COUNT_VALUE")
	str_DZ_COUNT_VALUE=TransferHTML(str_DZ_COUNT_VALUE,"[html-format]")
%>
<tr>
  <td>调用条数：</td>
  <td><input name="str_DZ_COUNT_VALUE" type="text" id="str_DZ_COUNT_VALUE" value="<%=str_DZ_COUNT_VALUE%>" size="4" />
    </td>
</tr>
<%
	
	Dim str_DZ_ISREPLY
	str_DZ_ISREPLY=gbook_gravatar_Config.Read("DZ_ISREPLY")
	str_DZ_ISREPLY=TransferHTML(str_DZ_ISREPLY,"[html-format]")
%>

<tr>
  <td height="29">显示回复：</td>
  <td><label>
        <input type="radio" name="str_DZ_ISREPLY" value="1" id="isreply_1" />
        显示</label>
    <label>
      <input type="radio" name="str_DZ_ISREPLY" value="0" id="isreply_0" />
        不显示</label>
         <script language="javascript" type="text/javascript">document.getElementById('isreply_<%=str_DZ_ISREPLY%>').checked=true;</script>
    </td>
</tr>
<%
	
	Dim str_DZ_USERIDS
	str_DZ_USERIDS=gbook_gravatar_Config.Read("DZ_USERIDS")
	str_DZ_USERIDS=TransferHTML(str_DZ_USERIDS,"[html-format]")
%>

<tr>
  <td height="29">不显示用户：</td>
  <td><input name="str_DZ_USERIDS" type="text" id="str_DZ_USERIDS" value="<%=str_DZ_USERIDS%>" size="65" />
    </td>
</tr>

<%
 	Dim str_DZ_STYLE_VALUE
 	str_DZ_STYLE_VALUE=gbook_gravatar_Config.Read("DZ_STYLE_VALUE")
 	str_DZ_STYLE_VALUE=TransferHTML(str_DZ_STYLE_VALUE,"[html-format]")
 %>
 <tr>
   <td height="29">外观样式：</td>
   <td>
    <!--<p>当前样式：<%=str_DZ_STYLE_VALUE%></p>-->
    <!--<p>-->
    <input type="radio" name="str_DZ_STYLE_VALUE" id="str_DZ_STYLE_VALUE_1" value="1" />&nbsp;<label for="str_DZ_STYLE_VALUE_1">昵称+留言</label>&nbsp;&nbsp;
    <input type="radio" name="str_DZ_STYLE_VALUE" id="str_DZ_STYLE_VALUE_2" value="2" onClick="javascript:document.getElementById('str_DZ_WH_VALUE').value='16';" />&nbsp;<label for="str_DZ_STYLE_VALUE_2">小头像+留言</label>&nbsp;&nbsp;
    <input type="radio" name="str_DZ_STYLE_VALUE" id="str_DZ_STYLE_VALUE_3" value="3"onclick="javascript:document.getElementById('str_DZ_WH_VALUE').value='32';" />&nbsp;<label for="str_DZ_STYLE_VALUE_3">大头像+昵称+留言</label>
    <!--</p>-->
    <script language="javascript" type="text/javascript">document.getElementById('str_DZ_STYLE_VALUE_<%=str_DZ_STYLE_VALUE%>').checked=true;</script>
     </td>
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