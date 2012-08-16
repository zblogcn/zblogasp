<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
If CheckPluginState("CustomMeta")=False Then Call ShowError(48)
BlogTitle="CustomMeta自定义数据字段"
Dim c
Set c=New TConfig
c.Load "CustomMeta"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->


<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
<div class="SubMenu"><a href="main.asp"><span class="m-left">文章页面自定义数据字段</span></a><a href="cate.asp"><span class="m-left m-now">分类自定义数据字段</span></a></div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>

<form id="form" name="form" method="post" action="save.asp">

<input type="hidden" name="edtZC_STATIC_MODE" id="edtZC_STATIC_MODE" value="<%=ZC_STATIC_MODE%>" />
<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
<tr><th width='40%'>自定义数据段名称</th><th width='40%'>注释</th><th width='20%'></th></tr>

<%
Dim m,i
Set m=New TMeta
m.LoadString=c.Read("CateMeta")
For i=LBound(m.Names)+1 To UBound(m.Names)
Response.Write "<tr><td><input style='margin:10px 10px;width:80%;' name='MetaName'  type='text' value='"&m.Names(i)&"' /></td><td><input style='margin:10px 10px;width:80%;' name='MetaNote'  type='text' value='"&m.GetValue(m.Names(i))&"' /></td><td align='center'><input name='' type='button' class='button' onclick='$(this).parent().parent().remove();return false;' value='删除'/></td></tr>"
Next
%>
<tr><td><input id="newMetaName" style='margin:10px 10px;width:80%;' name="MetaName"  type="text" value="" /><span class="star">(*)</span></td><td><input id="newMetaNote" style='margin:10px 10px;width:80%;' name="MetaNote"  type="text" value="" /></td><td width='10%' align='center'><input name="" type="submit" class="button" value="新建"/></td></tr>
</table>
<p><span class="note">自定义数据段名称必须是小写英文字母,数字和下划线_的组合</span></p>
<input name="" type="submit" class="button" value="保存"/>
</form>
<script type="text/javascript">
$(".button[value='新建']").click( function() {

if($("#newMetaName").val().toLowerCase().match(/[a-z0-9_]{1,30}/)==null){
  alert("自定义数据段名称必须是小写英文字母,数字和下划线_的组合");
  return false;
}


});
</script>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

