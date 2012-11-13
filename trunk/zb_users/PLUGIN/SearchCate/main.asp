<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
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
If CheckPluginState("SearchCate")=False Then Call ShowError(48)
BlogTitle="按分类搜索"
Dim Config
Set Config=New TConfig
Config.Load "SearchCate"
Select Case Request.QueryString("act")
	Case "save"
		'Response.Write Join(Split(Request.Form,"&"),vbCrlf)
		Dim i
		For Each i In Request.Form
			Config.Write i,Request.Form(i)
		Next
		Config.Save
		Dim F
		Call GetFunction()
		Set F=Functions(FunctionMetas.GetValue("searchpanel"))
		F.Content=Request.Form("htmcode")
		F.Save
		Call ClearGlobeCache
		Call LoadGlobeCache
		Call BlogRebuild_Default
		Response.Write "保存并写入侧栏成功！"
		Response.End()
End Select
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
.hide{display:none}
</style>
<script type="text/javascript" src="jquery.form.js"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">设置</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
    <p>事情完工以后，本插件你爱删就删，反正也不会怎么样。</p>
    <form id="form1" name="form1" method="post" action="?act=save" onsubmit="return false">
    <table width="100%" border="1">
      <tr height="32px">
        <th scope="col" style="width:150px">项目</th>
        <th scope="col">配置</th>
      </tr>
      <tr height="32px">
        <th scope="row">是否开启分类搜索</th>
        <td>
          <input type="radio" name="opencate" id="opencate" value="1" <%=IIf(Config.Read("opencate")=1,"checked=""checked""","""")%> />
          <label for="opencate">打开</label>
          <input type="radio" name="opencate" id="opencate1" value="0" <%=IIf(Config.Read("opencate")<>1,"checked=""checked""","""")%> />
          <label for="opencate1">不打开</label>
        </td>
      </tr>
      <tr height="32px" class="cate" name="cate">
        <th scope="row">分类显示模板</th>
        <td><select name="show" id="showl">
        <option value="1" <%=IIf(Config.Read("show")=1,"selected=""selected""","""")%> >下拉框</option>
        <option value="2" <%=IIf(Config.Read("show")=2,"selected=""selected""","""")%> >复选框</option>
        <option value="3" <%=IIf(Config.Read("show")>2,"selected=""selected""","""")%> >自定义</option>
        </select>
        <input name="show_template" id="show_template" type="text" value="<%=IIf(Config.Read("show_template")<>"",TransferHTML(Config.Read("show_template"),"[html-format]"),"&lt;option value=&quot;&lt;!cateid!&gt;&quot;&gt;&lt;!catename!&gt;&lt;/option&gt;")%>" style="width:80%" onchange='$("#showl").val(3)'/>
        </td>
      </tr>
      <tr height="32px" class="cate" name="cate">
        <th scope="row">显示哪些分类</th>
        <td id="cates">				<%
				Dim subCate,ins
				ins=Config.Read("Cate")'0, 1, 2
				For Each subCate In Categorys
					If IsObject(subCate) Then
						Response.Write "<p><label><input type=""checkbox"" name=""Cate"" _cateid="""&subCate.ID&""" _catename="""&subCate.Name&""" id=""Cate"&subCate.ID&""" value="""&subCate.ID&""" onchange=""Run()"" "&IIf(InStr(ins,subCate.ID&", ")>0 Or InStr(ins,", "&subCate.ID)>0 Or ins=CStr(subCate.ID),"checked=""checked""","")&" />"&subCate.Name&"</label></p>"
						Response.Write vbCrlf&"				"
					End If
				Next
				%></td>
      </tr>
      <tr height="32px">
        <th scope="row">得到HTML代码</th>
        <td><label for="htmcode"></label>
          <textarea name="htmcode" id="htmcode" cols="45" rows="5" style="width:80%"><%=TransferHTML(Config.Read("htmcode"),"[textarea]")%></textarea></td>
      </tr>
    </table>
    <p>&nbsp;</p>
    <input type="submit" value="保存配置并写入到侧栏" class="button" id="submit"/>
    </form>
    <p>&nbsp;</p>
  </div>
</div>
<script type="text/javascript">
$("[name=opencate]").click(function(){
	if($("[name=opencate]:checked").val()==0){$("tr[name=cate]").hide()}else{$("tr[name=cate]").show()}
	Run()
	})
$("#showl").change(function() {
    switch ($(this).val()){
		case "1":$('#show_template').val('<option value="<!cateid!>" ><!catename!></option>');break;
		case "2":$('#show_template').val('<input value="<!cateid!>" name="cate" id="cate<!cateid!>" type="radio"/><label for="cate<!cateid!>"><!catename!></label>');break;
		default:$('#show_template').val('');break;
	};
	Run();
});
function Run(){
	var j=$("#htmcode");
	var _1="<form method=\"post\" action=\"<#ZC_BLOG_HOST#>/zb_system/cmd.asp?act=Search\"><input type=\"text\" name=\"edtSearch\" id=\"edtSearch\""+
	" size=\"12\" /> <input type=\"submit\" value=\"<#ZC_MSG087#>\" name=\"btnPost\" id=\"btnPost\" /></form>";
	var _6="<select name=\"cate\" style=\"width:50%\">";
	var _7="</select>  ";
	var _2="<form methpd=\"get\" action=\"<#ZC_BLOG_HOST#>/search.asp\"><input type=\"text\" name=\"q\" id=\"q\""+
	" style=\"width:50%\" />"
	var _3="<input type=\"submit\" value=\"<#ZC_MSG087#>\" name=\"btnPost\" id=\"btnPost\" /></form>";
	var _4=$("#show_template").val()
	if($("[name=opencate]:checked").val()==0){j.val(_1);return false}
	else{
		var _5=$("#cates input:checked");
		if(_5.length==0){j.val(_1);return false}
		var strtemp="",obj;
		for(var i=0;i<_5.length;i++){
			obj=$(_5[i]);
			strtemp+=_4.replace(/<!catename!>/g,obj.attr("_catename")).replace(/<!cateid!>/g,obj.attr("_cateid"));
		}
		j.val(_2+($("#showl").val()=="1"?_6:"")+strtemp+($("#showl").val()=="1"?_7:"")+_3);
	}
}
$("#submit").click(function(){
	$("#form1").ajaxForm(function(s){alert(s)});
})
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
