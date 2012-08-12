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
If CheckPluginState("GuestBook")=False Then Call ShowError(48)
BlogTitle="留言本"
Dim objConfig,a,b,c
c=0
Set objConfig=New TConfig
objConfig.Load("GuestBook")
If objConfig.Read("g")<>"" Then
	If CInt(objConfig.Read("g"))=0  then b=true
Else
	b=true
End If
Dim objRS
if b=false then
	if request.QueryString("act")="save" then
		if not isempty(request.form("id")) then
			a=request.form("id")
			Call CheckParameter(a,"int",0)
			Set objRs=Server.CreateObject("adodb.recordset")
			objRs.Open "SELECT [log_ID] FROM [blog_Comment] WHERE [log_ID]="&a,obJConn,1,3
			Do Until objRs.Eof
				objRs("log_ID")=CInt(objConfig.Read("g"))
				objRs.Update
				objRs.MoveNext
				c=c+1
			Loop
			Call BuildArticle(a,False,True)
			Call BuildArticle(CInt(objConfig.Read("g")),False,True)
			BlogReBuild_Comments
			Call SetBlogHint_Custom("已经成功将ID="&a&"的文章的"&c&"条评论迁移到ID="&CInt(objConfig.Read("g"))&"的页面中！")
			Set objRs=Nothing
		end if
	end if 
end if

Set objRs=objConn.Execute("SELECT [log_ID] FROM [blog_Comment] WHERE [log_ID]=0")
If Not objRs.Eof Then Call SetBlogHint_Custom("检测到有1.8的留言未升级！请在下面指定一个页面后点击“迁移留言本”将1.8的留言升级到2.0！")
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function zsx(){
	var a=$('#id').attr('value');
	if (a==""){
		a=0;
		$('#id').attr('value','0');
	};
	if(a!=0){
		return confirm("是否要将ID为"+a+"的文章的全部评论移动到ID为<%=objConfig.Read("g")%>的页面？")
	}
}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left">设定留言本页面</span></a><a href="b.asp"><span class="m-left m-now">迁移留言本</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script><% if b=false then%>
<form id="form1" name="form1" method="post" action="?act=save" onsubmit="return zsx()">
<label for="id">请选择迁移文章ID<a href="javascript:;" onclick="$('#id').attr('value',0);$('#form1').submit()">【点击这里从1.8升级】</a></label><input type="number" id="id" name="id" min="0" max="" style="width:100%" value="0"/>
<p>
<input name="" type="submit" class="button" value="保存"/></p>
</form><%else 
response.write "未指定留言本！"
end if%>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
