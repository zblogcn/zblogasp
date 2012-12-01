<%@ CODEPAGE=65001 %>
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
On Error Resume Next
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

If CheckPluginState("Totoro")=False Then Call ShowError(48)

BlogTitle="TotoroⅢ（基于TotoroⅡ的Z-Blog的评论管理审核系统增强版）"

Dim o,m,s,t,h,n,n1
If Request.QueryString("type")="test" Then
	On Error Resume Next
	Set o=New RegExp
	t=Request.Form("string")
	o.Global=True
	o.IgnoreCase=True
	o.Pattern="("&Request.Form("regexp")&")"
	'Set m=o.Execute(t)
	h=TransferHTML(t,"[html-format]")
	t=h
	'For Each s in m
		'h=Replace(h,s.value,"<span style=""background-color:#92d050"">"&s.value&"</span>")
	'Next
	h=o.replace(h,"<span style=""background-color:#92d050"">$1</span>")
	n1=RunTime
	If Err.Number<>0 Then
		h="正则有误："& Err.Description
		h=h&"<br/>可能的情况是：<ol><li>少打了某个符号</li><li>没有在[ ] ( ) ^ . ? !等符号前加\</li></ol>"
	Else
		n="用时"&n1&"ms"
		if (t<>h) then h=n&"<br/><br/>检测到黑词或敏感词：<br/><br/>"&h  else h=n&"<br/><br/>"&h
	End If
	
	
	
	
	
	Response.Write h
	Response.End
End If
%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><a href="setting.asp"><span class="m-left">TotoroⅢ设置</span></a><a href="regexptest.asp"><span class="m-right m-now">黑词测试</span></a><a href="onlinetest.asp"><span class="m-right">模拟测试</span></a></div>
          <div id="divMain2">
          <table width='100%' style='padding:0px;margin:1px;line-height:20px' cellspacing='0' cellpadding='0'>
          <tr height="40"><td width="50%">输入待测试内容</td>
            <td>结果</td>
          </tr>
          <tr><td><textarea rows="6" name="test" id="test" style="width:99%" ></textarea></td>
            <td rowspan="4" style="text-indent:0;vertical-align:top"><div id="result"></div></td>
          </tr>
          <tr height="40"><td>输入黑词列表或过滤词列表</td>
            </tr>
          <tr><td><textarea rows="6" name="regexp" id="regexp" style="width:99%" ></textarea></td>
            </tr>
          <tr><td><input type="button" class="button" value="提交测试" id="buttonsubmit"/></td></tr>
          </table>
          </div>
        </div>
        <script type="text/javascript">
		$(document).ready(function(e) {
            $("#buttonsubmit").bind("click",function(){
				var o=$.ajax({
					url:"regexptest.asp?type=test",
					async:false,
					type:"POST",
					data:{"string":$("#test").attr("value"),"regexp":$("#regexp").attr("value")},
					dataType:"script",
					/*success:function(data){
						alert(data);
					}*/
				});
				$("#result").html(o.responseText);
			})
			
        });
        </script>
<script type="text/javascript">ActiveLeftMenu("aCommentMng");</script>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
