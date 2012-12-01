<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
BlogTitle="分页条优化选项"

'读取配置
Call BetterPagebar_Config
 
If (Not IsEmpty(Request.QueryString("s"))) Then Call BlogReBuild_Default

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
	<style>
		small{
			font-size:12px;
		}
		h3{
			font-size:16px;
			background:#E4F2FD;
			height:20px;
			margin:10px 5px;
			padding:3px 5px;
		}
		h3.s{
			background:#FFF6EF;
			border:1px solid #FFA65F;
		}
		p{
			line-height:150%;
			margin:5px 10px;
		}
		hr{
			visibility:visible;
			margin:10px;
			border: 1px solid #E4F2FD;
		}
	</style>
	<title><%=BlogTitle%></title>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader"><%=BlogTitle%></div>
<div id="divMain2">

<form id="edit" name="edit" method="post" action="savesetting.asp">

	<h3 id="normal">一般选项</h3>

	<p>
	<input type="checkbox" value="True" id="AlwaysShow" name="AlwaysShow" <%
	If BetterPagebar_AlwaysShow then
		Response.Write " checked=""checked"">"
	else
		Response.Write ">"
	End if %>
	 <strong>是否总是显示分页向导（首页、上页、下页、尾页）？</strong>
	</p>
	
	
	<h3 id="text">文本选项</h3>
	
	<p><input type="text" value="<%=BetterPagebar_FristPage%>" size="30" id="FristPage" name="FristPage"/> <strong> 首页文字 </strong><br/></p>
	<p><input type="text" value="<%=BetterPagebar_FristPage_Tip%>" size="30" id="FristPage_Tip" name="FristPage_Tip"/> <strong> 首页链接提示 </strong><br/></p>
<hr/>			
	<p><input type="text" value="<%=BetterPagebar_LastPage%>" size="30" id="LastPage" name="LastPage"/> <strong> 尾页文字 </strong><br/></p>
	<p><input type="text" value="<%=BetterPagebar_LastPage_Tip%>" size="30" id="LastPage_Tip" name="LastPage_Tip"/> <strong> 尾页链接提示 </strong><br/></p>
<hr/>		
	<p><input type="text" value="<%=BetterPagebar_PrvePage%>" size="30" id="PrvePage" name="PrvePage"/> <strong> 上一页文字 </strong><br/></p>
	<p><input type="text" value="<%=BetterPagebar_PrvePage_Tip%>" size="30" id="PrvePage_Tip" name="PrvePage_Tip"/> <strong> 上一页链接提示 </strong><br/></p>
<hr/>			
	<p><input type="text" value="<%=BetterPagebar_NextPage%>" size="30" id="NextPage" name="NextPage"/> <strong> 下一页文字 </strong><br/></p>
	<p><input type="text" value="<%=BetterPagebar_NextPage_Tip%>" size="30" id="NextPage_Tip" name="NextPage_Tip"/> <strong> 下一页链接提示 </strong><br/></p>
<hr/>	
	<p><input type="text" value="<%=BetterPagebar_Extend%>" size="30" id="Extend" name="Extend"/> <strong> 过渡文字</strong> 
	<small>  * 其输出格式为 &lt;span&nbsp;class="extend"&gt;...&lt;/span&gt;。</small>	
	</p>

<hr/>	
<p><input type="submit" class="button" id="btnPost" value="提交"/>
	<br/><br/>
	<small>  * </small>
</p>	

<hr/>
</form>

</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
