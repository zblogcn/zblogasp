<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<!-- #include file="function.asp" -->

<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ReplaceWord")=False Then Call ShowError(48)
BlogTitle="敏感词替换器"
replaceword.init()
If Request.QueryString("act")="export" Then
	Response.ContentType="application/octet-stream"
	For i=0 To replaceword.words.length-1
		Response.Write IIf(replaceword.regex(i)=False,0,1) & "|"
		Response.Write TransferHTML(replaceword.str(i),"[textarea]")& "|"
		Response.Write TransferHTML(replaceword.rep(i),"[textarea]")& "|"
		Response.Write TransferHTML(replaceword.des(i),"[textarea]")& "|"
		Response.Write vbCrlf
		
	Next
	Response.End
End If
Dim i
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">

</style>
<script type="text/javascript" src="jquery.form.js" language="javascript"></script>

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=replaceword.submenu(IIf(Request.QueryString("act")="export",2,1))%></div>
          <div id="divMain2">
            <form id="form1" method="post" action="save.asp?act=import">
            	<%If Request.QueryString("act")<>"export" Then%>
                <p>导入有2种方式：一种是直接覆盖config.asp，一种是在下面批量添加。</p>
              <p>导入格式： 正则（开启为1，关闭为0）|敏感词|替换词|注释</p>
              <p>如： </p>
              <p>1|fuck|****|脏话</p>
              <p>0|sb||</p>
              <p>
              <%End If%>
                <label for="txaContent"></label>
                <textarea name="txaContent" id="txaContent" rows="10" style="width:50%"><%

				%></textarea>
              </p>
              <p>
              <%If Request.QueryString("act")<>"export" Then%>
                <p><input type="radio" name="type" id="type1" value="1" checked="checked"/>
                <label for="type1">不清空当前已有内容</label></p>
                <p><input type="radio" name="type" id="type2" value="2" />
                <label for="type2">清空后导入，此操作不可恢复，建议先<a href="?act=export">导出</a>，做好备份。</label></p>
                
              </p>
              <p><input type="submit" class="button" value="提交"/></p>
              <%End If%>
            </form>
          </div>
        </div>
        <div id="dialog" style="display:none"> </div>
        <script type="text/javascript">
			ActiveTopMenu("aPlugInMng");
			$(document).ready(function(){
				bmx2table();
				$("form").submit(function(){
					$(this).ajaxForm(function(s){
						var j=eval("("+s+")");
						if(j.success){
							showDialog("保存成功！","提示",function(){location.href="main.asp";});
						}else{showDialog("保存失败，错误ID："+j.error)}
					});
					return false;
				});
			});
			function showDialog(text,title,enter){
				if(enter==undefined) enter=function() {$(this).dialog("close");}
				var j=$('#dialog');
				j.html(text);
				j.dialog({
					modal: true,
					title: (title==undefined?"提示":title),
					buttons: {
						"确定": enter
					}
				})
			}
			
        </script> 
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
