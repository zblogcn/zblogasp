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
<%
Call System_Initialize()
Call Watermark_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("Watermark")=False Then Call ShowError(48)
BlogTitle="图片水印设置"

Dim act:act = Request.QueryString("act")
If act="save" Then Call SaveConfig()

Function SaveConfig()

	Dim strWATERMARK_TYPE,strWATERMARK_FONTBOLD,strWATERMARK_FONTQUALITY,strWATERMARK_FONTSIZE,strWATERMARK_FONTCOLOR,strWATERMARK_TEXT,strWATERMARK_QUALITY,strWATERMARK_WIDTH_POSITION,strWATERMARK_HEIGHT_POSITION,strWATERMARK_LOGO,strWATERMARK_ALPHA
	
	strWATERMARK_TYPE = Request.Form("strWATERMARK_TYPE")
	strWATERMARK_FONTBOLD = Request.Form("strWATERMARK_FONTBOLD")
	strWATERMARK_FONTQUALITY = Request.Form("strWATERMARK_FONTQUALITY")
	strWATERMARK_FONTSIZE = Request.Form("strWATERMARK_FONTSIZE")
	strWATERMARK_FONTCOLOR = Request.Form("strWATERMARK_FONTCOLOR")
	strWATERMARK_TEXT = Request.Form("strWATERMARK_TEXT")
	strWATERMARK_QUALITY = Request.Form("strWATERMARK_QUALITY")
	strWATERMARK_WIDTH_POSITION = Request.Form("strWATERMARK_WIDTH_POSITION")
	strWATERMARK_HEIGHT_POSITION = Request.Form("strWATERMARK_HEIGHT_POSITION")
	strWATERMARK_LOGO = Request.Form("strWATERMARK_LOGO")
	strWATERMARK_ALPHA = Request.Form("strWATERMARK_ALPHA")
	
	Watermark_Config.Write "TYPE",strWATERMARK_TYPE
	Watermark_Config.Write "FONTBOLD",strWATERMARK_FONTBOLD
	Watermark_Config.Write "FONTQUALITY",strWATERMARK_FONTQUALITY
	Watermark_Config.Write "FONTSIZE",strWATERMARK_FONTSIZE
	Watermark_Config.Write "FONTCOLOR",strWATERMARK_FONTCOLOR
	Watermark_Config.Write "TEXT",strWATERMARK_TEXT
	Watermark_Config.Write "WIDTH_POSITION",strWATERMARK_WIDTH_POSITION
	Watermark_Config.Write "HEIGHT_POSITION",strWATERMARK_HEIGHT_POSITION
	Watermark_Config.Write "QUALITY",strWATERMARK_QUALITY
	Watermark_Config.Write "LOGO",strWATERMARK_LOGO
	Watermark_Config.Write "ALPHA",strWATERMARK_ALPHA
	Watermark_Config.Save

	Call SetBlogHint(True,False,False)
	Response.Redirect "main.asp"

End Function
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
			<form name="edit" method="post" action="main.asp?act=save">
			<table width="100%" style="padding:0px;margin:1px;" cellspacing="0" cellpadding="0">
				<tr>
					<td style="width:32%"><p align="left">水印类型</p></td>
					<td><p>
					  <input type="radio" name="strWATERMARK_TYPE" value="1" <%=IIF(WATERMARK_TYPE=1,"checked","")%> />文字水印
					  <input type="radio" name="strWATERMARK_TYPE" value="2" <%=IIF(WATERMARK_TYPE=2,"checked","")%> />图片水印</p></td>
				</tr>
				<tr>
					<td><p align="left">水印水平位置</p></td>
					<td><p>
					<input type="radio" name="strWATERMARK_WIDTH_POSITION" value="left" <%=IIF(WATERMARK_WIDTH_POSITION="left","checked","")%> />左
					<input type="radio" name="strWATERMARK_WIDTH_POSITION" value="center" <%=IIF(WATERMARK_WIDTH_POSITION="center","checked","")%> />中
					<input type="radio" name="strWATERMARK_WIDTH_POSITION" value="right" <%=IIF(WATERMARK_WIDTH_POSITION="right","checked","")%> />右</p></td>
				</tr>
				<tr>
					<td><p align="left">水印垂直位置</p></td>
					<td><p>
					<input type="radio" name="strWATERMARK_HEIGHT_POSITION" value="top" <%=IIF(WATERMARK_HEIGHT_POSITION="top","checked","")%> />上
					<input type="radio" name="strWATERMARK_HEIGHT_POSITION" value="center" <%=IIF(WATERMARK_HEIGHT_POSITION="center","checked","")%> />中
					<input type="radio" name="strWATERMARK_HEIGHT_POSITION" value="bottom" <%=IIF(WATERMARK_HEIGHT_POSITION="bottom","checked","")%> />下</p></td>
				</tr>
				<tr>
					<td><p align="left">图片压缩质量(0-100,0为最低,100为最高)</p></td>
					<td><p>
					<input name="strWATERMARK_QUALITY" id="quality_range" style="width:230px;" type="range" min="0" max="100" value="<%=WATERMARK_QUALITY%>" onchange="document.getElementById('quality_num').value=this.value" /> <input id="quality_num" style="width:50px;text-align:center;vertical-align:top;" type="number" min="0" max="100" value="<%=WATERMARK_QUALITY%>" onchange="document.getElementById('quality_range').value=this.value" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">水印图片路径(插件内的相对路径,<span title="需aspjpeg1.8以上">支持png和gif透明</span>)</p></td>
					<td><p>
					<input name="strWATERMARK_LOGO" style="width:95%" type="text" value="<%=WATERMARK_LOGO%>" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">水印图片透明度(<span title="aspjpeg1.8以上仅对jpg有效">0-1,0为完全透明,1为不透明</span>)</p></td>
					<td><p>
					<input name="strWATERMARK_ALPHA" id="opaque_range" style="width:230px;" type="range" min="0" max="1" step="0.1" value="<%=WATERMARK_ALPHA%>" onchange="document.getElementById('opaque_num').value=this.value" /> <input id="opaque_num" style="width:50px;text-align:center;vertical-align:top;" type="number" min="0" max="1" step="0.1" value="<%=WATERMARK_ALPHA%>" onchange="document.getElementById('opaque_range').value=this.value" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">水印文字</p></td>
					<td><p>
					<input name="strWATERMARK_TEXT" style="width:95%" type="text" value="<%=WATERMARK_TEXT%>" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">文字质量(0默认,1草稿,2校样,3不抗锯齿,4抗锯齿)</p></td>
					<td><p>
					<input name="strWATERMARK_FONTQUALITY" style="width:95%" type="number" min="0" max="4" value="<%=WATERMARK_FONTQUALITY%>" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">文字大小</p></td>
					<td><p>
					<input name="strWATERMARK_FONTSIZE" style="width:95%" type="number" min="1" max="100" value="<%=WATERMARK_FONTSIZE%>" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">文字颜色</p></td>
					<td><p>
					<input name="strWATERMARK_FONTCOLOR" type="color" value="<%=WATERMARK_FONTCOLOR%>" />
					</p></td>
				</tr>
				<tr>
					<td><p align="left">文字是否粗体</p></td>
					<td><p>
					<input type="text" name="strWATERMARK_FONTBOLD" class="checkbox" value="<%=WATERMARK_FONTBOLD%>" />
				</tr>
			</table>
			<p>
				<input type="submit" class="button" value="提交" id="btnPost" />
				<input type="reset" class="button" value="重置" id="btnPost" />
			</p>
			</form>
          </div>
        </div>
<script type="text/javascript">
// <![CDATA[
$(document).ready(function(){
	var r = document.getElementById("quality_range");
	if(r.type == "text"){
		$("#opaque_range").hide();
		$("#quality_range").hide();
	}
});
// ]]>
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%Call System_Terminate()%>