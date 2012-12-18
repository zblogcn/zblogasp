<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\..\c_option.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\..\..\plugin\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If (Not IsEmpty(Request.QueryString("s"))) Then Call metro_savetofile("custom.css")

BlogTitle="Metro主题配置"

Dim strBodyBg,aryBodyBg
Dim strHdBg,aryHdBg
Dim strColor,aryColor
Dim c
Set c = New TConfig
c.Load("metro")
If c.Exists("vesion")=true Then
	strBodyBg=c.read("custom_bodybg")
	strHdBg=c.read("custom_hdbg")
	strColor=c.read("custom_color")
	aryBodyBg=Split(strBodyBg,"|")
	aryHdBg=Split(strHdBg,"|")
	aryColor=Split(strColor,"|")
End If 

Dim i,a
a=Array("","左","中","右")
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<link href="evol.colorpicker.css" rel="stylesheet" /> 
<script src="evol.colorpicker.min.js" type="text/javascript"></script>
<script src="custom.js" type="text/javascript"></script>
<style>
table input{padding: 0;margin:0.25em 0;}
</style>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
	<div id="ShowBlogHint">
	<%Call GetBlogHint()%>
	</div>
	<div class="divHeader"><%=BlogTitle%></div>
	<div class="SubMenu"></div>
	<div id="divMain2"> 
		<form action="save.asp" method="post">
			<table width="100%" border="1" width="100%" class="tableBorder">
				<tr>
					<th scope="col" height="32" width="150px">主题背景</th>
					<th scope="col"></th>
				</tr>
				<tr>
					<td>页面背景</td>
					<td>
						<div id="bgcolor">
							背景颜色：<input id="bodybgc0" name="bodybg0"  value=<%=aryBodyBg(0)%> /> 
						</div>
						<div >
							<input type="checkbox" id="bodybgc5" name="bodybg5" <%=IIf(aryBodyBg(5)="True","checked=""checked""","")%> value="True"/> <label for="bodybgc5">使用背景图</label>
						</div>
						<div id="bodybgmain" <%=IIf(aryBodyBg(5)="","style=""display:none""","")%>>
							背景图：<input id="bgurl" name="bodybg1"  value=<%=aryBodyBg(1)%> /> 
							<div id="bodybgs">背景设定：
							<input type="checkbox" id="bodybg2r" name="bodybg2" <%=IIf(InStr(aryBodyBg(2),"repeat")>0,"checked=""checked""","")%> value="repeat"/><label for="bodybg2r">平铺</label>
							<input type="checkbox" id="bodybg2f" name="bodybg2" <%=IIf(InStr(aryBodyBg(2),"fixed")>0,"checked=""checked""","")%> value="fixed"/><label for="bodybg2f">固定</label>
							</div> 
							<div id="bgpx"> 对齐方式： 			  
							<%	For i=1 To 3	%>
							<input type="radio" id="bgpx<%=i%>" name="bodybg3" value="<%=i%>" <%=IIf(i=int(aryBodyBg(3)),"checked=""checked""","")%> /><label for="bgpx<%=i%>">居<%=a(i)%></label>
							<%	Next  %>
							</div> 
							<input type="hidden" id="bgpy" name="bodybg4"  value=<%=aryBodyBg(4)%> />
						</div>
					</td>
				</tr>
				<tr>
					<td>顶部背景</td>
					<td>
						<div >顶部高度：<input id="hdbgph" type="text" name="hdbg5"  value=<%=aryHdBg(5)%> />(单位：px)</div>
						<div id="hdbgcolor" <%=IIf(aryHdBg(6)="True","style=""display:none""","")%>>
							<input type="checkbox" id="hdbgc0" name="hdbg0" <%=IIf(aryHdBg(0)="transparent","checked=""checked""","")%> value="transparent"/><label for="hdbgc0"> 背景透明（不透明情况下使用主色为背景色）</label>
						</div>
						<div >
							<input type="checkbox" id="hdbgc6" name="hdbg6" <%=IIf(aryHdBg(6)="True","checked=""checked""","")%> value="True"/> <label for="hdbgc6">使用背景图</label>
						</div>
						<div id="hdbgmain" <%=IIf(aryHdBg(6)="","style=""display:none""","")%>>
							背景图：<input id="hdbgurl" name="hdbg1"  value=<%=aryHdBg(1)%> /> 
							<div id="hdbgs">背景设定：
							<input type="checkbox" id="hdbg2r" name="hdbg2" <%=IIf(InStr(aryHdBg(2),"repeat")>0,"checked=""checked""","")%> value="repeat"/><label for="hdbg2r">平铺</label>
							<input type="checkbox" id="hdbg2f" name="hdbg2" <%=IIf(InStr(aryHdBg(2),"fixed")>0,"checked=""checked""","")%> value="fixed"/><label for="hdbg2f">固定</label>
							</div> 
							<div id="hdbgpx"> 对齐方式： 			  
								<%	For i=1 To 3	%>
										<input type="radio" id="hdbgpx<%=i%>" name="hdbg3" value="<%=i%>" <%=IIf(i=int(aryHdBg(3)),"checked=""checked""","")%> /><label for="hdbgpx<%=i%>">居<%=a(i)%></label>
								<%	Next 	%>
							</div> 
						<input id="hdbgpy" type="hidden" name="hdbg4"  value=<%=aryHdBg(4)%> />
						</div>
					</td>
				</tr>
			</table>

			<table width="100%" border="1" width="100%" class="tableBorder">
				<tr>
					<th scope="col" height="32" width="150px">颜色配置</th>
					<th scope="col">（预设方案：<a  style="cursor: pointer;" onclick="loadConfig(theme_config.default);">默认</a>  <a  style="cursor: pointer;" onclick="loadConfig(theme_config.green);">绿色</a>）</th>
				</tr>
				<tr>
					<td>主色（深）</td>
					<td><input id="colorP1" name="color"  value=<%=aryColor(0)%> /></td>
				</tr>
				<tr>
					<td>次色（浅）</td>
					<td><input  id="colorP2" name="color"  value=<%=aryColor(1)%> /></td>
				</tr>
				<tr>
					<td>字体颜色</td>
					<td><input  id="colorP3" name="color"  value=<%=aryColor(2)%> /></td>
				</tr>
				<tr>
					<td>链接颜色</td>
					<td><input  id="colorP4" name="color"  value=<%=aryColor(3)%> /></td>
				</tr>
				<tr>
					<td>文章背景色</td>
					<td><input  id="colorP5" name="color"  value=<%=aryColor(4)%> /></td>
				</tr>
			</table>
			<input name="ok" type="submit" class="button" value="提交"/>
		</form>
	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
<script type="text/javascript">
ActiveTopMenu("ametroManage");
</script> 

<%Call System_Terminate()%>