<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
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

BlogTitle="WhitePage主题配置"

If Request.QueryString("save").Count>0 And Request.Form.Count>0 Then

	'Tconfig
	Dim c
	Set c = New TConfig
	c.Load("WhitePage")

	If Request.Form("bgcolor").Count=1 Then  c.Write "custom_bgcolor",Request.Form("bgcolor")
	If Request.Form("headtitle").Count=1 Then  c.Write "custom_headtitle",Request.Form("headtitle")
	If Request.Form("pagetype").Count=1 Then  c.Write "custom_pagetype",Request.Form("pagetype")
	If Request.Form("pagewidth").Count=1 Then  c.Write "custom_pagewidth",Request.Form("pagewidth")
	c.Save
	Set c =Nothing

	Call SetBlogHint(True,Empty,Empty)

End If


%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
	<div id="ShowBlogHint">
	<%Call GetBlogHint()%>
	</div>
	<div class="divHeader2"><%=BlogTitle%></div>
	<div class="SubMenu"></div>
	<div id="divMain2"> 
		<form action="?save" method="post">
			<table width="100%" border="1" width="100%" class="tableBorder">
				<tr>
					<th scope="col"  height="32" width="20%">整体设置</th>	
					<th></th>
				</tr>
				<tr>
					<td scope="col"  height="52">页面类型</td>					
					<td >
						<label><input type="radio" id="pt1" name="pagetype" value="1" />默认：图片阴影型（直角）</label>
						&nbsp;&nbsp;
						<label><input type="radio" id="pt2" name="pagetype" value="2" />CSS3阴影（直角）</label>
						&nbsp;&nbsp;
						<label><input type="radio" id="pt3" name="pagetype" value="3" />CSS3阴影（圆角）</label>
						&nbsp;&nbsp;
						<label><input type="radio" id="pt4" name="pagetype" value="4" />平面无阴影（直角）</label>
					</td>
				</tr>
				<tr>
					<td scope="col"  height="52">页面宽度</td>					
					<td >
						<label><input type="radio" id="pw1" name="pagewidth" value="1200" />1200px</label>
						&nbsp;&nbsp;
						<label><input type="radio" id="pw2" name="pagewidth" value="1000" />1000px</label>
					</td>
				</tr>
				<tr>
					<td scope="col"  height="52">标题对齐</td>					
					<td >
						<label><input type="radio" id="ht1" name="headtitle" value="left" />标题居左</label>
						&nbsp;&nbsp;
						<label><input type="radio" id="ht2" name="headtitle" value="center" />标题居中</label>
					</td>
				</tr>
			</table>

			<table width="100%" border="1" width="100%" class="tableBorder">
				<tr>
					<th scope="col" height="32" width="20%">颜色配置</th>
					<th scope="col">				
					<div  style="float:left;margin: 0.25em"></div>
					<div id="loadconfig"></div>
					</th>
				</tr>
				<tr>
					<td>背景色</td>
					<td>
						<label><input type="radio" id="bg0"  name="bgcolor" value="#FFFFFF" /><font color="#FFFFFF">#FFFFFF</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg1"  name="bgcolor" value="#FFA07A" /><font color="#FFA07A">#FFA07A</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg2"  name="bgcolor" value="#8FBC8B" /><font color="#8FBC8B">#8FBC8B</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg3"  name="bgcolor" value="#A9A9A9" /><font color="#A9A9A9">#A9A9A9</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg4"  name="bgcolor" value="#6699FF" /><font color="#6699FF">#6699FF</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg5"  name="bgcolor" value="#EE82EE" /><font color="#EE82EE">#EE82EE</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg6"  name="bgcolor" value="#9370DB" /><font color="#9370DB">#9370DB</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg7"  name="bgcolor" value="#FF7F50" /><font color="#FF7F50">#FF7F50</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg8"  name="bgcolor" value="#DEB887" /><font color="#DEB887">#DEB887</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg9"  name="bgcolor" value="#FFE4C4" /><font color="#FFE4C4">#FFE4C4</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg10" name="bgcolor" value="#7FFFD4" /><font color="#7FFFD4">#7FFFD4</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg11" name="bgcolor" value="#FFC0CB" /><font color="#FFC0CB">#FFC0CB</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg12" name="bgcolor" value="#BDB76B" /><font color="#BDB76B">#BDB76B</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg13" name="bgcolor" value="#D3D3D3" /><font color="#D3D3D3">#D3D3D3</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg14" name="bgcolor" value="#EEE8AA" /><font color="#EEE8AA">#EEE8AA</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg15" name="bgcolor" value="#98FB98" /><font color="#98FB98">#98FB98</font></label>&nbsp;&nbsp;
						<label><input type="radio" id="bg16" name="bgcolor" value="#FFB6C1" /><font color="#FFB6C1">#FFB6C1</font></label>&nbsp;&nbsp;
					</td>
				</tr>
			</table>
			<input name="ok" type="submit" class="button" value="保存配置"/>
		</form>
	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
<script type="text/javascript">
ActiveTopMenu("aWhitePageManage");
</script> 

<%Call System_Terminate()%>