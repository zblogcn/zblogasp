<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp"-->
<%
Dim bError
bError=False
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)

Call AppCentre_InitConfig
'BlogVersion="130128"
'ZC_BLOG_VERSION="Z-Blog 2.1 Phoenix Build 130128"
If Request.QueryString("restore")="now" Then
	Response.Clear
	Response.Write AppCentre_Update_Restore(Request.Form("build"),Request.Form("filename"),Request.Form("crc32"))
	Response.End
End If


If Request.QueryString("update")="download" Then
	Response.Clear
	Response.Write AppCentre_Update_Download(Request.Form("filename"))
	Response.End
End If

If Request.QueryString("update")="install" Then
	Response.Clear
	Response.Write AppCentre_Update_Install()
	Response.End
End If

If Request.QueryString("update")="success" Then
	Response.Clear
	Call SetBlogHint_Custom("恭喜您升级到最新版的Z-Blog.")
	Response.Redirect BlogHost & "zb_system/cmd.asp?act=SettingMng&update=" & Request.QueryString("file")
	Response.End
End If

If Request.QueryString("last")="now" Then
	Response.Clear
	Response.Write AppCentre_CheckSystemLast
	Response.End
End If


If Request.QueryString("check")="now" Then
	Call AppCentre_CheckSystemIndex(BlogVersion)
End If

Dim PathAndCrc32
Set PathAndCrc32=New TMeta

Dim objXmlFile,strXmlFile,item
Dim fso, f, f1, fc, s
Set fso = CreateObject("Scripting.FileSystemObject")


If fso.FileExists(BlogPath & "zb_users/cache/"&BlogVersion&".xml") Then

	strXmlFile =BlogPath & "zb_users/cache/"&BlogVersion&".xml"

	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.load(strXmlFile)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else

			for each item in objXmlFile.documentElement.SelectNodes("file")
				PathAndCrc32.SetValue item.getAttributeNode("name").Value,item.getAttributeNode("crc32").Value
			next

		End If
	End If
End If


If CLng(Request.QueryString("crc32"))>0 Then

	Response.Clear
	If CLng(Request.QueryString("crc32"))<=Round(PathAndCrc32.Count/10)+1 Then

		Dim i,j,k,l,m,n
		k=CLng(Request.QueryString("crc32"))
		i=(k-1)*10+1
		j=k*10
		m="<img src=\'"&BlogHost&"zb_system/image/admin/ok.png\' width=\'16\' alt=\'\' />"
		n="<a href=\'javascript:void(0)\' onclick=\'restore(this)\' title=\'还原系统文件\'><img src=\'"&BlogHost&"zb_system/image/admin/exclamation.png\' width=\'16\' alt=\'\' /></a>"
		For l=i To j
			If l>PathAndCrc32.Count Then Exit For
			If CRC32(BlogPath & vbsunescape(PathAndCrc32.Names(l)))<>PathAndCrc32.Values(l) Then
				Response.Write "$('#td"&l&"').html('"&n&"').parent().addClass(""check_conflict"");_conflict+=1;_count.html(_conflict);"
			Else
				Response.Write "$('#td"&l&"').html('"&m&"').parent().addClass(""check_normal"");"
			End If
		Next
	Else
		Call DelToFile(BlogPath & "zb_users/cache/"&BlogVersion&".xml")
	End If
	Response.End

End If


BlogTitle="应用中心-系统更新检查"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu">
            <% AppCentre_SubMenu(7)%>
          </div>
          <div id="divMain2">
            <form method="post" action="">
              <table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
                <tr>
                  <th width='50%'>当前版本</th>
                  <th>最新版本</th>
                </tr>
                <tr>
                  <td align='center' id='now'>Z-Blog <%=ZC_BLOG_VERSION%></td>
                  <td align='center' id='last'></td>
                </tr>
              </table>
              <p>
                <input id="updatenow" type="button" onClick="update();return false;" style="visibility:hidden;" value="升级新版程序" />
              </p>
			  <hr/>

              <div class="divHeader">校验系统核心文件&nbsp;&nbsp;<span id="checknow"><a href="update.asp?check=now" title="开始校验"><img src="Images/refresh.png" width="16" alt="校验" /></a></span></div>
			  <div>进度<span id="status">0</span>%；已发现<span id="count">0</span>个修改过的系统文件。<div id="bar"></div></div>
              <table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
                <tr>
                  <th width='78%'>文件名</th>
                  <th id="_s">状态</th>
                </tr>

<%
Dim FileManageEnabled
FileManageEnabled=CheckPluginState("FileManage")
Dim a,b,c,d,e
b=0
For Each a In PathAndCrc32.Names

If b>0 Then

c=vbsunescape(a)

Response.Write "<tr><td>"
Response.Write AppCentre_GetTypeIco(c)
Response.Write " <span>"& c &"</span><cite style='display:none;'>"& PathAndCrc32.GetValue(c) &"</cite></td><td id='td"&b&"' align='center'>"& e &"</td></tr>"
Response.Flush

End If
b=b+1
Next


%>
              </table>
              <p> </p>
            </form>
          </div>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script> 
        <script type="text/javascript">
			var _max = parseInt("<%=Round(PathAndCrc32.Count/10)+2%>"),
			_conflict = 0,
			_sort = 0;
			var _bar = $("#bar"),
			_status = $("#status"),
			_count = $("#count");
			
			function crc32(i) {
				if (i == 1) {
					$("#checknow").html("<img src='" + bloghost + "zb_system/image/admin/loading.gif' alt='loading' width='16' />");
					$("#checknow").parent().click(function() {
						return false
					});
				}
				_bar.prev().hide();
				$.get("update.asp?crc32=" + i, 
				function(data) {
					if (data !== "") {
						i = i + 1;
						_bar.progressbar({
							value: (i / _max) * 100
			
			
						});
						_status.html(parseInt((i / _max) * 100));
						eval(data);
						crc32(i);
			
			
			
					} else {
						_bar.hide();
						$("#checknow").hide();
						_bar.prev().show();
						_status.html(100);
						$("#_s").html("<a href='javascript:void(0);'>筛选</a>").find("a").click(function() {
							var o = $(this);
							switch (_sort) {
								case 0:
								$(".check_normal").hide();
								_sort = 1;
								break;
								case 1:
								$(".check_normal").show();
								$(".check_conflict").hide();
								_sort = 2;
								break;
								case 2:
								$(".check_conflict").show();
								_sort = 0;
								break;
			
			
							}
							return false
			
			
						});
			
			
					}
			
			
			
				});
			
			
			
			}
			
			function checklast(now, last) {
				var n = now.toString().match(/[0-9]{6}/);
				var l = last.toString().match(/[0-9]{6}/);
				if (l - n > 0) {
					$("form").attr("action", n + "-" + l);
					$("form input:button").css("visibility", "inherit");
			
			
				}
			
			
			
			}
			
			function update() {
			
				var s = Math.random().toString();
				var j = document.createElement("div");
				j.id = "dialog_" + s;
				j.innerHTML = "正在下载更新包<br/>";
				$(j).dialog({
					title: "提示",
					modal: true,
					buttons: {
						"取消": function() {
							//$(this).dialog("close");
							location.href = "update.asp"
			
			
						}
			
			
					}
			
			
				});
			
				update_download(j);
				//update_install(j);
			
			
			}
			
			
			function update_download(j) {
				$.post("update.asp?update=download", 
				{
					"filename": $("form").attr("action")
			
			
				},
				function(data) {
			
					if (data != "") {
						$(j).append(data + "<br/>");
						update_install(j)
			
			
					} else {
						$(j).append("升级失败<br/>");
			
			
					}
			
			
				});
			
			
			}
			
			function update_install(j) {
				$(j).append("开始安装文件包<br/>");
				delallJQueryUIDialogButton($(j));
				$.post("update.asp?update=install", 
				{
					"filename": $("form").attr("action")
			
			
				},
				function(data) {
			
					if (data != "") {
						$(j).append(data + "<br/>");
						createnewJQueryUIDialogButton($(j), "确定", 
						function() {
							update_success(j)
			
						});
			
			
					} else {
						$(j).append("升级失败<br/>");
						createnewJQueryUIDialogButton($(j), "确定", 
						function() {
							$(j).dialog("close")
			
						});
			
			
					}
			
			
				});
			
			
			
			}
			
			function update_success(j) {
				location.href = "update.asp?update=success&file=" + $("form").attr("action");
			
			
			}
			
			function restore(t) {
				$("#checknow").attr("src", bloghost + "zb_system/image/admin/loading.gif");
				$("#checknow").parent().click(function() {
					return false
				});
				$("#checknow").show();
				var b = $("#now").html().match(/[0-9]{6}/);
				var f = $(t).parent().prev().find("span").html();
				var c = $(t).parent().prev().find("cite").html();
				$.post("update.asp?restore=now", 
				{
					"build": b.toString(),
					"filename": f,
					"crc32": c
			
				},
				function(data) {
					$("#checknow").hide();
					if (data != "") {
						//alert(data);
						$(t).find("img").attr("src", bloghost + 'zb_system/image/admin/ok.png');
			
					} else {
						alert("更新失败");
			
					}
			
			
				});
			
			
			}
			
			
			function createnewJQueryUIDialogButton(jqobj, text, e) {
				var objid = "btn_" + parseInt(new Date().valueOf().toString() + Math.random() * 1000);
				jqobj.parent().find('.ui-dialog-buttonset').append('<button type="button" id="' + objid + '" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only" role="button" aria-disabled="false"><span class="ui-button-text">' + text + '</span></button>');
				$("#" + objid).click(e)
			
			}
			
			
			function delallJQueryUIDialogButton(jqobj) {
				jqobj.parent().find('.ui-dialog-buttonset').html("");
			
			}
			
			
			$(document).ready(function() {
			
				$.get("update.asp?last=now", 
				function(data) {
					$("#last").html("Z-Blog " + data);
					checklast($("#now").html(), $("#last").html());
			
			
				});
						
			<%
			If Request.QueryString("check")="now" Then
				Response.Write "crc32(1)"
			End If
			%>
			
						});
   
   </script>
        <%
	If login_pw<>"" Then
		Response.Write "<script type='text/javascript'>$('div.footer_nav p').html('&nbsp;&nbsp;&nbsp;<b>"&login_un&"</b>您好,欢迎来到APP应用中心!').css('visibility','inherit');</script>"
	End If
%>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
		
Function AppCentre_GetTypeIco(c)
	On Error Resume Next
	If FileManageEnabled And Not bError Then 
		Response.Write FileManage_GetTypeIco(c)
		If Err.Number<>0 Then bError=True:Response.Write "<img src='Images/document_empty.png' width='16' alt='' />"
	Else
		Response.Write "<img src='Images/document_empty.png' width='16' alt='' />"
	End If
End Function
%>