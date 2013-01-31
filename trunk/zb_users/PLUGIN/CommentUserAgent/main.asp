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
'检查非法链接
Call CheckReference("")

Dim UserAgent,Req
Req=Request.QueryString("list")
If IsEmpty(Req) Then Req=0
Req=CInt(Req)
UserAgent=Request.ServerVariables("HTTP_USER_AGENT").Item
'检查权限
'If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("CommentUserAgent")=False Then Call ShowError(48)
BlogTitle="评论UA"
Dim objConfig
Set objConfig=New TConfig
objConfig.Load "CommentUserAgent"
	Dim aryUA(29)
	aryUA(0)="Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.19 (KHTML, like Gecko) Chrome/25.0.1323.1 Safari/537.19"
	aryUA(1)="Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.19 (KHTML, like Gecko) Chrome/25.0.1323.1 Safari/537.19"
	aryUA(2)="Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; WOW64; Trident/6.0)"
	aryUA(3)="Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)"
	aryUA(4)="Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
	aryUA(5)="Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/5.0)"
	aryUA(6)="Mozilla/5.0 (Windows NT 6.2; WOW64; rv:18.0) Gecko/20100101 Firefox/18.0"
	aryUA(7)="Mozilla/5.0 (Windows NT 6.1; rv:18.0) Gecko/20100101 Firefox/18.0"
	aryUA(8)="Opera/9.80 (Windows NT 6.2; WOW64; Edition IBIS) Presto/2.12.388 Version/12.11"
	aryUA(9)="Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.802.30 Safari/535.1 SE 2.X MetaSr 1.0"
	aryUA(10)="Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.92 Safari/537.1 LBBROWSER"
	aryUA(11)="Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.1 (KHTML, like Gecko) Maxthon/4.0.3.1000 Chrome/22.0.1229.79 Safari/537.1"
	aryUA(12)="Mozilla/4.0 (compatible; MSIE 5.00; Windows 98) "
	aryUA(13)="Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0; SV1; QQBrowser/7.0.2433.202)"
	aryUA(14)="Mozilla/5.0 (Windows; U; Windows NT 5.1; zh_CN) AppleWebKit/534.7 (KHTML, like Gecko) Chrome/18.0 BIDUBrowser/2.x Safari/534.7"
	aryUA(15)="Mozilla/5.0 (Linux; U; Android 2.3.6; zh-cn; SCH-I779 Build/GINGERBREAD) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1"
	aryUA(16)="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Ubuntu Chromium/23.0.1271.97 Chrome/23.0.1271.97 Safari/537.11"
	aryUA(17)="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_2) AppleWebKit/536.26.17 (KHTML, like Gecko) Version/6.0.2 Safari/536.26.17"
	aryUA(18)="Mozilla/5.0 (X11; Linux i686; rv:16.0) Gecko/16.0 Firefox/16.0"
	aryUA(19)="Mozilla/5.0 (iPhone; CPU iPhone OS 5_0 like Mac OS X) AppleWebKit/534.46 (KHTML, like Gecko) Version/5.1 Mobile/9A334 Safari/7534.48.3"
	aryUA(20)="Mozilla/5.0 (iPad; CPU OS 5_0 like Mac OS X) AppleWebKit/534.46 (KHTML, like Gecko) Version/5.1 Mobile/9A334 Safari/7534.48.3"
	aryUA(21)="Mozilla/5.0 (Windows NT 5.1; Avant TriCore) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.97 Safari/537.11"
	aryUA(22)="Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.11 TaoBrowser/2.0 Safari/536.11"
	aryUA(23)="Mozilla/5.0 (Linux; U; Android 2.3.5; zh-cn; M9 Build/GRJ90) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1"
	aryUA(24)="Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; SV1; SaaYaa)"
	aryUA(25)="Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322; TencentTraveler)"
	aryUA(26)="Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.1.16) Gecko/20121020 Iceweasel/3.5.16 (like Firefox/3.5.16)"
	aryUA(27)="Mozilla/5.0 (X11; Linux i686) AppleWebKit/534.34 (KHTML, like Gecko) rekonq/1.1 Safari/534.34"
	aryUA(28)="Opera/9.80 (Windows NT 6.1; WOW64; U; en) Presto/2.10.229 Version/11.64"
	aryUA(29)="Mozilla/5.0 (X11; U; FreeBSD amd64; en; rv:1.8.1.20) Gecko/20090422 Epiphany/2.22 Firefox/2.0.0.20"

If Request.QueryString("act")="design" Then
	Dim objRs,objRs2
	Set objRs=objConn.Execute("SELECT [comm_id] FROM [blog_Comment] WHERE [comm_Agent]=''")
	Do Until objRs.Eof
		Randomize
		objConn.Execute "UPDATE [blog_Comment] SET [comm_Agent]='"&aryUA(Int(30 * Rnd))&" (Unreal_CommentUserAgentAutoGenerate)' WHERE [comm_id]=" & objRs("comm_id")
		objRs.MoveNext
	Loop
	Call SetBlogHint_Custom("SUCCESS!")
ElseIf Request.QueryString("act")="back" Then
	objConn.Execute "UPDATE [blog_Comment] SET [comm_Agent]='' WHERE [comm_Agent] LIKE '%(Unreal_CommentUserAgentAutoGenerate)'"
ElseIf Request.QueryString("act")="save" Then
	Dim i,sss(),mmm,jjj
	Redim sss(-1)
	mmm=0
	jjj=Split(Request.Form("selecttags").Item,", ")
	For i=0 To 13 
		If jjj(i)="True" Then 
			Redim Preserve sss(mmm)
			sss(mmm)=i
			mmm=mmm+1
		End If
		objConfig.Write "Item",Join(sss,",")
		objConfig.Save
		Call SetBlogHint(True,Empty,Empty)
	Next
End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
td{white-space:nowrap}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=commentuseragent.functions.submenu(Req)%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
            <%
				If Req=0 Then
			%>
            <%
				
				Dim k
				k=objConfig.Read("Item")
				k=","&k&","
				
				Dim s
				Set s=detect_webbrowser(UserAgent)
				%>
            <%
				Dim o
				Set o=detect_platform(UserAgent)
				
				%>
            <p>以下，您可以选择要在评论后跟上哪些内容，也可以自己在模板里插入标签。</p>
            <form method="post" enctype="application/x-www-form-urlencoded" action="main.asp?act=save">
              <table width="100%">
                <thead>
                  <tr height="32">
                    <th width="05%">显示</th>
                    <th width="20%">标签</th>
                    <th width="10%">效果</th>
                    <th>详细</th>
                  </tr>
                </thead>
                <tbody>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(0)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/useragent#&gt;</td>
                    <td><%=TransferHTML(UserAgent,"[html-format]")%></td>
                    <td>User-Agent</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(1)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/platform/src#&gt;</td>
                    <td><%=BlogHost & "zb_users/plugin/commentuseragent/img/" & o.fullfilename%></td>
                    <td>设备的图片地址</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(2)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/platform/system#&gt;</td>
                    <td><%=o.text%></td>
                    <td>设备的详细说明</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(3)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/platform/version#&gt;</td>
                    <td><%=o.ver%></td>
                    <td>设备版本号</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(4)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/platform/link#&gt;</td>
                    <td><%=o.link%></td>
                    <td>设备官网</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(5)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/platform/img#&gt;</td>
                    <td><%="<img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & o.fullfilename&"' width='16 height='16' alt='"&o.text&"系统' />"%></td>
                    <td>设备Logo</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(6)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/platform#&gt;</td>
                    <td><%="<span class='cmtua_platform'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & o.fullfilename&"' width='16 height='16' alt='"&o.text&"系统' />"&o.text&"</span>"%></td>
                    <td>设备</td>
                  </tr>

                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(7)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/browser/src#&gt;</td>
                    <td><%=BlogHost & "zb_users/plugin/commentuseragent/img/" & s.fullfilename%></td>
                    <td>浏览器图片地址</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(8)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/browser/browser#&gt;</td>
                    <td><%=s.text%></td>
                    <td>浏览器详细说明</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(9)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/browser/version#&gt;</td>
                    <td><%=s.ver%></td>
                    <td>浏览器版本号</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(10)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/browser/link#&gt;</td>
                    <td><%=s.link%></td>
                    <td>浏览器官网</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(11)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/browser/img#&gt;</td>
                    <td><%="<img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & s.fullfilename&"' width='16 height='16' alt='"&s.text&"浏览器' />"%></td>
                    <td>浏览器Logo</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(12)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/browser#&gt;</td>
                    <td><%="<span class='cmtua_browser'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & s.fullfilename&"' width='16 height='16' alt='"&s.text&"浏览器' />"&s.text&"</span>"%></td>
                    <td>浏览器</td>
                  </tr>
                  <tr height="32">
                    <td><input type="text" class="checkbox" name="selecttags" value="<%=GConfig(13)%>" /></td>
                    <td>&lt;#zsxsoft/cmtua/all#&gt;</td>
                    <td><%="<div class='cmtua'><span class='cmtua_platform'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & o.fullfilename&"' width='16 height='16' alt='"&o.text&"系统' />"&o.text&"</span><span class='cmtua_browser'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & s.fullfilename&"' width='16 height='16' alt='"&s.text&"浏览器' />"&s.text&"</span></div>"%></td>
                    <td>直接显示全部信息</td>
                  </tr>
                </tbody>
              </table>
              <p><input type="submit" class="button" value="保存设置" /></p>
            </form>
            <%
			ElseIf req=1 Then
			%>
            <p>插件内大部分图标和代码来自WordPress插件wp-useragent,zsx将其转换为ASP后做成了Z-Blog插件。</p>
            <p>Z-Blog 2.0 Doomsday 121221和之前版本Z-Blog的Totoro有一个BUG，得到的评论均无User-Agent，所以一定会显示“Unknown”。</p>
            <p>您可以
              <input type="button" class="button" value="点击这里" onClick="location.href='?act=design'"/>
              为它们随便设计一个User-Agent。</p>
            <p>当然，如果您蛋疼，也可以
              <input type="button" class="button" value="点击这里" onClick="location.href='?act=back'"/>
              把随便设计的切换回来</p>
            <%
			ElseIf req=2 Then
			%>
            	<table width="100%"><thead><tr height="32"><td width="50%">图</td><td>UA</td></tr></thead><tbody>
            <%
				Dim kk
				kk=Split(LoadFromFile(Server.MapPath("test.txt"),"utf-8"),vbCrlf)
				'kk=aryUA
				For i=0 To Ubound(kk)
					Set o=detect_platform(kk(i))
			%>
            	<tr height="32"><td><%="<span class='cmtua_platform'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & o.fullfilename&"' width='16 height='16' alt='"&o.text&"系统' />"&o.text&"</span>"%>
            <%
					Set s=detect_webbrowser(kk(i))
			%>
            	<%="<span class='cmtua_browser'><img src='"&BlogHost & "zb_users/plugin/commentuseragent/img/" & s.fullfilename&"' width='16 height='16' alt='"&s.text&"浏览器' />"&s.text&"</span>"%></td>
            	<td><%=kk(i)%></td></tr>
            <%
				Next
			%>
            	</tbody></table>
            <%
			End If%>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>

<%

Function GConfig(code)
	If InStr(k,","&code&",") Then GConfig=True Else GConfig=False
End Function
%>