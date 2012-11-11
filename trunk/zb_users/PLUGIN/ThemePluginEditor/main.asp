<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ThemePluginEditor")=False Then Call ShowError(48)
BlogTitle="主题插件生成器 v1.1"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function newtr(This){
	var m=window.prompt("请输入一个文件名","xxxx.html");
	if(m==null) return false;
	$(This).parent().parent().before("<tr><td>"+m+"</td><td>"+"<select name=\"type_"+m+"\"><option value=\"1\" selected=\"selected\">文本</option><option value=\"2\">二进制</option></select></td><td><input type=\"text\" id=\""+m+"\" name=\"include_"+m+"\" value=\"\" style=\"width:98%\"/><input type=\"hidden\" id=\""+m+"_2\" name=\"new_"+m+"\"/></td><td align=\"center\"><a href='javascript:;' onclick='$(this).parent().parent().remove()'>删除</a></td></tr>");bmx2table();
}
var HAHAHA=false;
function shelp(){
	if(HAHAHA){$("#help").hide();HAHAHA=false}else{$("#help").show();HAHAHA=true}
}
function rename(obj,isnew){
	
	var _this=$(obj);
	var p=_this.parent().parent().children("td")
	var fs=$(p[0]).html();
	var j=prompt("请输入新文件名",fs);
	if(j!==null){
		$.get("save.asp",
		{
			"act":"rename"
			,"name":fs
			,"newname":j
		}
		,function(d){
			$(p[0]).html(j);
			$(p[1]).attr("name","type_"+j);
			$(p[2]).children().attr("id",j);
			$(p[2]).children().attr("name","include_"+j);
		})
	}
	
	
}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%>   <a href='javascript:;' onclick='shelp()'>帮助</a></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <div id="help" style="display:none">
            <p>这个插件，可以降低主题开发者的开发难度，让开发者把时间放于制作更加精美的主题而不是为了制作一个后台而苦恼。</p>
            <p>你需要给主题INCLUDE文件夹下添加需要引用的文件，这里就会自动出现文件名。同理，删除INCLUDE下的文件，这里也会相应删除。但是生成的主题插件始终不受到影响。</p>
            <p>若您的主题已有并非本插件生成的主题插件，请不要使用本插件！本插件必须修改主题xml。若您未修改xml，您可以通过<a href="howtouse.asp">这个页面</a>来手动生成xml。</p>
            <p><span style="color:red">注意：若当前主题已有自带插件，则请备份主题目录下PLUGIN文件夹</span></p>
            <p>常见使用问题：</p>
            <ol>
            <li>主题出现两个配置按钮：让用户点击“网站设置”-->“提交”即可。</li>
            <li>如何停用本主题插件：切换到其他主题-->编辑theme.xml，去掉<a href="howtouse.asp">这个页面</a>所述内容-->删除PLUGIN目录-->再切换回原主题。</li>
            </ol>
            <!--<p>更为详细的帮助请看：<a href="http://www.zsxsoft.com/archives/261.html" target="_blank">http://www.zsxsoft.com/archives/261.html</a></p>-->
            <div id="help001" style="display:none">
            <p>“文本”指文本文件，即txt、htm、js、css等允许用户直接修改的文件。“二进制”指图片、视频等无法直接修改的，让用户上传的文件。</p>
            <p>&nbsp;</p>
            <p>建议Logo等使用“二进制”，“广告位”“标语”等用“文本”。</p>
            </div>
            <div id="help002" style="display:none">
            <p>指展现给用户看的文字</p>
            </div></div>
            <form action="save.asp" method="post">
              <table width="100%" border="1" width="100%" class="tableBorder">
              <tr>
              <th scope="col" height="32" width="150px">文件名</th>
              <th scope="col" width="100px">文件类型 <a id="help01" href="$help001?width=320" class="betterTip" title="帮助">？</a></th>
              <th scope="col">文件注释 <a id="help02" href="$help002?width=320" class="betterTip" title="帮助">？</a></th>
              <th scope="col" width="100px"></th>
            </tr>
            <%
			Dim objConfig
			Set objConfig=New TConfig
			objConfig.Load "ThemePluginEditor_"&ZC_BLOG_THEME
			Dim oFso,oF
			Set oFso=Server.CreateObject("scripting.filesystemobject")
			If oFSO.FolderExists(BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\include")=False Then
				oFSO.CreateFolder BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\INCLUDE"
			End If
			Set oF=oFso.GetFolder(BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\include").Files
			Dim oS,s
			For Each oS In oF
			s=TransferHTML(oS.Name,"[html-format]")
			%>
            <tr>
            <td><%=oS.Name%></td>
            <td><select name="type_<%=s%>">
            <option value="1">文本</option>
            <option value="2"<%=IIf(isHTML(s),""," selected=""selected"")")%>>二进制</option></select></td><td>
            <input type="text" id="<%=s%>" name="include_<%=s%>" value="<%=objConfig.Read(s)%>" style="width:98%"/>
            </td>
            <td align="center"><a href='javascript:;' onclick='rename(this)'>改名</a>&nbsp;&nbsp;<a href='javascript:;' onclick='var _this=this;if(confirm("确定要删除吗？删除了不可恢复！")){$.get("save.asp",{"act":"del","name":$($(this).parent().parent().children("td")[0]).html()},function(d){$(_this).parent().parent().remove();})}'>删除</a></td>
            </tr>
            
            <%
			Next
			%>
            <tr id="new"><td><a href='javascript:;' onclick='newtr(this)'>新建..</a></td><td></td><td></td><td></td></tr>
            </table>
            <input type="submit" value="提交" class="button"/>
            <input type="hidden" name="delete" value="" id="delete"/>
          </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%
Function isHTML(s)
	Dim j
	If Instr(s,".") Then
		j=Mid(s,InstrRev(s,".")+1)
	Else
		j=""
	End If
	Select Case LCase(j)
		Case "html","htm","css","js","txt","xml","asp" isHTML=True
	End Select
End Function
%>