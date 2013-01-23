<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBDK")=False Then Call ShowError(48)
BlogTitle=zbdk_title

%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="../css/BlogConfig.css" type="text/css" media="screen"/>
<script type="text/javascript" src="../script/templatetags.js"></script>
<script type="text/javascript">

</script>
<style type="text/css">
td {
	text-align: center
}
</style>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <%=ZBDK.submenu.Export("TemplateTags")%> </div>
          <div id="divMain2">
            <div class="DIVBlogConfig">
              <div class="DIVBlogConfignav">
                <ul name="tree" id="tree">
                </ul>
              </div>
              <div id="result" class="DIVBlogConfigcontent">
                <div class="DIVBlogConfigtop"><span id="templatename">请选择</span></div>
                <table width="100%" style='padding:0px;' cellspacing='0' cellpadding='0' id="table_tr">
                  <tr height='32'>
                    <th width="25%">标签</th>
                    <th>注释</th>
                    <th>其他说明</th>
                  </tr>
                </table>
              </div>
              <div class="clear"></div>
            </div>
          </div>
        </div>
      </div>
    </div>
    <script type="text/javascript">
$(document).ready(function() {
	$("#tree").html(function(){
		var o=template_tags.filename,str="";
		for(var i=0;i<o.length;i++){
			str+="<li><a href='javascript:;' title='"+o[i].data+"'>"+(o[i].filename!="all"?o[i].filename+".html":o[i].data)+"</a></li>"
		};
		return str;
	});
	$("#tree li a").click(function(){
		var o=$(this);
		$("#templatename").html(o.text()+" "+o.attr("title"));
		
	})
});
</script> 
    <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->