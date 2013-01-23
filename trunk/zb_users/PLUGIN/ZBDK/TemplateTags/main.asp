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

                  
                </table>
              </div>
              <div class="clear"></div>
            </div>
          </div>
        </div>
      </div>
    </div>
    <script type="text/javascript">
	var _head="<tr height='32'><th width='25%'>标签</th><th>注释</th><th>其他说明</td></tr>";
$(document).ready(function() {
	$("#tree").html(function(){
		var o=template_tags.filename,str="";
		for(var i=0;i<o.length;i++){
			str+="<li><a href='javascript:;' title='"+o[i].data+"' _filename='"+o[i].filename+"'>"+(!o[i].isglobal?o[i].filename+".html":o[i].data)+"</a></li>"
		};
		return str;
	});
	$("#tree li a").click(function(){
		
		var o=$(this),s=template_tags.tags;
		var pp=o.attr("_filename"),jj=o.attr("title");
		$("#templatename").html(pp+" "+jj);
		var _tr="<tr height='32'><td><input type='text' value='<!tag!>' style='width:100%'/></td><td><!note!></td><td><!msg!></td></tr>";
		var j,k,l,m,str;
		str+=_head;
		for(j=0;j<s.length;j++){
			if(s[j].file.join(",").indexOf(pp.split(".html")[0])>-1){
				if(!s[j].ajax){
					k=s[j].tags;
					for(l=0;l<k.length;l++){
						m=_tr;
						m=m.replace("<!tag!>","<#"+k[l].tag+"#>").replace("<!msg!>",k[l].msg).replace("<!note!>",k[l].note);
						str+=m;
					}
				}
				else{
					s[j].ajax_config.success=function(data){
						str=_head;
						$("#table_tr").html(str+data);
						bmx2table()
					}
					str+="<tr height='32'><td>请稍候，获取中</td><td></td><td></td></tr>"
					$.ajax(s[j].ajax_config);
					
				}
			}
		}
		$("#table_tr").html(str);
		bmx2table()
		
	})
});
</script> 
    <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->