</head>
<body>
<div id="header">
  <div class="top">
    <div class="logo"><img src="<%=GetCurrentHost%>zb_system/image/admin/logo.png" alt="Z-Blog" title="Z-Blog"/></div>
    <div class="user"> <img src="<%=GetCurrentHost%>zb_system/image/admin/avatar.png" width="40" height="40" id="avatar" alt="Avatar" />
      <div class="username"><%=ZVA_User_Level_Name(BlogUser.Level)%>：<%=BlogUser.Name%></div>
      <div class="userbtn"><a class="profile" href="<%=GetCurrentHost%>" title="" target="_blank"><%=ZC_MSG065%></a> <a class="logout" href="<%=GetCurrentHost%>ZB_SYSTEM/cmd.asp?act=logout" title=""><%=ZC_MSG020%></a></div>
    </div>
    <div class="menu">
      <ul id="topmenu">
<%
Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(ZC_MSG006,"http://www.rainbowsoft.org/","","_blank"))
%>
        <%=Response_Plugin_Admin_Top%>
      </ul>
    </div>
  </div>
</div>
<div id="main">
<!--#include file="admin_left.asp"-->
<div class="main_right">
  <div class="yui">
    <div class="content">
<script type="text/javascript">
<!--
function Batch2Tip(s){$("#batch p").html(s)}
function BatchContinue(){$("#batch p").before("<iframe style='width:20px;height:20px;' frameborder='0' scrolling='no' src='<%=GetCurrentHost%>zb_system/cmd.asp?act=batch'></iframe>");$("#batch img").remove();}
function BatchBegin(){};
function BatchEnd(){};
-->
</script>
<%
If IsObject(Session("batch"))=True Then
If Session("batch").Count>0 Then
	If Session("batch").Count= Session("batchorder") Then
		'
		'Session("batchtime")=0
%>
<div id="batch">
<iframe style="width:20px;height:20px;" frameborder="0" scrolling="no" src="<%=GetCurrentHost%>zb_system/cmd.asp?act=batch"></iframe><p><%=ZC_MSG110%>...</p>
</div>
<%
	Else
%>
<div id="batch"><img src="<%=GetCurrentHost%>zb_system/image/admin/warning.png" width="20"/><p><%=ZC_MSG273%></p></div>
<script type="text/javascript">
$("#batch a").bind("click", function(){ BatchContinue();$("#batch p").html("<%=ZC_MSG109%>...");});
</script>

<%
	End If
Else
	Session("batchorder")=0
End If
End If
%>