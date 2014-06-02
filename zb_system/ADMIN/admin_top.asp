</head>
<body>
<div id="header">
  <div class="top">
    <div class="logo"><a href="<%=BlogHost%>" title="<%=TransferHTML(ZC_BLOG_TITLE,"[html-format]")%>" target="_blank"><img src="<%=BlogHost%>zb_system/image/admin/logo.png" alt="Z-Blog"/></a></div>
    <div class="user"> <a href="<%=BlogHost%>zb_system/cmd.asp?act=UserEdt&amp;id=<%=BlogUser.ID%>" title="<%=ZC_MSG078%>"><img src="<%=BlogHost%>zb_system/image/admin/avatar.png" width="40" height="40" id="avatar" alt="Avatar" /></a>
      <div class="username"><%=ZVA_User_Level_Name(BlogUser.Level)%>：<%=BlogUser.FirstName%></div>
      <div class="userbtn"><a class="profile" href="<%=BlogHost%>" title="" target="_blank"><%=ZC_MSG065%></a>&nbsp;&nbsp;<a class="logout" href="<%=BlogHost%>zb_system/cmd.asp?act=logout" title=""><%=ZC_MSG020%></a></div>
    </div>
    <div class="menu">
      <ul id="topmenu">
<%
Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(GetRights("vrs"),ZC_MSG006,"http://www.zblogcn.com/","","_blank"))
%>
        <%=ResponseAdminTopMenu(Response_Plugin_Admin_Top)%>
      </ul>
    </div>
  </div>
</div>
<div id="main">
<!--#include file="admin_left.asp"-->
<div class="main_right">
  <div class="yui">
    <div class="content">
<%
If IsObject(Session("batch"))=True Then
If Session("batch").Count>0 Then
	If Session("batch").Count= Session("batchorder") Then
		'Session("batchtime")=0
%>
<div id="batch">
<iframe style="width:20px;height:20px;" frameborder="0" scrolling="no" src="<%=BlogHost%>zb_system/cmd.asp?act=batch"></iframe><p><%=ZC_MSG110%>...</p>
</div>
<%
	Else
%>
<div id="batch"><img src="<%=BlogHost%>zb_system/image/admin/error.png" width="16"/><p><%=ZC_MSG273%></p></div>
<%
	End If
Else
	Session("batchorder")=0
End If
End If
%>