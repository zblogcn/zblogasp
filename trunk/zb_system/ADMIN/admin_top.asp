</head>
<body>
<div id="header">
  <div class="top">
    <div class="logo"><img src="<%=GetCurrentHost%>ZB_SYSTEM/image/admin/logo.png" alt="Z-Blog" title="Z-Blog"/></div>
    <div class="user"> <img src="<%=GetCurrentHost%>ZB_SYSTEM/image/admin/avatar.png" width="40" height="40" id="avatar" alt="Avatar" />
      <div class="username"><%=ZVA_User_Level_Name(BlogUser.Level)%>：<%=BlogUser.Name%></div>
      <div class="userbtn"><a class="profile" href="<%=GetCurrentHost%>" title="" target="_blank"><%=ZC_MSG065%></a> <a class="logout" href="<%=GetCurrentHost%>ZB_SYSTEM/cmd.asp?act=logout" title=""><%=ZC_MSG020%></a></div>
    </div>
    <div class="menu">
      <ul id="topmenu">
<%
Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(ZC_MSG340,"http://www.rainbowsoft.org/","_blank"))
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