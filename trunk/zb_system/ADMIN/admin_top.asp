</head>
<body>
<div id="header">
  <div class="top">
    <div class="logo"><img src="<%=GetCurrentHost%>ZB_SYSTEM/image/IMG/logo.png" alt="Z-Blog" title="Z-Blog"/></div>
    <div class="user"> <img src="<%=GetCurrentHost%>ZB_SYSTEM/image/IMG/avatar.png" width="40" height="40" class="avatar" alt="Avatar" />
      <div class="username"><%=ZVA_User_Level_Name(BlogUser.Level)%>：<%=BlogUser.Name%></div>
      <div class="userbtn"><a class="profile" href="<%=GetCurrentHost%>" title="" target="_blank"><%=ZC_MSG065%></a> <a class="logout" href="../cmd.asp?act=logout" title=""><%=ZC_MSG020%></a></div>
    </div>
    <div class="menu">
      <ul>
        <li><a href="<%=GetCurrentHost%>zb_system/cmd.asp?act=admin"><%=ZC_MSG245%></a></li>
        <li><a href="<%=GetCurrentHost%>zb_system/cmd.asp?act=SettingMng"><%=ZC_MSG247%></a></li>
        <li><a href="http://www.rainbowsoft.org/" target="_blank"><%=ZC_MSG340%></a></li>
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