</head>
<body>
<div id="header">
  <div class="top">
    <div class="logo"><img src="<%=ZC_BLOG_HOST%>ZB_SYSTEM/image/IMG/logo.png" alt="Z-Blog" title="Z-Blog"/></div>
    <div class="user"> <img src="<%=ZC_BLOG_HOST%>ZB_SYSTEM/image/IMG/avatar.png" width="40" height="40" class="avatar" alt="Avatar" />
      <div class="username"><%=ZVA_User_Level_Name(BlogUser.Level)%>：<%=BlogUser.Name%></div>
      <div class="userbtn"><a class="logout" href="../cmd.asp?act=logout" title=""><%=ZC_MSG020%></a> <a class="profile" href="../cmd.asp?act=UserMng" title=""><%=ZC_MSG070%></a></div>
    </div>
    <div class="menu">
      <ul>
        <li><a href="<%=ZC_BLOG_HOST%>"><%=ZC_MSG065%></a></li>
        <li><a href="<%=ZC_BLOG_HOST%>zb_system/cmd.asp?act=AskFileReBuild"><%=ZC_MSG247%></a></li>
        <li><a href="#"><%=ZC_MSG341%></a></li>
        <li><a href="http://www.rainbowsoft.org/" target="_blank"><%=ZC_MSG340%></a></li>
        <%=Response_Plugin_AdminTop_Plugin%>
      </ul>
    </div>
  </div>
</div>


<div id="main">
<!--#include file="admin_left.asp"-->
<div class="main_right">
  <div class="yui">
    <div class="content">