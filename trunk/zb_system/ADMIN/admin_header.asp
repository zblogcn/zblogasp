<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="robots" content="nofollow" />
<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> 
<title><%=BlogTitle%></title>
<link href="<%=GetCurrentHost%>ZB_SYSTEM/CSS/admin2.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/script/common.js" type="text/javascript"></script>
<link rel="stylesheet" href="<%=GetCurrentHost%>ZB_SYSTEM/CSS/jquery.bettertip.css" type="text/css" media="screen" />
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/script/jquery.bettertip.pack.js" type="text/javascript"></script>
<script language="JavaScript" src="<%=GetCurrentHost%>ZB_SYSTEM/script/jquery-ui-1.8.21.custom.min.js" type="text/javascript"></script>
<%If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE 6.0;")>0 Then%>
<!--[if IE 6]>
<script src="<%=GetCurrentHost%>zb_system/script/iepng.js" type="text/javascript"></script>
<script type="text/javascript">
   GtPNG.fix('div, ul, img, li, input, span, a');  //EvPNG.fix('包含透明PNG图片的标签'); 多个标签之间用英文逗号隔开。
</script>
<![endif]-->
<%End If%>
<%=Response_Plugin_Admin_Header%>