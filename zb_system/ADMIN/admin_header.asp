<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="robots" content="nofollow" />
<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> 
<link href="<%=ZC_BLOG_HOST%>ZB_SYSTEM/CSS/admin2.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="<%=ZC_BLOG_HOST%>ZB_SYSTEM/script/common.js" type="text/javascript"></script>
<title><%=BlogTitle%></title>
<!--[if IE 6]>
<script src="<%=ZC_BLOG_HOST%>ZB_SYSTEM/Script/iepng.js" type="text/javascript"></script>
<script type="text/javascript">
   EvPNG.fix('div, ul, img, li, input, span, a');  //EvPNG.fix('包含透明PNG图片的标签'); 多个标签之间用英文逗号隔开。
</script>
<![endif]-->
<!--高亮JS-->
<script type="text/javascript">
function currentpage(){
    if(!document.getElementsByTagName) return false;
    if(!document.getElementById) return false;
    if(!document.getElementById('nav')) return false;
    var nav = document.getElementById('nav');
    var links = nav.getElementsByTagName('a');
    for (var i=0;i<links.length;i++){
        var linkurl =  links[i].getAttribute('href');
        var currenturl = document.location.href;
            if(currenturl.indexOf(linkurl)!=-1){
                links[i].className = 'on';
                return true;
            }
    }
}
</script>
