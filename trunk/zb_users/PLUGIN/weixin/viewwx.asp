<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.CacheControl="no-cache" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%
Call System_Initialize()

If Request("wid") Then
	Dim viewid,ZC_PU_WEIXIN_TITLE,ZC_PU_WEIXIN_TIME,ZC_PU_WEIXIN_POSTER,ZC_PU_WEIXIN_CONTENT,ZC_PU_WEIXIN_INTRO
	viewid = Request("wid")
	
	Dim objArticle
	Set objArticle=New TArticle
	
	If objArticle.LoadInfoByID(viewid) Then
		ZC_PU_WEIXIN_TIME = objArticle.PostTime
		ZC_PU_WEIXIN_TITLE = objArticle.TITLE
		'ZC_PU_WEIXIN_POSTER = Users(objArticle.AuthorID).Name
		ZC_PU_WEIXIN_CONTENT = objArticle.Content
		ZC_PU_WEIXIN_INTRO = TransferHTML(objArticle.Intro)
	End If
Else
	Response.End()
End If
%>

<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd"> <html xmlns="http://www.w3.org/1999/xhtml"><head><title><%=ZC_PU_WEIXIN_TITLE%></title><meta http-equiv=Content-Type content="text/html;charset=utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0"><meta name="apple-mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-status-bar-style" content="black"><meta name="format-detection" content="telephone=no"><style type="text/css">html{background:#FFF;color:#000}body,div,dl,dt,dd,h1,h2,h3,h4,h5,h6,pre,code,form,fieldset,legend,input,textarea,p,blockquote,th,td{margin:0;padding:0}table{border-collapse:collapse;border-spacing:0}fieldset,img{border:0}address,caption,cite,code,dfn,th,var{font-style:normal;font-weight:normal}ol,ul{list-style:none}caption,th{text-align:left}h1,h2,h3,h4,h5,h6{font-size:100%;font-weight:normal}q:before,q:after{content:''}abbr,acronym{border:0;font-variant:normal}sup{vertical-align:text-top}sub{vertical-align:text-bottom}input,textarea,select{font-family:inherit;font-size:inherit;font-weight:inherit}input,textarea,select{font-size:100%}legend{color:#000}html{background-color:#f8f7f5}body{background:#f8f7f5;color:#222;font-family:Helvetica,STHeiti STXihei,Microsoft JhengHei,Microsoft YaHei,Tohoma,Arial;height:100%;padding:15px 15px 0;position:relative}body>.tips{display:none;left:50%;padding:20px;position:fixed;text-align:center;top:50%;width:200px;z-index:100}.page{padding:15px}.page .page-error,.page .page-loading{line-height:30px;position:relative;text-align:center}.btn{background-color:#fcfcfc;border:1px solid #ccc;border-radius:5px;box-shadow:0 1px 4px rgba(0,0,0,0.3);color:#222;cursor:pointer;display:block;font-size:15px;font-weight:bold;margin:15px 0;moz-box-shadow:0 1px 4px rgba(0,0,0,0.3);padding:10px;text-align:center;text-decoration:none;webkit-box-shadow:0 1px 4px rgba(0,0,0,0.3)}.icons{background:url(../images/icons151fd8.png) no-repeat 0 0;border-radius:5px;height:25px;overflow:hidden;position:relative;width:25px}.icons.arrow-r{background:url(../images/brand_profileinweb_arrow@2x151fd8.png) no-repeat center center;background-size:100%;height:16px;width:12px}.icons.check{background-position:-25px 0}#activity-detail .page-bizinfo .header #activity-name{color:#000;font-size:20px;font-weight:bold;word-break:normal;word-wrap:break-word}#activity-detail .page-bizinfo .header #post-date{color:#8c8c8c;font-size:11px;margin:0}#activity-detail .page-bizinfo #biz-link.btn{background:url(../images/brand_profileinweb_bg@2x151fd8.png) no-repeat center center;background-size:100% 100%;border:0;border-radius:0;box-shadow:none;height:42px;padding:12px;padding-left:62px;position:relative;text-align:left}#activity-detail .page-bizinfo #biz-link.btn:hover{background-image:url(../images/brand_profileinweb_bg_HL@2x151fd8.png)}#activity-detail .page-bizinfo #biz-link.btn .arrow{position:absolute;right:15px;top:25px}#activity-detail .page-bizinfo #biz-link.btn .logo{height:42px;left:5px;overflow:hidden;padding:6px;position:absolute;top:6px;width:42px}#activity-detail .page-bizinfo #biz-link.btn .logo img{position:relative;width:42px;z-index:10}#activity-detail .page-bizinfo #biz-link.btn .logo .circle{background:url(../images/brand_photo_middleframe@2x151fd8.png) no-repeat center center;background-size:100% 100%;height:54px;left:0;position:absolute;top:0;width:54px;z-index:100}#activity-detail .page-bizinfo #biz-link.btn #nickname{color:#454545;font-size:15px;text-shadow:0 1px 1px white}#activity-detail .page-bizinfo #biz-link.btn #weixinid{color:#a3a3a3;font-size:12px;line-height:20px;text-shadow:0 1px 1px white}#activity-detail .page-content{margin:18px 0 0;padding-bottom:18px}#activity-detail .page-content .media{margin:18px 0}#activity-detail .page-content .media img{width:100%}#activity-detail .page-content .text{color:#3e3e3e;line-height:1.5;width:100%;overflow:hidden;zoom:1}#activity-detail .page-content .text p{min-height:1.5em;min-height:1.5em}#activity-list .header{font-size:20px}#activity-list .page-list{border:1px solid #ccc;border-radius:5px;margin:18px 0;overflow:hidden}#activity-list .page-list .line.btn{border-radius:0;margin:0;text-align:left}#activity-list .page-list .line.btn .checkbox{height:25px;line-height:25px;padding-left:35px;position:relative}#activity-list .page-list .line.btn .checkbox .icons{background-color:#ccc;left:0;position:absolute;top:0}#activity-list .page-list .line.btn.off .icons{background-image:none}.vm{vertical-align:middle}.tc{text-align:center}.db{display:block}.dib{display:inline-block}.b{font-weight:700}.clr{clear:both}.text img{max-width:100%!important;height:auto!important}.page-url{padding-top:18px}.page-url-link{color:#607fa6;font-size:14px;text-decoration:none;text-shadow:0 1px #fff;-webkit-text-shadow:0 1px #fff;-moz-text-shadow:0 1px #fff}#nickname{overflow: hidden;white-space: nowrap;text-overflow: ellipsis;max-width: 90%;}ol,ul{list-style-position:inside;}</style></head> 
<body id="activity-detail"><div class="page-bizinfo"><div class="header"><h1 id="activity-name"><%=ZC_PU_WEIXIN_TITLE%></h1><span id="post-date"><%=ZC_PU_WEIXIN_TIME%></span></div></div><div class="page-content"><div class="text"><%=ZC_PU_WEIXIN_CONTENT%></div>
<!--<p class="page-url"><a href="javascript:void(0)" onclick="viewSource()" class="page-url-link">阅读原文<br/> 本文章内容为《<%=ZC_BLOG_TITLE%>》通过 Z-Blog微信插件生成，以供文章在微信上查看获得最佳体验。</a></p>-->
</div>
<!--<script src="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/js/jquery-1.8.3.min1530d1.js"></script>
	<script src="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/js/wxm-core1530d0.js"></script>
	<script id="txt-desc" type="txt/text"><%=ZC_PU_WEIXIN_INTRO%></script>
	<script id="txt-title" type="txt/text"><%=ZC_PU_WEIXIN_TITLE%></script>
	<script id="txt-sourceurl" type="txt/text"><%=bloghost%>ZB_USERS/plugin/weixin/view.asp?wid=<%=viewid%>#wechat_redirect</script>
    <script>(function(){var b=jQuery(".text"),a=b.html();_em=jQuery("<p></p>").html("a").css({display:"inline"}),_init=function(){_em.appendTo(b);var d=a,c=Math.floor(b.width()/_em.width()),e=new RegExp("[a-z1-9]{"+c+",}","ig");_em.remove();d=d.replace(/>[^<]+</g,function(f){return f.replace(e,function(h){var i=h,g=[];while(i.length>c){g.push(i.substr(0,c));i=i.substr(c)}g.push(i);return g.join("<br/>")})});b.html(d)};jQuery(window).on("resize",_init).trigger("resize")})();function getStrFromTxtDom(a){return jQuery("#txt-"+a).html().replace(/&lt;/g,"<").replace(/&gt;/g,">")}function viewSource(){var a=navigator.userAgent.toLowerCase();var b=function(){if(/IEMobile/i.test(a)){return true}else{return false}};if(b()){jQuery(".page-url-link:first").attr("href",getStrFromTxtDom("sourceurl"));return}jQuery.ajax({url:"/mp/appmsg/show-ajax"+location.search,async:false,type:"POST",timeout:2000,data:{url:getStrFromTxtDom("sourceurl")},complete:function(){location.href=getStrFromTxtDom("sourceurl")}});return false};</script>-->
	</body>
</html>
