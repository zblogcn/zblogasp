<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%

'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:   大猪(myllop)
'// 版权所有:    www.izhu.org
'// 技术支持:    myllop@qq.com
'// 程序名称:    大猪淘淘
'// 程序版本:    1.0
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>


<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->

<!-- #include file="include.asp" -->
<!-- #include file="function.asp" -->

<%
Call System_Initialize()

LoadGlobeCache
Dim ArtList
Set ArtList=New TArticleList

ArtList.LoadCache

ArtList.template="SINGLE"

ArtList.Title="大猪淘淘"

Dim id
Dim t_content
Dim taotao
Dim strTagCloud()
Dim i,j
Dim headstr'定义加载css样式
dim adc1,adc2,adc3,adc4
dim t_rndnumber , t_rndName
t_rndnumber = RndNumber(1,7)
t_rndName = rndName(t_rndnumber)



id=request.QueryString("id")
if not isnumeric(id) then 
taotao = "没有找到您要查看的信息，回去吧，骚年"
end if


taotao = "<link rel=""stylesheet"" type=""text/css"" media=""all"" href="""&ZC_BLOG_HOST&"plugin/dztaotao/images/css.css"" />"&vbcrlf
taotao = taotao & "<link rel=""Stylesheet"" href="""&ZC_BLOG_HOST&"plugin/dztaotao/uploadify.css"" />"&vbcrlf
taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"plugin/dztaotao/swfobject.js""></script>"&vbcrlf
taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"plugin/dztaotao/jquery.uploadify.js""></script>"&vbcrlf
taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"plugin/dztaotao/images/artZoom.js""></script>"&vbcrlf

if is_release = BlogUser.Level then

'发表淘淘表单
taotao = taotao & "<form id=""form1"" enctype=""multipart/form-data"" method=""post"" ><div class=""dialog"" id=""dialog"" style="""">	<div class=""trans-box"" id=""dialogBoxtalk"">    	<div class=""dialog-title""><img height=""23"" width=""165"" src=""/PLUGIN/dztaotao/images/img-talk.png""><div class=""dialog-talktip""></div><a id=""dialogClose"" class=""dialog-close"" onfocus=""this.blur()"" href=""javascript:void(0);"" onclick=""closeDialog();return false;""></a></div><div id=""msg"" style=""display:""></div><div id=""deldiv""></div><textarea id=""s_content"" class=""comment-textarea"" rows="""" cols="""" name=""s_content"" style=""color: rgb(153, 153, 153);""></textarea><input type=""hidden"" name=""u_img"" id=""u_img""><input type=""hidden"" name=""s_img"" id=""s_img""><div><input type=""file"" name=""uploadify"" id=""uploadify"" /><a href=""javascript:$('#uploadify').uploadifyUpload()"" style=""display:none"">上传</a> <a href=""javascript:$('#uploadify').uploadifyClearQueue()"" style=""display:none""> 取消上传</a><div id=""fileQueue""></div></div><div class=""dialog-set""><span class=""talk-label"">昵称：<input type=""text"" value=""匿名"" id=""username"" class=""label-txt"" name=""username"" style=""color:#999;""><input type=""text"" style=""display:none;""></span>    <div class=""btn-talk""><span>博客：<input type=""text"" value="""" id=""s_site"" class=""label-txt"" name=""s_site"" style=""color:#999;""></span><a class=""btn-dialog submit"" href=""javascript:;"" onclick=""subInfo()"">发表</a></div>    <br clear=""all""> </div> <div class=""pink-con""> <p><span class=""highlight"">备注：</span>给我们讲一个，让我们和你一起乐哈哈~</p><p>您发表的内容我们会进行审核，正文中包含链接地址，广告，垃圾信息，政治相关或色情描写的内容将会被删除。</p> </div>    </div></div></form>"&vbcrlf


'发表按钮
taotao = taotao & "<div class=""btnTablk-box""><object height=""100"" width=""120"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""><param value="""&ZC_BLOG_HOST&"PLUGIN/dztaotao/images/talk.swf"" name=""movie""><param value=""high"" name=""quality""><param value=""transparent"" name=""wmode""><embed height=""100"" width=""120"" wmode=""transparent"" type=""application/x-shockwave-flash"" pluginspaging=""http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"" quality=""high"" src="""&ZC_BLOG_HOST&"PLUGIN/dztaotao/images/talk.swf""></object><a onfocus=""this.blur()"" id=""btnTalk"" onclick=""showDialog();return false;"" href=""javascript:;""></a></div>"&vbcrlf

end if

'banner条
taotao = taotao & "<div id=""banner"" class=""banner""><div id=""switch_img"" class=""switch_img"">"
if ads_img1 <> "" and ads_url1 <> "" then
taotao = taotao & "<a target=""_blank"" href="""&ads_url1&""" style=""z-index: 0; left: 650px;""><img src="""&ads_img1&""" width=""660"" height=""90""></a>"
adc1 = "{l:"""&ads_url1&""",s:"""&ads_img1&"""}"
end if
if ads_img2 <> "" and ads_url2 <> "" then
taotao = taotao & "<a target=""_blank"" href="""&ads_url2&""" style=""z-index: 1; left: 0px;""><img  width=""660"" height=""90""  src="""&ads_img2&"""></a>"
adc2 = ",{l:"""&ads_url2&""",s:"""&ads_img2&"""}"
end if
if ads_img3 <> "" and ads_url3 <> "" then
taotao = taotao & "<a target=""_blank"" href="""&ads_url3&""" style=""z-index: 0; left: 650px;""><img src="""&ads_img3&""" width=""660"" height=""90""></a>"
adc3 = ",{l:"""&ads_url3&""",s:"""&ads_img3&"""}"
end if
if ads_img4 <> "" and ads_url4 <> "" then
taotao = taotao & "<a target=""_blank"" href="""&ads_url4&""" style=""z-index: 0; left: 650px;""><img src="""&ads_img4&""" width=""660"" height=""90""></a>"
adc4 = ",{l:"""&ads_url4&""",s:"""&ads_img4&"""}"
end if
taotao = taotao & "</div><span id=""switchButton"" class=""switch_button"">"
if ads_img1 <> "" and ads_url1 <> "" then taotao = taotao & "<a href=""###"" class="""">1</a>" end if
if ads_img2 <> "" and ads_url2 <> "" then taotao = taotao & "<a href=""###"" class=""current"">2</a>" end if
if ads_img3 <> "" and ads_url3 <> "" then taotao = taotao & "<a href=""###"" class="""">3</a>" end if
if ads_img4 <> "" and ads_url4 <> "" then taotao = taotao & "<a href=""###"" class="""">4</a>" end if
taotao = taotao & "</span></div>"&vbcrlf

'发表成功后插入新增内容
taotao = taotao & "<div id=""newInsert""></div>"


Dim objRS
dim r_rs,img
dim r_recordcount
Set objRS=objConn.Execute("SELECT * FROM [dz_taotao] where [id]="&id&"")
If (Not objRS.bof) And (Not objRS.eof) Then

	if objRS("img")<>"" then img = "<img src=""upload/"&objRS("s_img")&""">" else img = ""  end if

	t_content=objRS("content")
	
		taotao = taotao & "<div id=""item-"&objRS("id")&""" class=""item""><div class=""item-list""><div id=""listText-"&objRS("id")&""" class=""list-text"">"&UBBCode(objRS("content"),"[face][link][autolink][font][code][image][media][flash]")&"<br><a class=""miniImg artZoom"" rel=""upload/"&objRS("img")&""" href=""upload/"&objRS("img")&""">"&img&"</a></div><div class=""list-text""><div class=""list-interaction""> "&vbcrlf & vbcrlf
		
		
		'分享代码
		taotao = taotao & "<div id=""shareLayer"&objRS("id")&""" class=""share-layer"" style=""display:""><dl class=""item-share""><dt>分享到:</dt><dd><a href=""http://service.weibo.com/share/share.php?url="&ZC_BLOG_HOST&"plugin/dztaotao/view.asp?id="&objRS("id")&"&type=3&count=&appkey=&title="&server.URLEncode("大猪淘淘——"&left(objRS("content"),130))&"&pic="&ZC_BLOG_HOST&"plugin/dztaotao/upload/"&objRS("img")&"&ralateUid=&rnd=1337756006442"" target=""_blank"" title=""转帖到新浪微博"" id=""share_sina"" class=""btn-share-sina""></a></dd><dd><a href=""http://share.renren.com/share/buttonshare.do?link="&ZC_BLOG_HOST&"plugin/dztaotao/view.asp?id="&objRS("id")&"&title="&server.URLEncode("大猪淘淘——"&left(objRS("content"),130))&""" target=""_blank"" title=""转帖到人人网"" class=""btn-share-rr""></a></dd><dd><a href=""###"" onclick=""open_share('kx','"&objRS("id")&"')"" title=""转帖到开心网"" id=""share_kx"" class=""btn-share-kx""></a></dd><dd><a href=""http://share.v.t.qq.com/index.php?c=share&a=index&appkey=&site="&ZC_BLOG_HOST&"&pic="&ZC_BLOG_HOST&"plugin/dztaotao/upload/"&objRS("img")&"&title="&server.URLEncode("大猪淘淘——"&left(objRS("content"),120))&"&url="&ZC_BLOG_HOST&"plugin/dztaotao/view.asp?id="&objRS("id")&""" target=""_blank"" title=""推荐到QQ微博"" id=""share_tqq"" class=""btn-share-tqq""></a></dd></dl></div>"&vbcrlf & vbcrlf
		
		taotao = taotao & "           </div><div class=""clear""></div></div></div><div class=""item-infor""><div class=""infor-text""><img src=""/PLUGIN/dztaotao/images/default.jpg""> <span>"&objRS("username")&"</span> <span>"&objRS("addtime")&" 发布</span></div><div class=""infor-set""><a onclick=""dingUp("&objRS("id")&")"" class=""btn-up"" onfocus=""this.blur()"" href=""javascript:;"">称赞</a> <span id=""ding_"&objRS("id")&""" class=""scroe-up highlight"">"&objRS("ttop")&"</span> <a onclick=""dingDown("&objRS("id")&")"" class=""btn-down"" onfocus=""this.blur()"" href=""javascript:;"">鄙视</a> <span class=""scroe-down highlight"" id=""tread_"&objRS("id")&""">"&objRS("tread")&"</span> | <a onclick=""showReply("&objRS("id")&")"" class=""comment"" id=""commtent-"&objRS("id")&""" onfocus=""this.blur()"" title=""点击展开评论"" href=""javascript:;"">评论("&objRS("comments")&")</a></div></div>"&vbcrlf
		
		'发表评论框
		taotao = taotao &"<div class=""item-comment"" style=""background:#fff;"" >"
		
		taotao = taotao & "<div id=""blueCon-"&objRS("id")&""" class=""blue-con"" style=""padding: 10px 10px 0pt;"">"&vbcrlf
		
		taotao = taotao & "<table border=""0""><tr><td>发表评论<div id=""shortcut-key"&objRS("id")&"""></div></td></tr><tr><td><textarea id=""r_content_"&objRS("id")&""" class=""comment-textarea"" style=""height:80px;"" name=""r_content_"&objRS("id")&"""></textarea></td></tr>   <tr style=""display:none;""><td>昵称：<input type=""text"" name=""r_username_"&objRS("id")&""" id=""r_username_"&objRS("id")&""" value="""&t_rndName&""">    邮箱：<input type=""text"" name=""r_email_"&objRS("id")&""" id=""r_email_"&objRS("id")&""">    网址：<input type=""text"" name=""r_site_"&objRS("id")&""" id=""r_site_"&objRS("id")&"""></td></tr></table>"&vbcrlf
		

		taotao = taotao & "<div class=""discuss-login""><a id=""send-"&objRS("id")&""" class=""btn-send"" href=""javascript:;"" onclick=""postCmt("&objRS("id")&")"">发表评论</a><span class=""comments-leave"">最好不要超过200个字符</span></div></div>"&vbcrlf

		
		'发表评论结束
		taotao = taotao &"</div>"
		
		'广告
		taotao = taotao &"<div class=""item-comment"" style=""background:#fff;"" >"
		taotao = taotao & "<div id=""item-ad1"" class=""item""><div style=""margin:5px 6px;"">"&content_ads&"</div></div></div>"&vbcrlf

		

		
		taotao = taotao & "<div class=""item-comment"" style=""background:#fff;"" id=""item-comment-"&objRS("id")&"""><div class=""clear""></div>"&vbcrlf
		
		
		
		
		
		taotao = taotao & "<div id=""msg-"&objRS("id")&""" class=""comment-msg""></div>          "&vbcrlf
		
		'评论列表
		taotao = taotao & "<ul id=""comments-order-401083"" class=""comments-order""> <li class=""current"" order=""time""><a href=""###"">评论列表</a></li>                   </ul>"
		
		taotao = taotao & "<div id=""comments-"&objRS("id")&""" class=""comment-list"">"&vbcrlf & vbcrlf
		
		

		
		'用来放置新插入评论
		taotao = taotao & "<div id=""newInsertCmt"&objRS("id")&"""></div>"
		
		set r_rs=objConn.execute("select * from dz_comment where tt_id = "&objRS("id")&"")
		if not r_rs.eof then
		do while not r_rs.eof
		taotao = taotao & "<!--comment start--><div id=""jitem-"&r_rs("id")&""" class=""item""><div class=""comment-box""><a href="""&r_rs("u_site")&""" class=""discuss-pic""><img height=""32"" width=""32"" src=""http://passport.maxthon.cn/_image/avatar-demo.png""></a><div class=""discuss-con""><div class=""con-bar dash-boder""><a href="""&r_rs("u_site")&""" class=""name"">"&r_rs("u_sername")&"</a><span class=""time"">"&r_rs("addtime")&"发表</span> </div><p>"&r_rs("content")&"</p></div><div class=""clear""></div></div></div><!-- end comment-->"
		r_rs.movenext
		loop
		end if
		r_rs.close:set r_rs=nothing

		taotao = taotao & "</div><div id=""all-"&objRS("id")&""" class=""comment-all"">共有"&objRS("comments")&"条评论 | <a href=""javascript:;"" onclick=""showReply("&objRS("id")&")"">收起评论</a></div></div></div>"

    Else
		taotao = "<div>暂无数据</div>"
End If
objRS.Close
Set objRS=Nothing


taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"plugin/dztaotao/core.js""></script>"&vbcrlf & vbcrlf


taotao = taotao & "<script>var banner=(function(){var b=["&adc1&adc2&adc3&adc4&"];var a=function(c){this.build(b);this.setOptions(c);this.oSwithButton=this.options.oSwithButton;this.oSwithImg=this.options.oSwithImg;this.iSwithButton=$(this.oSwithButton).find(""a"");this.iSwithImg=$(this.oSwithImg).find(""a"");this.timer=null;this.init();var d=this;$(this.oSwithButton).bind(""mouseover"",function(){d.stop();}).bind(""mouseout"",function(){d.autoButton(""auto"");});this.clickButton();this.autoButton(""auto"");};a.prototype={build:function(f){var c=f.length;var h="""";var g="""";var d=0;for(var e=0;e<c;e++){d=e+1;if(f[e].l===""""){h+='<a href=""###""><img width=""660"" height=""85"" src=""'+f[e].s+'"" /></a>';}else{h+='<a href=""'+f[e].l+'"" target=""_blank""><img width=""660"" height=""85"" src=""'+f[e].s+'"" /></a>';}g+='<a href=""###"">'+d+""</a>"";}$(""#banner"").html('<div class=""switch_img"" id=""switch_img"">'+h+'</div><span class=""switch_button"" id=""switchButton"">'+g+""</span>"");},setOptions:function(c){this.options={oSwithButton:""#switchButton"",oSwithImg:""#switch_img""};$.extend(this.options,c||{});},init:function(){$(this.iSwithButton[0]).addClass(""current"");$(this.iSwithImg[0]).css({""z-index"":1,left:0});},clickButton:function(){var d=this;for(var c=0;c<this.iSwithButton.length;c++){(function(){var e=c;$(d.iSwithButton[e]).click(function(){if($(d.iSwithButton[e]).attr(""class"")==""current""){return;}for(var f=0;f<c;f++){if(e==f){$(this).addClass(""current"");$(d.iSwithImg[e]).css({""z-index"":3});$(d.iSwithImg[e]).animate({left:""-=650px""},{duration:500,complete:function(){for(var g=0;g<f;g++){g==e?$(this).css({""z-index"":1}):$(d.iSwithImg[g]).css({""z-index"":0,left:""650px""});}}});}else{$(d.iSwithImg[f]).stop(true,true);$(d.iSwithButton[f]).removeClass();}}});})();}},pointer:function(){for(var c=0;c<this.iSwithButton.length;c++){if($(this.iSwithButton[c]).attr(""class"")==""current""){return c;}}},action:function(e,f){var h=this;var g=this.pointer();switch(e.toLowerCase()){case""right"":if(g>=(h.iSwithButton.length-1)){g=-1;}break;case""left"":if(g<=0){g=h.iSwithButton.length;}break;}var c=g+f;for(var d=0;d<h.iSwithButton.length;d++){if(d==(c)){$(h.iSwithButton[c]).addClass(""current"");$(h.iSwithImg[c]).css({""z-index"":3});$(h.iSwithImg[c]).animate({left:""-=650px""},{duration:500,complete:function(){for(var i=0;i<d;i++){i==(c)?$(this).css({""z-index"":1}):$(h.iSwithImg[i]).css({""z-index"":0,left:""650px""});}}});}else{$(h.iSwithImg[d]).stop(true,true);$(h.iSwithButton[d]).removeClass();}}},stop:function(){clearTimeout(this.timer);},autoButton:function(){var c=this;if(arguments[0]==""auto""){this.timer=window.setTimeout(function(){c.autoButton(""auto"");c.action(""right"",1);},3000);}}};return{init:function(){var c=new a();}};})();$(document).ready(function(){banner.init();});</script>"&vbcrlf & vbcrlf


ArtList.SetVar "template:article-single",taotao
	
ArtList.html = replace(ArtList.html,"<#article/title#>",""&left(t_content,40)&"...")

ArtList.Build

Response.Write ArtList.html
%><!-- <%=RunTime()%>ms --><%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>