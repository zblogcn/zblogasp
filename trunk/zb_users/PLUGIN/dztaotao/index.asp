<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:   大猪(myllop)
'// 版权所有:    www.izhu.org
'// 技术支持:    myllop@qq.com
'// 程序名称:    大猪滔滔
'// 程序版本:    1.0
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="function.asp" -->
<%
Call System_Initialize()

LoadGlobeCache

Call dztaotao_Initialize()

Dim ArtList
Set ArtList=New TArticle

ArtList.LoadCache

'ArtList.template="page"

ArtList.Title="滔滔"

Dim taotao
Dim strTagCloud()
Dim i,j
Dim headstr'定义加载css样式
dim adc1,adc2,adc3,adc4
dim t_rndnumber , t_rndName

if BlogUser.Name = "来宾" then
t_rndnumber = RndNumber(1,7)
t_rndName = rndName(t_rndnumber)
else
t_rndName = BlogUser.Name
end if



taotao = "<link rel=""stylesheet"" type=""text/css"" media=""all"" href="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/images/css.css"" />"&vbcrlf
taotao = taotao & "<link rel=""Stylesheet"" href="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/uploadify.css"" />"&vbcrlf
taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/swfobject.js""></script>"&vbcrlf
taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/jquery.uploadify.js""></script>"&vbcrlf
taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/images/artZoom.js""></script>"&vbcrlf


if int(DZTAOTAO_RELEASE_VALUE) = cint(BlogUser.Level) or cint(DZTAOTAO_RELEASE_VALUE) = 5 then

'发表淘淘表单
taotao = taotao & "<form id=""form1"" enctype=""multipart/form-data"" method=""post"" ><div class=""dialog"" id=""dialog"" style="""">	<div class=""trans-box"" id=""dialogBoxtalk"">    	<div class=""dialog-title""><img height=""23"" width=""165"" src="""&ZC_BLOG_HOST&"zb_users/PLUGIN/dztaotao/images/img-talk.png""><div class=""dialog-talktip""></div><a id=""dialogClose"" class=""dialog-close"" onfocus=""this.blur()"" href=""javascript:void(0);"" onclick=""closeDialog();return false;""></a></div><div id=""msg"" style=""display:""></div><div id=""deldiv""></div><textarea id=""s_content"" class=""comment-textarea"" rows="""" cols="""" name=""s_content"" style=""color: rgb(153, 153, 153);""></textarea>"&vbcrlf

'上传图片部分
if DZTAOTAO_ISIMG_VALUE = 1 then
taotao = taotao & "<input type=""hidden"" name=""u_img"" id=""u_img""><input type=""hidden"" name=""s_img"" id=""s_img""><div><input type=""file"" name=""uploadify"" id=""uploadify"" /><a href=""javascript:$('#uploadify').uploadifyUpload()"" style=""display:none"">上传</a> <a href=""javascript:$('#uploadify').uploadifyClearQueue()"" style=""display:none""> 取消上传</a><div id=""fileQueue""></div></div>"&vbcrlf
end if

taotao = taotao & "<div class=""dialog-set""><span class=""talk-label"">昵称：<input type=""text"" value="""&BLogUser.Name&""" id=""username"" class=""label-txt"" name=""username"" style=""color:#999;""><input type=""text"" style=""display:none;""></span>    <div class=""btn-talk""><span>博客：<input type=""text"" value="""&BlogUser.HomePage&""" id=""s_site"" class=""label-txt"" name=""s_site"" style=""color:#999;""></span><a class=""btn-dialog submit"" href=""javascript:;"" onclick=""subInfo();return false;"" id=""submit_btn"">发表</a></div>    <br clear=""all""> </div> <div class=""pink-con""> <p><span class=""highlight"">备注：</span>给我们讲一个，让我们和你一起乐哈哈~</p><p>您发表的内容我们会进行审核，正文中包含链接地址，广告，垃圾信息，政治相关或色情描写的内容将会被删除。</p> </div>    </div></div></form>"&vbcrlf


'发表按钮
taotao = taotao & "<div class=""btnTablk-box""><object height=""100"" width=""120"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""><param value="""&ZC_BLOG_HOST&"zb_users/PLUGIN/dztaotao/images/talk.swf"" name=""movie""><param value=""high"" name=""quality""><param value=""transparent"" name=""wmode""><embed height=""100"" width=""120"" wmode=""transparent"" type=""application/x-shockwave-flash"" pluginspaging=""http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"" quality=""high"" src="""&ZC_BLOG_HOST&"zb_users/PLUGIN/dztaotao/images/talk.swf""></object><a onfocus=""this.blur()"" id=""btnTalk"" onclick=""showDialog();return false;"" href=""javascript:;""></a></div>"&vbcrlf

end if


'发表成功后插入新增内容
taotao = taotao & "<div id=""newInsert""></div>"&vbcrlf
taotao = taotao & "<div id=""RequestText""></div>"&vbcrlf

'taotao循环外层
taotao = taotao & "<div id=""taotao"" style=""width:"&DZTAOTAO_PAGEWIDTH_VALUE&"px;"">"&vbcrlf

Dim objRS
dim r_rs
dim r_recordcount
Set objRS=objConn.Execute("SELECT * FROM [dz_taotao] where itype = 0 ORDER BY [id] desc")
If (Not objRS.bof) And (Not objRS.eof) Then
	Dim CurrentPage,F
	Dim TotalPut
	objRS.MoveFirst
	If Trim(Request.Querystring("Page"))<>"" Or Not IsNumeric(Request.Querystring("Page")) Then
	CurrentPage=Clng(Request.Querystring("Page"))
	Else
	CurrentPage=1
	End If
	TotalPut=objConn.ExeCute("Select Count(id) From dz_taotao where itype = 0",0,1)(0)
	If CurrentPage<>1 Then
		If (CurrentPage-1)*DZTAOTAO_PAGECOUNT_VALUE<TotalPut Then
		objRS.Move(CurrentPage-1)*DZTAOTAO_PAGECOUNT_VALUE
		End If
	End If
	Dim N,K
	dim dz_ii,dz_img
	dz_ii=1
	If (TotalPut Mod DZTAOTAO_PAGECOUNT_VALUE)=0 Then
	N=TotalPut \ DZTAOTAO_PAGECOUNT_VALUE
	Else
	N=TotalPut \ DZTAOTAO_PAGECOUNT_VALUE+1
	End If
	For F=1 To DZTAOTAO_PAGECOUNT_VALUE
	If Not objRS.Eof Then
	

	if objRS("img")<>"" then dz_img = "<img src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/upload/"&objRS("s_img")&""">" else dz_img = ""  end if
	
		taotao = taotao & "<div id=""item-"&objRS("id")&""" class=""item""><div class=""item-list""><div id=""listText-"&objRS("id")&""" class=""list-text"">"&UBBCode(objRS("content"),"[face][link][autolink][font][code][image][media][flash]")&"<br><a class=""miniImg artZoom"" rel="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/upload/"&objRS("img")&""" href="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/upload/"&objRS("img")&""">"&dz_img&"</a></div><div class=""list-text""><div class=""list-interaction""> "&vbcrlf & vbcrlf
		
		
		'分享代码
		taotao = taotao & "<div id=""shareLayer"&objRS("id")&""" class=""share-layer"" style=""display:""><dl class=""item-share""><dt>分享到:</dt><dd><a href=""http://service.weibo.com/share/share.php?url="&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/view.asp?id="&objRS("id")&"&type=3&count=&appkey=&title="&server.URLEncode("大猪淘淘——"&left(objRS("content"),130))&"&pic="&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/upload/"&objRS("img")&"&ralateUid=&rnd=1337756006442"" target=""_blank"" title=""转帖到新浪微博"" id=""share_sina"" class=""btn-share-sina""></a></dd><dd><a href=""http://share.renren.com/share/buttonshare.do?link="&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/view.asp?id="&objRS("id")&"&title="&server.URLEncode("大猪淘淘——"&left(objRS("content"),130))&""" target=""_blank"" title=""转帖到人人网"" class=""btn-share-rr""></a></dd><dd><a href=""###"" onclick=""open_share('kx','"&objRS("id")&"')"" title=""转帖到开心网"" id=""share_kx"" class=""btn-share-kx""></a></dd><dd><a href=""http://share.v.t.qq.com/index.php?c=share&a=index&appkey=&site="&ZC_BLOG_HOST&"&pic="&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/upload/"&objRS("img")&"&title="&server.URLEncode("大猪淘淘——"&left(objRS("content"),120))&"&url="&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/view.asp?id="&objRS("id")&""" target=""_blank"" title=""推荐到QQ微博"" id=""share_tqq"" class=""btn-share-tqq""></a></dd></dl></div>"&vbcrlf & vbcrlf
		
		taotao = taotao & "           </div><div class=""clear""></div></div></div><div class=""item-infor""><div class=""infor-text""><img src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/images/default.jpg"" id=""authoricon""> <span>"&objRS("username")&"</span> <span>"&objRS("addtime")&" 发布</span></div><div class=""infor-set""><a onclick=""dingUp("&objRS("id")&")"" class=""btn-up"" onfocus=""this.blur()"" href=""javascript:;"">称赞</a> <span id=""ding_"&objRS("id")&""" class=""scroe-up highlight"">"&objRS("ttop")&"</span> <a onclick=""dingDown("&objRS("id")&")"" class=""btn-down"" onfocus=""this.blur()"" href=""javascript:;"">鄙视</a> <span class=""scroe-down highlight"" id=""tread_"&objRS("id")&""">"&objRS("tread")&"</span> | <a onclick=""showReply("&objRS("id")&")"" class=""comment"" id=""commtent-"&objRS("id")&""" onfocus=""this.blur()"" title=""点击展开评论"" href=""javascript:;"">评论("&objRS("comments")&")</a></div></div><div class=""item-comment"" style=""display:none"" id=""item-comment-"&objRS("id")&"""><div class=""clear""></div>         <div id=""blueCon-"&objRS("id")&""" class=""blue-con"" style=""padding: 10px 10px 0pt;"">"&vbcrlf
		
		taotao = taotao & "<table border=""0""><tr><td><div id=""shortcut-key"&objRS("id")&"""></div></td></tr><tr><td><textarea id=""r_content_"&objRS("id")&""" class=""comment-textarea"" name=""r_content_"&objRS("id")&"""></textarea></td></tr>   <tr style=""display:;""><td>昵称：<input type=""text"" name=""r_username_"&objRS("id")&""" id=""r_username_"&objRS("id")&""" value="""&t_rndName&""">    邮箱：<input type=""text"" name=""r_email_"&objRS("id")&""" value="""&BlogUser.Email&""" id=""r_email_"&objRS("id")&""">    网址：<input type=""text"" name=""r_site_"&objRS("id")&""" id=""r_site_"&objRS("id")&""" value="""&BlogUser.HomePage&"""></td></tr></table>"&vbcrlf
		

		taotao = taotao & "<div class=""discuss-login""><a id=""send-"&objRS("id")&""" class=""btn-send"" href=""javascript:;"" onclick=""postCmt("&objRS("id")&")"">发表评论</a><span class=""comments-leave"">最好不要超过200个字符</span></div></div><div id=""msg-"&objRS("id")&""" class=""comment-msg""></div>          <div id=""comments-"&objRS("id")&""" class=""comment-list"">"&vbcrlf & vbcrlf
		
		'用来放置新插入评论
		'taotao = taotao & "<div id=""newInsertCmt"&objRS("id")&"""></div>"&vbcrlf & vbcrlf
		
		set r_rs=objConn.execute("select * from dz_comment where tt_id = "&objRS("id")&" and itype=0 order by id desc")
		if not r_rs.eof then
		do while not r_rs.eof
		taotao = taotao & "<!--comment start--><div id=""jitem-"&r_rs("id")&""" class=""item""><div class=""comment-box""><a href="""&r_rs("u_site")&""" class=""discuss-pic"">"
			if r_rs("u_email") <> "" then
			taotao = taotao & "<img height=""32"" width=""32"" src=""http://www.gravatar.com/avatar/"&MD5(r_rs("u_email"))&"?s=40&d=http%3A%2F%2Fwww.gravatar.com%2Favatar%2Fad516503a11cd5ca435acc9bb6523536%3Fs%3D40&r=G"">"&vbcrlf & vbcrlf
			else
			taotao = taotao & "<img height=""32"" width=""32"" src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/images/default.jpg"">"&vbcrlf & vbcrlf
			end if
		taotao = taotao & "</a><div class=""discuss-con""><div class=""con-bar dash-boder""><a href="""&r_rs("u_site")&""" class=""name"" target=""_blank"">"&r_rs("u_sername")&"</a><span class=""time"">"&r_rs("addtime")&"发表</span> </div><p>"&r_rs("content")&"</p></div><div class=""clear""></div></div></div><!-- end comment-->"&vbcrlf & vbcrlf
		r_rs.movenext
		loop
		end if
		r_rs.close:set r_rs=nothing

		taotao = taotao & "</div><div id=""all-"&objRS("id")&""" class=""comment-all"">共有"&objRS("comments")&"条评论 | <a href=""javascript:;"" onclick=""showReply("&objRS("id")&")"">收起评论</a> | <a href=""view.asp?id="&objRS("id")&""">更多</a></div></div></div>"&vbcrlf

	objRS.MoveNext
	End If
    Next
    Else
		taotao = "<div>暂无数据</div>"
End If
objRS.Close
Set objRS=Nothing

K=CurrentPage

taotao = taotao & "</div>"&vbcrlf

taotao = taotao & "<div class=""t_pages"">" &ExportPageBar(k,n,10,"index.asp?page=") & "</div>"&vbcrlf

taotao = taotao & "<script type=""text/javascript"" src="""&ZC_BLOG_HOST&"zb_users/plugin/dztaotao/core.js""></script>"&vbcrlf & vbcrlf


''''''''''''''''''''
ArtList.FType=ZC_POST_TYPE_PAGE
ArtList.Content=taotao
ArtList.Title=DZTAOTAO_TITLE_VALUE
ArtList.FullRegex="{%host%}/{%alias%}.html"


If ArtList.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
	ArtList.Build
	Response.Write ArtList.html
End If

%><!-- <%=RunTime()%>ms --><%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>