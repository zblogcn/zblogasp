<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
Call ZBQQConnect_Initialize()

Call CheckReference("")
If CheckPluginState("ZBQQConnect")=False Then Call ShowError(48)
If BlogUser.Level>1 Then Response.End

BlogTitle="ZBQQConnect-插件配置"

%>
<%=ZBQQConnect_Config.Load("ZBQQConnect")%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
function showqk(){
	$("#how").toggleClass('hidden')
}
</script>
<script type="text/javascript" language="javascript" src="../../../zb_system/ADMIN/ueditor/third-party/SyntaxHighlighter/shCore.js"></script>
<link rel="stylesheet" href="../../../zb_system/ADMIN/ueditor/third-party/SyntaxHighlighter/shCoreDefault.css"/>
<style type="text/css">
input[type="text"] {
	width: 90%
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
<div id="ShowBlogHint">
<%Call GetBlogHint()%>
</div>
<div class="divHeader">ZBQQConnect</div>
<div class="SubMenu"><%=ZBQQConnect_SBar(3)%></div>
<form id="form1" name="form1" method="post" action="savesetting.asp">
<div id="divMain2">
<div class="content-box"><!-- Start Content Box -->
<div class="content-box-header">
<ul class="content-box-tabs">
<li><a href="#tab0" class="default-tab"><span>全局配置</span></a></li>
<li><a href="#tab1"><span>同步设置</span></a></li>
<li><a href="#tab2"><span>评论设置</span></a></li>
<li><a href="#tab3"><span>QQ登陆设置</span></a></li>
<li><a href="#tab5"><span>头像设置</span></a></li>
<li><a href="#tab4"><span>关于</span></a></li>
</ul>
<div class="clear"></div>
</div>
<!-- End .content-box-header -->

<div class="content-box-content" id="qqcbox">
<div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab0">
<table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
<tr>
<td width="100%"><p>QQ登录APP ID:</p>
<p>
<input name="AppId" type="text" id="ap" value="<%=ZBQQConnect_Config.Read("AppID")%>"/>
</p>
<p>QQ登录KEY:</p>
<p>
<input name="KEY" type="text" id="as" value="<%=ZBQQConnect_Config.Read("KEY")%>"/>
</p>
<p><a href="http://www.zsxsoft.com/archives/231.html#how" target="_blank">如何获得？</a></p>

</td>
</tr>
</table>
<div class="clear"></div>
</div>
<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab1">
<table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
<tr>
<td width="100%"><p></p>
<p>
<input name="a" id="a" type="checkbox" <%=d(ZBQQConnect_Config.Read("a"))%> />
<label for="a"><font color="#0000CC">发表文章时默认同步到QQ空间</font></label>
</p>
<p>
<input name="b" id="b" type="checkbox" <%=d(ZBQQConnect_Config.Read("b"))%> />
<label for="b"><font color="#009900">发表文章时默认同步到腾讯微博</font></label>
</p>
<p>
<input name="c" id="c" type="checkbox" <%=d(ZBQQConnect_Config.Read("c"))%> />
<label for="c">同步时自带文章第一张图片</label>
</p>
<p><font color="#009900">同步到微博内容（%i——截断字数的摘要；%b——博客名称；%u——文章地址；%t——文章标题）</font></p>
<p>
<label for="content"></label>
<input type="text" name="content" id="content" value="<%=ZBQQConnect_Config.Read("content")%>"/>
</p>
<p></p></td>
</tr>
</table>
</div>
<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
<table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
<tr>
<td width="100%"><p></p>
<p>
<input name="d" id="d" type="checkbox" <%=d(ZBQQConnect_Config.Read("d"))%> />
<label for="d">发表评论时自动同步</label>
</p>
<p>
<input name="e" id="e" type="checkbox" <%=d(ZBQQConnect_Config.Read("e"))%> />
<label for="e"><font color="#0000CC">以分享的方式同步到QQ空间</font></label>
</p>
<p>
<input name="f" id="f" type="checkbox" <%=d(ZBQQConnect_Config.Read("f"))%> />
<label for="f"><font color="#009900">以评论的方式同步到腾讯微博（若该文章没有同步到微博则评论也不同步）</font></label>
</p>
<p>
<input name="g" id="g" type="checkbox" <%=d(ZBQQConnect_Config.Read("g"))%> />
<label for="g"><font color="#0000CC">使用管理员的身份同步到空间（无论有无勾选，若用户没有绑定QQ，则使用管理员的身份同步。微博暂只能使用管理员的身份同步）</font></label>
</p>
<p> <font color="#009900">同步到微博内容（%a—评论作者；%c——评论摘要）</font> </p>
<p>
<label for="pl"></label>
<input type="text" name="pl" id="pl" value="<%=ZBQQConnect_Config.Read("pl")%>"/>
</p>
<p></p></td>
</tr>
</table>
</div>
<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
<table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
<tr>
<td width="100%"><p></p>
<p>
<input name="h" id="h" type="checkbox" <%=d(ZBQQConnect_Config.Read("h"))%> />
<label for="h"><font color="#0000CC">允许使用QQ登录</font></label>
</p>
<p>
<input name="i" id="i" type="checkbox" <%=d(ZBQQConnect_Config.Read("i"))%> />
<label for="i"><font color="#0000CC">允许使用QQ注册帐号（不过必须先启用注册组件）</font></label>
</p>
<p>注：如果您不启用注册组件，那QQ登录也没什么用处了。</p></td>
</tr>
</table>
</div>
<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab5">
<table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
<tr>
<td width="100%"><p></p>
<p>
<label>
<input type="radio" name="a1" value="0" id="a1_0"<%=e(ZBQQConnect_Config.Read("a1"),0)%>/>
<font color="#0000CC"> 评论显示微博头像（若无，则使用系统自带）</font></label>
</p>
<p>
<label>
<input type="radio" name="a1" value="1" id="a1_1"<%=e(ZBQQConnect_Config.Read("a1"),1)%>/>
<font color="#0000CC">评论显示空间头像（若无，则使用系统自带）</font></label>
</p>
<p>
<label>
<input type="radio" name="a1" value="2" id="a1_2"<%=e(ZBQQConnect_Config.Read("a1"),2)%>/>
评论使用系统自带头像</label>
</p>
<p>注：评论内需要有<#article/comment/avatar#>才可以使用本功能！</p>
<p></p></td>
</tr>
</table>
</div>
<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab4">
<table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
<tr>
<td width="100%"><p>插件名：ZBQQConnect</p>
<p>插件版本:v1.0</p>
<p>插件制作者：ZSXSOFT</p>
<p>微博：<a href="http://t.qq.com/zhengshixin163" target="_blank">http://t.qq.com/zhengshixin163</a></p>
<p>网站：<a href="http://www.zsxsoft.com" target="_blank">http://www.zsxsoft.com</a></p></td>
</tr>
</table>
</div>
</div>
<!-- content-box-content --> 

</div>
<!-- content-box -->
<p>蓝色字体部分为需要QQ登录，绿色为需要登录微博。其他颜色则为不需要。</p>
<p>
<input type="submit" value="提交" class="button"/>
</p>
</div>
<!-- divMain2 -->

</form>
</div>
<script type="text/javascript">

SyntaxHighlighter.all();

</script>
<%
function d(v)
	d=iif(v="True"," checked=""checked"" ","")
end function
function e(s,b)
	e=iif(cint(s)=b," checked=""checked"" ","")
end function
%>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->