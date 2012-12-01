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
init_qqconnect()

Call CheckReference("")
If CheckPluginState("QQConnect")=False Then Call ShowError(48)
If BlogUser.Level=5 Then Response.End
BlogTitle="QQ互联"

%>

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
        <div class="divHeader"><%=BlogTitle%></div>
        <div class="SubMenu"><%=QQConnect.functions.navbar(1)%></div>
        <form id="form1" name="form1" method="post" action="savesetting.asp">
          <div id="divMain2">
            <div class="content-box"><!-- Start Content Box -->
              <div class="content-box-header">
                <ul class="content-box-tabs">
                  <% If BlogUser.Level=1 Then %>
				  <li><a href="#tab0" class="default-tab"><span>全局配置</span></a></li>
                  <li><a href="#tab1"><span>同步设置</span></a></li>
                  <li><a href="#tab2"><span>评论设置</span></a></li>
                  <li><a href="#tab3"><span>QQ登陆设置</span></a></li>
                  <li><a href="#tab6"><span>头像设置</span></a></li>
                  <%End If%>
                  <li><a href="#tab7"<%=IIf(BlogUser.Level=1,""," class=""default-tab""")%>><span>用户设置</span></a></li>
                  <li><a href="#tab4"><span>关于</span></a></li>
                </ul>
                <div class="clear"></div>
              </div>
              <!-- End .content-box-header -->
              
              <div class="content-box-content" id="qqcbox">
              <% If BlogUser.Level=1 Then %>
                <div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab0">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <th width="200px">配置项</th>
                      <th>内容</th>
                    <tr>
                      <td>QQ登录APP ID <a href="http://www.zsxsoft.com/archives/231.html#how" target="_blank" title="如何获得">？</a></td>
                      <td><input name="AppId" type="text" id="ap" value="<%=qqconnect.tconfig.read("AppID")%>"/></td>
                    </tr>
                    <tr>
                      <td>QQ登录KEY</td>
                      <td><input name="KEY" type="text" id="as" value="<%=qqconnect.tconfig.read("KEY")%>"/></td>
                    </tr>
                  </table>
                  <div class="clear"></div>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab1">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <th width="200px">配置项</th>
                      <th>内容</th>
                    </tr>
                    <tr>
                      <td><font color="#0000CC">发表文章时默认同步到QQ空间</font></td>
                      <td><input name="a" id="a" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("a"))%> /></td>
                    </tr>
                    <tr>
                      <td><font color="#009900">发表文章时默认同步到腾讯微博</font></td>
                      <td><input name="b" id="b" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("b"))%> /></td>
                    </tr>
                    <tr>
                      <td>同步时自带文章第一张图片</td>
                      <td><input name="c" id="c" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("c"))%> /></td>
                    </tr>
                    <tr>
                      <td>同步到微博内容</td>
                      <td><p>（%i——截断字数的摘要；%b——博客名称；%u——文章地址；%t——文章标题）</p>
                        <input type="text" name="content" id="content" value="<%=qqconnect.tconfig.read("content")%>"/></td>
                    </tr>
                  </table>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <th width="200px">配置项</th>
                      <th>内容</th>
                    </tr>
                    <tr>
                      <td>发表评论时自动同步</td>
                      <td><input name="d" id="d" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("d"))%> /></td>
                    </tr>
                    <tr>
                      <td><font color="#0000CC">以分享的方式同步到QQ空间</font></td>
                      <td><input name="e" id="e" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("e"))%> /></td>
                    </tr>
                    <tr>
                      <td><div id="help002" style="display:none"> 若该文章没有同步到微博则评论也不同步 </div>
                        <font color="#009900">以评论的方式同步到腾讯微博</font><a id="help02" href="$help002?width=320" class="betterTip" title="帮助">？</a></td>
                      <td><input name="f" id="f" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("f"))%> /></td>
                    </tr>
                    <tr>
                      <td><div id="help001" style="display:none"> 无论有无勾选，若用户没有绑定QQ，则使用管理员的身份同步。微博暂只能使用管理员的身份同步 </div>
                        <font color="#0000CC">使用管理员的身份同步到空间</font><a id="help01" href="$help001?width=320" class="betterTip" title="帮助">？</a></td>
                      <td><input name="g" id="g" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("g"))%> /></td>
                    </tr>
                    <tr>
                      <td><font color="#009900">同步到微博内容</font></td>
                      <td><p>%a—评论作者；%c——评论摘要</p>
                        <input type="text" name="pl" id="pl" value="<%=qqconnect.tconfig.read("pl")%>"/></td>
                    </tr>
                  </table>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <th width="200px">配置项</th>
                      <th>内容</th>
                    </tr>
                    <tr>
                      <td><font color="#0000CC">允许使用QQ登录</font></td>
                      <td><input name="h" id="h" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("h"))%> /></td>
                    </tr>
                    <tr>
                      <td><div id="help003" style="display:none"> 必须打开注册组件 </div>
                        <font color="#0000CC">允许使用QQ注册帐号</font><a id="help03" href="$help003?width=320" class="betterTip" title="帮助">？</a></td>
                      <td><input name="i" id="i" type="text" class="checkbox" <%=d(qqconnect.tconfig.read("i"))%> /></td>
                    </tr>
                  </table>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab6">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <td width="100%"><p></p>
                        <p>
                          <label>
                            <input type="radio" name="a1" value="0" id="a1_0"<%=e(qqconnect.tconfig.read("a1"),0)%>/>
                            <font color="#0000CC"> 评论显示微博头像（若无，则使用系统自带）</font></label>
                        </p>
                        <p>
                          <label>
                            <input type="radio" name="a1" value="1" id="a1_1"<%=e(qqconnect.tconfig.read("a1"),1)%>/>
                            <font color="#0000CC">评论显示空间头像（若无，则使用系统自带）</font></label>
                        </p>
                        <p>
                          <label>
                            <input type="radio" name="a1" value="2" id="a1_2"<%=e(qqconnect.tconfig.read("a1"),2)%>/>
                            评论使用系统自带头像</label>
                        </p>
                        <p>注：评论内需要有<#article/comment/avatar#>才可以使用本功能！</p>
                        <p></p></td>
                    </tr>
                  </table>
                </div>
                <%End If%>
                <div class="tab-content<%=IIf(BlogUser.Level=1,""," default-tab")%>" style='border:none;padding:0px;margin:0;' id="tab7">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <th width="200px">配置项</th>
                      <th>内容</th>
                    </tr>
                    <tr>
                      <td><div id="help005" style="display:none"> 只对当前登录用户<%=BlogUser.FirstName%>生效</div>
                        <font color="#0000CC">评论同步到QQ空间</font><a id="help05" href="$help005?width=320" class="betterTip" title="帮助">？</a></td>
                      <td><input type="text" class="checkbox" name="synctoqzone" id="synctoqzone" value="<%
					  
					  Dim s
					  s=BlogUser.Meta.GetValue("qqconnect_sync")
					  If s="False" Then Response.Write False Else Response.Write True%>"/></td>
                    </tr>
                  </table>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab4">
                  <table class="tableBorder"  style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' width="100%">
                    <tr>
                      <td width="100%"><p>插件名：QQConnect</p>
                        <p>插件版本:v1.0</p>
                        <p>插件制作者：ZSXSOFT</p>
                        <p>微博：<a href="http://t.qq.com/zhengshixin163" target="_blank">http://t.qq.com/zhengshixin163</a></p>
                        <p>网站：<a href="http://www.zsxsoft.com" target="_blank">http://www.zsxsoft.com</a></p></td>
                    </tr>
                  </table>
                </div>
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
ActiveLeftMenu("anewQQConnect");
</script>
      <%
function d(v)
	d=iif(v="True"," value=""True"" "," value=""False"" ")
end function
function e(s,b)
	e=iif(cint(s)=b," checked=""checked"" ","")
end function
%>
      <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->