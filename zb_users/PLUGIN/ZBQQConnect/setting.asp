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
input[type="text"]{width:100%}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
    <div id="divMain">
      <div id="ShowBlogHint">
        <%Call GetBlogHint()%>
      </div>
      <div class="divHeader">ZBQQConnect</div>
      <div class="SubMenu"><%=ZBQQConnect_SBar(3)%></div>
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
            <form id="form1" name="form1" method="post" action="savesetting.asp">
              <div class="content-box-content" id="qqcbox">
                <div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab0">
                  <p>APP ID:
                    <input name="AppId" type="text" id="ap" value="<%=ZBQQConnect_Config.Read("AppID")%>"/>
                  </p>
                  <p><br />
                    KEY:
                    <input name="KEY" type="text" id="as" value="<%=ZBQQConnect_Config.Read("KEY")%>"/>
                  </p>
                  <p><a href="javascript:void(0)" onclick="showqk()">如何获得？</a></p>
                  <div class="hidden" id="how">
                    <ol style="list-style-type:decimal;">
                      <li>
                        <p> 首先打开<a href="http://connect.qq.com/intro/login/" target="_blank">http://connect.qq.com/intro/login/</a>，点击申请加入 </p>
                      </li>
                      <li>
                        <p> 登录QQ </p>
                      </li>
                      <li>
                        <p> 填写个人信息，注册成为开发者 </p>
                      </li>
                      <li>
                        <p> 打开<a href="http://connect.qq.com/manage/" target="_blank">http://connect.qq.com/manage/</a>，点击右上角的“添加网站/应用” </p>
                      </li>
                      <li>
                        <p> 填写你的网站名称、域名。域名需要认证。如图 </p><p><img src="a.jpg" alt="" width="486" height="365"></p>

                        <blockquote><p> 认证方法： </p>
                        <p> 主题管理--&gt;修改主题--&gt;TEMPLATE--&gt;default.html，在&lt;#TEMPLATE_HEADER#&gt;下插入代码 </p></blockquote>
                      </li>
                      <li>
                        <p> 把得到的APP ID和KEY填写入插件即可<br />
                        </p>
                      </li>
                    </ol>
                  </div>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab1">
                  <p>
                    <input name="a" id="a" type="checkbox" <%=d(ZBQQConnect_Config.Read("a"))%> />
                    <label for="a">发表文章时默认同步到QQ空间</label>
                    <br/>
                    <input name="b" id="b" type="checkbox" <%=d(ZBQQConnect_Config.Read("b"))%> />
                   
                    <label for="b">发表文章时默认同步到腾讯微博</label>
                    <br />
                    <input name="c" id="c" type="checkbox" <%=d(ZBQQConnect_Config.Read("c"))%> />
                    <label for="c">同步时自带文章第一张图片</label>
                  </p>
                  <p>同步到腾讯微博内容（%i——截断字数的摘要；%b——博客名称；%u——文章地址；%t——文章标题；%c——分类）</p>
                  <p>
                    <label for="content"></label>
                    <input type="text" name="content" id="content" value="<%=ZBQQConnect_Config.Read("content")%>"/>
                  </p>
                  <p><br />
                </p>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
                  <input name="d" id="d" type="checkbox" <%=d(ZBQQConnect_Config.Read("d"))%> />
                  <label for="d">发表评论时自动同步</label>
                  <br/>
                  <input name="e" id="e" type="checkbox" <%=d(ZBQQConnect_Config.Read("e"))%> />
                  <label for="e">以分享的方式同步到QQ空间</label>
                  <br />
                  <input name="f" id="f" type="checkbox" <%=d(ZBQQConnect_Config.Read("f"))%> />
                  <label for="f">以评论的方式同步到腾讯微博（若该文章没有同步到微博则评论也不同步）</label>
                  <br />
                  <input name="g" id="g" type="checkbox" <%=d(ZBQQConnect_Config.Read("g"))%> />
                  <label for="g">使用管理员的身份同步（无论有无勾选，若用户没有绑定QQ，则使用管理员的身份同步）</label>
                  <br />
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
                  <input name="h" id="h" type="checkbox" <%=d(ZBQQConnect_Config.Read("h"))%> />
                  <label for="h">允许使用QQ登录</label>
                  <br/>
                  <input name="i" id="i" type="checkbox" <%=d(ZBQQConnect_Config.Read("i"))%> />
                  <label for="i">允许使用QQ注册帐号（不过必须先启用注册组件）</label>
                  
                  <br/>
                </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab5">
                <p>
                  <label>
                    <input type="radio" name="a1" value="0" id="a1_0"<%=e(ZBQQConnect_Config.Read("a1"),0)%>/>
                    评论显示微博头像（若无，则使用Gravatar）</label>
                  <br />
                  <label>
                    <input type="radio" name="a1" value="1" id="a1_1"<%=e(ZBQQConnect_Config.Read("a1"),1)%>/>
                    评论显示空间头像（若无，则使用Gravatar）</label>
                  <br />
                  <label>
                    <input type="radio" name="a1" value="2" id="a1_2"<%=e(ZBQQConnect_Config.Read("a1"),2)%>/>
                    评论显示Gravatar</label>
                  </p>
                <p>
                输入Gravatar地址（用<#EmailMD5#>来代替EMail的MD5,<#ZC_BLOG_HOST#>代替网站域名）
                  <input name="Gravatar" type="text" id="Gravatar" value="<%=ZBQQConnect_Config.Read("Gravatar")%>"/>
                </p>
                <p>调用代码： </p>
                <pre class="brush:html;toolbar:false">
&lt;!--此处&lt;#ZBQQConnect_Head#&gt;为调用地址，必须要有。其他视情况而定--&gt;
&lt;img width=&quot;32&quot; height=&quot;32&quot; alt=&quot;头像&quot; title=&quot;头像&quot; src=&quot;&lt;#ZBQQConnect_Head#&gt;&quot; /&gt;
                </pre>
                </p>
              </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab4"><br/>
                  插件名：ZBQQConnect<br/>
                  插件版本:v1.0<br/>
                  插件制作者：ZSXSOFT<br/>
                  微博：<a href="http://t.qq.com/zhengshixin163" target="_blank">http://t.qq.com/zhengshixin163</a><br/>
                  网站：<a href="http://www.zsxsoft.com" target="_blank">http://www.zsxsoft.com</a> </div>
              </div>
              <input type="submit" value="提交" class="button"/>
            </form>
          </div>
        </div>
        <script language="javascript">
SyntaxHighlighter.all();

function ChangeValue(obj){

	if (obj.value=="True")
	{
	obj.value="False";
	return true;
	}

	if (obj.value=="False")
	{
	obj.value="True";
	return true;
	}
}


    // Content box tabs:
		
		$('.content-box .content-box-content div.tab-content').hide(); // Hide the content divs
		$('ul.content-box-tabs li a.default-tab').addClass('current'); // Add the class "current" to the default tab
		$('.content-box-content div.default-tab').show(); // Show the div with class "default-tab"
		
		$('.content-box ul.content-box-tabs li a').click( // When a tab is clicked...
			function() { 
				$(this).parent().siblings().find("a").removeClass('current'); // Remove "current" class from all tabs
				$(this).addClass('current'); // Add class "current" to clicked tab
				var currentTab = $(this).attr('href'); // Set variable "currentTab" to the value of href of clicked tab
				$(currentTab).siblings().hide(); // Hide all content divs
				$(currentTab).show(); // Show the content div with the id equal to the id of clicked tab
				return false; 
			}
		);




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