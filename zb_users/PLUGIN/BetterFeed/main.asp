<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
BlogTitle="RSS Feed 优化选项"
'读取配置
Call BetterFeed_Config

If (Not IsEmpty(Request.QueryString("s"))) Then Call ExportRSS

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
	<style>
		small{
			font-size:12px;
		}
		h3{
			font-size:16px;
			background:#E4F2FD;
			height:20px;
			margin:10px 5px;
			padding:3px 5px;
		}
		h3.s{
			background:#FFF6EF;
			border:1px solid #FFA65F;
		}
		p{
			line-height:150%;
			margin:5px 10px;
		}
		hr{
			visibility:visible;
			margin:10px;
			border: 1px solid #E4F2FD;
		}
	</style>
	<title><%=BlogTitle%></title>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader"><%=BlogTitle%></div>
<div id="divMain2">

<p><b>优化导航</b> ：<a href="#normal">一般选项</a> · <a href="#comment">评论选项</a> · <a href="#related">相关文章</a> · <a href="#other">其它内容</a> · <a href="#about">插件说明</a></p>
<p>本插件加在原Feed后的内容，将按照如下排列顺序显示：版权申明－继续阅读－添加评论等－相关文章－本文评论－其它内容。</p>
<form id="edit" name="edit" method="post" action="savesetting.asp">

	<h3 id="normal">一般选项</h3>
	<p><strong>版权申明</strong><br/>
	<textarea id="copyright_message" name="copyright_message" cols="60" rows="5" class="code" style="width: 99%;"><%=BetterFeed_Copyright_message%></textarea><br/>
	<small>放空表示不显示版权声明。支持 HTML，你可以使用 %bloglink%、%blogtitle%、%permalink% 、%posttitle%来表示博客地址、博客标题，文章地址及文章标题。推荐使用<a href="http://cn.creativecommons.org/licenses/meet-the-licenses/">CC知识共享许可协议</a>来申明版权许可。</small>
	</p>

	<p>
	<input type="checkbox" value="True" id="addreadmoreinfeed" name="addreadmoreinfeed" <%
	If BetterFeed_Addreadmoreinfeed then
		Response.Write " checked=""checked"">"
	else
		Response.Write ">"
	End if %>
	 <strong>是否在 Feed 中显示继续阅读链接？</strong>
	</p>
	
	<p><strong>继续阅读链接的样式</strong><br/>
	<textarea id="readmore_message" name="readmore_message" cols="60" rows="1" class="code" style="width: 99%;"><%=BetterFeed_Readmore_message%></textarea><br/>
	<small>你可以使用 %permalink% 、%posttitle%分别表示文章地址及文章标题。</small>
	</p>
	<hr/>	
	
	<h3 id="comment">评论选项</h3>

	<p>
	<input type="checkbox" value="True" id="addcommentinfeed" name="addcommentinfeed" <%
	If BetterFeed_Addcommentinfeed then
		Response.Write " checked=""checked""/>"
	else
		Response.Write "/>"
	End if %> <strong>是否在 Feed 中显示添加评论，文章分类，TAGS？</strong>
	</p>
	
	<p><strong>显示的样式</strong><br/>
	<textarea id="addcomment_message" name="comment_message" cols="60" rows="5" class="code" style="width: 99%;"><%=BetterFeed_Comment_message%></textarea><br/>
	<small>支持 HTML。你可以使用 %permalink%，%posttitle%，%commentcount%，%category%，%tags% 分别表示文章地址，文章标题，评论数量，分类，Tags 。</small>
	</p>
		
	<p>
	<input type="checkbox" value="True" id="commentinfeed" name="commentinfeed" <%
	If BetterFeed_Commentinfeed then
		Response.Write " checked=""checked""/>"
	else
		Response.Write "/>"
	End if %> <strong>是否在 Feed 中显示评论？</strong>
	</p>	
	
	<p><input type="text" value="<%=BetterFeed_Commentinfeed_limit%>" size="2" id="commentinfeed_limit" name="commentinfeed_limit"/> <strong>显示最新评论的条目数量</strong><br/></p>
		
	<p><strong>在评论前显示</strong><br/><textarea id="commentinfeed_before" name="commentinfeed_before" cols="60" rows="1" style="width: 99%;"><%=BetterFeed_Commentinfeed_before%></textarea><br/><small>支持 HTML。</small></p>
	
	<p><strong>每条评论的模板</strong><br/><textarea id="commentinfeed_layout" name="commentinfeed_layout" cols="60" rows="5" class="code" style="width: 99%;"><%=BetterFeed_Commentinfeed_layout%></textarea><br/><small>支持 HTML。你可以使用 %permalink%，%authorlink%，"%revlink%"，%date%，%time%，%commentid%，%comment% 分别表示文章地址，作者链接，父评论作者，日期，时间，评论编号，评论内容。</small></p>
	
	<p><strong>在评论后显示</strong><br/><textarea id="commentinfeed_after" name="commentinfeed_after" cols="60" rows="1" style="width: 99%;"><%=BetterFeed_Commentinfeed_after%></textarea><br/><small>支持 HTML。评论内容前后的html标签应该对应闭合。</small></p>
	<hr/>	
	
	<h3 id="related">相关文章选项</h3>

	<p><input type="checkbox" value="True" id="relatedpostinfeed" name="relatedpostinfeed" <%
	If BetterFeed_Relatedpostinfeed then
		Response.Write " checked=""checked""/>"
	else
		Response.Write "/>"
	End if %> <strong>是否在 Feed 中包含相关文章？</strong></p>
	
	<p><input type="text" value="<%=BetterFeed_Relatedpostinfeed_limit%>" size="2" id="relatedpostinfeed_limit" name="relatedpostinfeed_limit"/> <strong>显示相关文章的条目数量</strong><br/></p>
	
	<p><strong>在相关文章条目之前</strong><br/><textarea id="relatedpostinfeed_before" name="relatedpostinfeed_before" cols="60" rows="1" style="width: 99%;"><%=BetterFeed_Relatedpostinfeed_before%></textarea><br/>
	</p>

	<p><strong>每条相关文章的模板</strong><br/><textarea id="relatedpostinfeed_layout" name="relatedpostinfeed_layout" cols="60" rows="5" class="code" style="width: 99%;"><%=BetterFeed_Relatedpostinfeed_layout%></textarea><br/><small>支持 HTML。你可以使用 %article_id%，%article_url%，%article_title%，%article_time% 分别表示文章编号，链接，标题，时间。</small></p>
	
	<p><strong>在相关文章的条目之后</strong><br/><textarea id="relatedpostinfeed_after" name="relatedpostinfeed_after" cols="60" rows="1" style="width: 99%;"><%=BetterFeed_Relatedpostinfeed_after%></textarea><br/><small>支持 HTML。前后的html标签应该对应闭合。</small></p>
	
	<p><strong>如果没有相关文章则显示</strong><br/><textarea id="relatedpostinfeed_sub" name="relatedpostinfeed_sub" cols="60" rows="1" style="width: 99%;"><%=BetterFeed_Relatedpostinfeed_sub%></textarea><br/><small>放空表示不启用这项功能。</small></p>
		
	<hr/>	
	
	<h3 id="other">其它内容</h3>
	
	<p><strong>其它内容添加</strong><br/>
	<textarea id="otherinfeed" name="otherinfeed" cols="60" rows="5" style="width: 99%;"><%=BetterFeed_Otherinfeed%></textarea><br/><small>放空表示不启用这项功能。可以加入订阅到阅读器的代码，Feed广告等等。在前面一些选项中也可以添加此类内容，不过整洁起见还是推荐加在这。
	</small></p>
	
<hr/>	
<p><input type="submit" class="button" id="btnPost" value="提交"/></p>	

</form>
<hr/>	
	<h3 id="about">插件说明</h3>
	<p>
	<small>所有可用标签均可按字面意思理解，需要特别说明的是：</small>
	<ol>
	<li>%bloglink%，%permalink%，%posttitle%是全局标签，所有选项中均可使用。</li>
	<li>%authorlink%标签显示的是留言者名称及其个人链接，但其未留链接时只显示名称。</li>
	<li>添加HTML代码需谨慎，确保代码标签闭合Feed页面才能正常显示。</li>
	</ol>
	<small>  *</small>
	</p>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

