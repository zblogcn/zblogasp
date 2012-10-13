<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>

<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AdvancedConfig")=False Then Call ShowError(48)
BlogTitle="AdvancedConfig"
Dim i
If Request.QueryString("act")="save" Then
	For i=1 To BlogConfig.Count
		If Not IsEmpty(Request.Form(BlogConfig.Meta.Names(i))) Then
			BlogConfig.Write BlogConfig.Meta.Names(i),Request.Form(BlogConfig.Meta.Names(i))
		End If
		
	Next
	BlogConfig.Save
	SaveConfig2Option
	Response.Redirect "main.asp"
End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=Response_Plugin_SettingMng_SubMenu%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("topmenu2");</script> 
            <form id="form1" name="form1" method="post" action="?act=save">
            
			<table width="100%"><tr height="40"><td width="20%">配置项</td><td>配置</td><td width="30%">备注</td></tr>
			<%
			
			
			For i=1 To BlogConfig.Count
				Response.Write "<tr><td>"& GetName(BlogConfig.Meta.Names(i)) & "</td><td>"& ExportConfig(BlogConfig.Meta.Names(i),vbsunescape(BlogConfig.Meta.Values(i))) &"</td><td>"& GetValue(BlogConfig.Meta.Names(i)) &"</td></tr>"
			Next
			%>
          </table>
          <input type="submit" value="提交"/>
          </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%

Function ExportConfig(s,m)
	ExportConfig="<input type=""text"" style=""width:95%"" name=""" & s & """ id=""" & s & """ value=""" & TransferHTML(m,"[html-format]") & """"
	If m="True" Or m="False" Then
		ExportConfig=ExportConfig & " class=""checkbox""/>"
	End If
End Function
%>
<script language="javascript" runat="server">
function GetName(s){
	switch(s.toUpperCase()){
	case "ZC_BLOG_HOST":return "域名配置"
	case "ZC_BLOG_TITLE":return "设置博客标题"
	case "ZC_BLOG_SUBTITLE":return "设置博客副标题"
	case "ZC_BLOG_NAME":return "博客名"
	case "ZC_BLOG_SUB_NAME":return "博客描述"
	case "ZC_BLOG_THEME":return "当前主题"
	case "ZC_BLOG_CSS":return "当前样式"
	case "ZC_BLOG_COPYRIGHT":return "底部版权"
	case "ZC_BLOG_MASTER":return "博客创始人"
	case "ZC_BLOG_LANGUAGE":return "博客语言"
	case "ZC_DATABASE_PATH":return "Access数据库路径"
	case "ZC_MSSQL_ENABLE":return "是否使用MSSQL"
	case "ZC_MSSQL_DATABASE":return "MSSQL数据库"
	case "ZC_MSSQL_USERNAME":return "MSSQL用户名"
	case "ZC_MSSQL_PASSWORD":return "MSSQL密码"
	case "ZC_MSSQL_SERVER":return "MSSQL服务器"
	case "ZC_USING_PLUGIN_LIST":return "启用插件列表"
	case "ZC_BLOG_CLSID":return "配置CLSID"
	case "ZC_TIME_ZONE":return "配置使用者时区"
	case "ZC_HOST_TIME_ZONE":return "配置服务器时区"
	case "ZC_UPDATE_INFO_URL":return "后台公告地址"
	case "ZC_MULTI_DOMAIN_SUPPORT":return "多域名支持"
	case "ZC_BLOG_VERSION":return "Z-Blog版本"
	case "ZC_COMMENT_TURNOFF":return "关闭评论"
	case "ZC_COMMENT_VERIFY_ENABLE":return "打开验证码"
	case "ZC_COMMENT_REVERSE_ORDER_EXPORT":return "评论倒序输出"
	case "ZC_COMMNET_MAXFLOOR":return "评论盖楼最大层数"
	case "ZC_VERIFYCODE_STRING":return "验证码可用字符（支持0-9A-Za-z）"
	case "ZC_VERIFYCODE_WIDTH":return "验证码宽度"
	case "ZC_VERIFYCODE_HEIGHT":return "验证码高度"
	case "ZC_DISPLAY_COUNT":return "每页显示数量"
	case "ZC_RSS2_COUNT":return "RSS显示数量"
	case "ZC_SEARCH_COUNT":return "搜索结果显示数量"
	case "ZC_PAGEBAR_COUNT":return "分页条显示数量"  
	case "ZC_MUTUALITY_COUNT":return "相关文章显示数量"
	case "ZC_COMMENTS_DISPLAY_COUNT":return "评论每页显示数量"
	case "ZC_USE_NAVIGATE_ARTICLE":return "文章页是否显示上下篇导航"
	case "ZC_RSS_EXPORT_WHOLE":return "是否输出全文RSS"
	case "ZC_MANAGE_COUNT":return "后台管理数量"
	case "ZC_REBUILD_FILE_COUNT":return "每次重建数目"
	case "ZC_REBUILD_FILE_INTERVAL":return "重建间隔"

	case "ZC_UBB_ENABLE":return "打开UBB"
	/*ZC_UBB_LINK_ENABLE
	ZC_UBB_FONT_ENABLE
	ZC_UBB_CODE_ENABLE
	ZC_UBB_FACE_ENABLE
	ZC_UBB_IMAGE_ENABLE
	ZC_UBB_MEDIA_ENABLE
	ZC_UBB_FLASH_ENABLE
	ZC_UBB_TYPESET_ENABLE
	ZC_UBB_AUTOLINK_ENABLE
	ZC_UBB_AUTOKEY_ENABLE*/
	case "ZC_EMOTICONS_FILENAME":return "表情文件名（弃用）"
	case "ZC_EMOTICONS_FILETYPE":return "表情后缀名"
	case "ZC_EMOTICONS_FILESIZE":return "表情大小（弃用）"
	case "ZC_UPLOAD_FILETYPE":return "允许上传后缀名"
	case "ZC_UPLOAD_FILESIZE":return "最大文件大小"
	case "ZC_UPLOAD_DIRBYMONTH":return "上传附件按月存档"
	case "ZC_UPLOAD_DIRECTORY":return "附件保存目录"
	case "ZC_USERNAME_MIN":return "最小用户名长度"
	case "ZC_USERNAME_MAX":return "最大用户名长度"
	case "ZC_PASSWORD_MIN":return "最小密码长度"
	case "ZC_PASSWORD_MAX":return "最大密码长度"
	case "ZC_EMAIL_MAX":return "最大邮件长度"
	case "ZC_HOMEPAGE_MAX":return "最大网址长度"
	case "ZC_CONTENT_MAX":return "最大评论长度"
	case "ZC_STATIC_TYPE":return "静态文件后缀名"
	case "ZC_STATIC_DIRECTORY":return "静态文件保存路径"
	case "ZC_TEMPLATE_DIRECTORY":return "主题模板保存路径"
	case "ZC_STATIC_MODE":return "当前模式（动态或伪静态）"
	case "ZC_ARTICLE_REGEX":return "文章网址格式"
	case "ZC_PAGE_REGEX":return "页面网址格式"
	case "ZC_CATEGORY_REGEX":return "分类网址格式"
	case "ZC_USER_REGEX":return "用户网址格式"
	case "ZC_TAGS_REGEX":return "标签网址格式"
	case "ZC_DATE_REGEX":return "日期网址格式"
	case "ZC_DEFAULT_REGEX":return "首页网址格式"
	case "ZC_DISPLAY_COUNT_WAP":return "文章列表单页显示文章数量"
	case "ZC_COMMENT_COUNT_WAP":return "单页显示评论数量"
	case "ZC_PAGEBAR_COUNT_WAP":return "文章列表评论条显示条数"
	case "ZC_SINGLE_SIZE_WAP":return "开启分页查看文章时单页字数"
	case "ZC_SINGLE_PAGEBAR_COUNT_WAP":return "WAP文章分页数（未启用）"
	case "ZC_FILENAME_WAP":return "WAP文件地址"
	case "ZC_WAPCOMMENT_ENABLE":return "打开WAP评论"
	case "ZC_DISPLAY_MODE_ALL_WAP":return "WAP显示全文"
	case "ZC_DISPLAY_CATE_ALL_WAP":return "WAP显示导航"
	case "ZC_DISPLAY_PAGEBAR_ALL_WAP":return "WAP分页条"
	case "ZC_WAP_MUTUALITY_LIMIT":return "WAP相关文章数量"
	case "ZC_SYNTAXHIGHLIGHTER_ENABLE":return "打开代码高亮"
	case "ZC_CODEMIRROR_ENABLE":return "打开源码编辑高亮"
	case "ZC_ARTICLE_EXCERPT_MAX":return  "自动截取摘要字数"
	case "ZC_UNCATEGORIZED_NAME":return "自定义未分类分类名"
	case "ZC_UNCATEGORIZED_ALIAS":return "自定义未分类分类别名"
	case "ZC_UNCATEGORIZED_COUNT":return "未分类分类文章计数"
	case "ZC_POST_STATIC_MODE":return "文章模式（动态、静态或伪静态）"

	default:return s
	}
}
function GetValue(s){
	switch(s.toUpperCase()){
	case "ZC_BLOG_HOST":return "2.0可以自动修改，已经废弃"
	case "ZC_BLOG_NAME":return " 在2.0中等同于ZC_BLOG_TITLE，没有任何用处"
	case "ZC_BLOG_SUB_NAME":return " 在2.0中等同于ZC_BLOG_SUBTITLE，没有任何用处"
	case "ZC_BLOG_THEME":
	case "ZC_BLOG_CSS":
	case "ZC_USING_PLUGIN_LIST":
	case "ZC_BLOG_CLSID":
	case "ZC_UPDATE_INFO_URL":
	case "ZC_BLOG_VERSION":
	case "ZC_UNCATEGORIZED_COUNT":	return "请不要任意修改"
	case "ZC_POST_STATIC_MODE":return "三种模式：ACTIVE   STATIC   REWRITE"
	case "ZC_MULTI_DOMAIN_SUPPORT":return "打开则以相对路径显示博客，不再绑定域名"
	case "ZC_ARTICLE_REGEX":return "网址格式，下面都是"
	}
}
</script>