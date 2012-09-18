<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'---------------------------------网站基本设置-----------------------------------
Call BlogConfig.Write("ZC_BLOG_HOST","http://localhost/")
Call BlogConfig.Write("ZC_BLOG_TITLE","My Blog")
Call BlogConfig.Write("ZC_BLOG_SUBTITLE","Hello, world!")
Call BlogConfig.Write("ZC_BLOG_NAME","My Blog")
Call BlogConfig.Write("ZC_BLOG_SUB_NAME","Hello, world!")
Call BlogConfig.Write("ZC_BLOG_THEME","default")
Call BlogConfig.Write("ZC_BLOG_CSS","default")
Call BlogConfig.Write("ZC_BLOG_COPYRIGHT","Copyright Your WebSite. Some Rights Reserved.")
Call BlogConfig.Write("ZC_BLOG_MASTER","zblogger")
Call BlogConfig.Write("ZC_BLOG_LANGUAGE","zh-CN")





'----------------------------数据库配置---------------------------------------
Call BlogConfig.Write("ZC_DATABASE_PATH","zb_users\data\#%20768d53283c63b13403f0.mdb")
Call BlogConfig.Write("ZC_MSSQL_ENABLE",False)
Call BlogConfig.Write("ZC_MSSQL_DATABASE","zb")
Call BlogConfig.Write("ZC_MSSQL_USERNAME","sa")
Call BlogConfig.Write("ZC_MSSQL_PASSWORD","")
Call BlogConfig.Write("ZC_MSSQL_SERVER","(local)\SQLEXPRESS")





'---------------------------------插件----------------------------------------
Call BlogConfig.Write("ZC_USING_PLUGIN_LIST","")








'-------------------------------全局配置-----------------------------------
Call BlogConfig.Write("ZC_BLOG_CLSID","BB1C5669-6E37-460C-F415-D287D7BBB59E")
Call BlogConfig.Write("ZC_TIME_ZONE","+0800")
Call BlogConfig.Write("ZC_HOST_TIME_ZONE","+0800")
Call BlogConfig.Write("ZC_UPDATE_INFO_URL","http://update.rainbowsoft.org/info/")
Call BlogConfig.Write("ZC_MULTI_DOMAIN_SUPPORT",False)




'留言评论
Call BlogConfig.Write("ZC_COMMENT_TURNOFF",False)
Call BlogConfig.Write("ZC_COMMENT_VERIFY_ENABLE",False)
Call BlogConfig.Write("ZC_COMMENT_NOFOLLOW_ENABLE",True)
Call BlogConfig.Write("ZC_COMMENT_REVERSE_ORDER_EXPORT",False)
Call BlogConfig.Write("ZC_COMMNET_MAXFLOOR",4)


'验证码
Call BlogConfig.Write("ZC_VERIFYCODE_STRING","0123456789")
Call BlogConfig.Write("ZC_VERIFYCODE_WIDTH",60)
Call BlogConfig.Write("ZC_VERIFYCODE_HEIGHT",20)


Call BlogConfig.Write("ZC_DISPLAY_COUNT",10)
Call BlogConfig.Write("ZC_RSS2_COUNT",10)
Call BlogConfig.Write("ZC_SEARCH_COUNT",25)
Call BlogConfig.Write("ZC_PAGEBAR_COUNT",15)
Call BlogConfig.Write("ZC_MUTUALITY_COUNT",10)
Call BlogConfig.Write("ZC_COMMENTS_DISPLAY_COUNT",10)





Call BlogConfig.Write("ZC_IMAGE_WIDTH",520)

Call BlogConfig.Write("ZC_USE_NAVIGATE_ARTICLE",True)

Call BlogConfig.Write("ZC_RSS_EXPORT_WHOLE",False)




'后台管理
Call BlogConfig.Write("ZC_MANAGE_COUNT",50)
Call BlogConfig.Write("ZC_REBUILD_FILE_COUNT",50)
Call BlogConfig.Write("ZC_REBUILD_FILE_INTERVAL",1)










'UBB转换
Call BlogConfig.Write("ZC_UBB_ENABLE",False)
Call BlogConfig.Write("ZC_UBB_LINK_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_FONT_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_CODE_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_FACE_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_IMAGE_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_MEDIA_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_FLASH_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_TYPESET_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_AUTOLINK_ENABLE",True)
Call BlogConfig.Write("ZC_UBB_AUTOKEY_ENABLE",False)




'表情相关
Call BlogConfig.Write("ZC_EMOTICONS_FILENAME","neutral|grin|happy|slim|smile|tongue|wink|surprised|confuse|cool|cry|evilgrin|fat|mad|red|roll|unhappy|waii|yell")
Call BlogConfig.Write("ZC_EMOTICONS_FILETYPE","png")
Call BlogConfig.Write("ZC_EMOTICONS_FILESIZE",16)




'上传相关
Call BlogConfig.Write("ZC_UPLOAD_FILETYPE","jpg|gif|png|jpeg|bmp|psd|wmf|ico|rpm|deb|tar|gz|sit|7z|bz2|zip|rar|xml|xsl|svg|svgz|doc|xls|wps|chm|txt|pdf|mp3|avi|mpg|rm|ra|rmvb|mov|wmv|wma|swf|fla|torrent|zpi|zti")
Call BlogConfig.Write("ZC_UPLOAD_FILESIZE",10485760)
Call BlogConfig.Write("ZC_UPLOAD_DIRBYMONTH",True)
Call BlogConfig.Write("ZC_UPLOAD_DIRECTORY","zb_users\upload")



'当前 Z-Blog 版本
Call BlogConfig.Write("ZC_BLOG_VERSION","2.0 Beta Build 120819")



'用户名,密码,评论长度等限制
Call BlogConfig.Write("ZC_USERNAME_MIN",4)
Call BlogConfig.Write("ZC_USERNAME_MAX",14)
Call BlogConfig.Write("ZC_PASSWORD_MIN",8)
Call BlogConfig.Write("ZC_PASSWORD_MAX",14)
Call BlogConfig.Write("ZC_EMAIL_MAX",30)
Call BlogConfig.Write("ZC_HOMEPAGE_MAX",100)
Call BlogConfig.Write("ZC_CONTENT_MAX",1000)










'---------------------------------静态化配置-----------------------------------


'{asp html shtml}
Call BlogConfig.Write("ZC_STATIC_TYPE","html")

Call BlogConfig.Write("ZC_STATIC_DIRECTORY","post")

Call BlogConfig.Write("ZC_TEMPLATE_DIRECTORY","template")



'ACTIVE MIX REWRITE
Call BlogConfig.Write("ZC_STATIC_MODE","ACTIVE")

Call BlogConfig.Write("ZC_ARTICLE_REGEX","{%host%}/{%post%}/{%alias%}.html")
Call BlogConfig.Write("ZC_PAGE_REGEX","{%host%}/{%alias%}.html")
Call BlogConfig.Write("ZC_CATEGORY_REGEX","{%host%}/catalog.asp?cate={%id%}")
Call BlogConfig.Write("ZC_USER_REGEX","{%host%}/catalog.asp?user={%id%}")
Call BlogConfig.Write("ZC_TAGS_REGEX","{%host%}/catalog.asp?tags={%alias%}")
Call BlogConfig.Write("ZC_DATE_REGEX","{%host%}/catalog.asp?date={%date%}")
Call BlogConfig.Write("ZC_DEFAULT_REGEX","{%host%}/catalog.asp")





'--------------------------WAP----------------------------------------
Call BlogConfig.Write("ZC_DISPLAY_COUNT_WAP",5)
Call BlogConfig.Write("ZC_COMMENT_COUNT_WAP",5)
Call BlogConfig.Write("ZC_PAGEBAR_COUNT_WAP",5)
Call BlogConfig.Write("ZC_SINGLE_SIZE_WAP",1000)
Call BlogConfig.Write("ZC_SINGLE_PAGEBAR_COUNT_WAP",5)

Call BlogConfig.Write("ZC_FILENAME_WAP","wap.asp")
Call BlogConfig.Write("ZC_WAPCOMMENT_ENABLE",True)
'全文
Call BlogConfig.Write("ZC_DISPLAY_MODE_ALL_WAP",True)
'显示分类导航
Call BlogConfig.Write("ZC_DISPLAY_CATE_ALL_WAP",True)
'分页条
Call BlogConfig.Write("ZC_DISPLAY_PAGEBAR_ALL_WAP",True)
'相关文章
Call BlogConfig.Write("ZC_WAP_MUTUALITY",True)
'数量
Call BlogConfig.Write("ZC_WAP_MUTUALITY_LIMIT",5)

'Response.Write BlogConfig.Count
'Response.Write BlogConfig.Count
BlogConfig.Save

%>
