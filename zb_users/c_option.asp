<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_option.asp
'// 开始时间:    2004.07.25
'// 最后修改:    
'// 备    注:    设置模块
'///////////////////////////////////////////////////////////////////////////////




'--------------------------------------------------------------------
Const adOpenForwardOnly=0
Const adOpenKeyset=1
Const adOpenDynamic=2
Const adOpenStatic=3

Const adLockReadOnly=1
Const adLockPessimistic=2
Const adLockOptimistic=3
Const adLockBatchOptimistic=4

Const ForReading=1
Const ForWriting=2
Const ForAppending=8

Const adTypeBinary=1
Const adTypeText=2

Const adModeRead=1
Const adModeReadWrite=3

Const adSaveCreateNotExist=1
Const adSaveCreateOverWrite=2
'--------------------------------------------------------------------




'--------------------------------------------------------------------
Const ZC_BLOG_CLSID="BB1C5669-6E37-460C-F415-D287D7BBB59E"



Const ZC_BLOG_WEBEDIT="ueditor"

Const ZC_TIME_ZONE="+0800"
Const ZC_HOST_TIME_ZONE="+0800"

Const ZC_MSG_COUNT=10
Const ZC_ARCHIVE_COUNT=0
Const ZC_PREVIOUS_COUNT=15
Const ZC_DISPLAY_COUNT=10
Const ZC_MUTUALITY_COUNT=10

Const ZC_MANAGE_COUNT=50
Const ZC_RSS2_COUNT=10
Const ZC_SEARCH_COUNT=25
Const ZC_PAGEBAR_COUNT=14

Const ZC_TAGS_DISPLAY_COUNT=50

Const ZC_COMMENTS_DISPLAY_COUNT=10

Const ZC_IMAGE_WIDTH=520
Const ZC_REBUILD_FILE_COUNT=50
Const ZC_REBUILD_FILE_INTERVAL=0

Const ZC_AUTO_NEWLINE=True
Const ZC_JAPAN_TO_HTML=False
Const ZC_USE_NAVIGATE_ARTICLE=True


Const ZC_COMMENT_TURNOFF=False
Const ZC_TRACKBACK_TURNOFF=True

Const ZC_VERIFYCODE_STRING="0123456789"
Const ZC_VERIFYCODE_WIDTH=60
Const ZC_VERIFYCODE_HEIGHT=20
Const ZC_COMMENT_VERIFY_ENABLE=False
Const ZC_COMMENT_NOFOLLOW_ENABLE=True
Const ZC_RSS_EXPORT_WHOLE=False

Const ZC_UBB_LINK_ENABLE=True
Const ZC_UBB_FONT_ENABLE=True
Const ZC_UBB_CODE_ENABLE=True
Const ZC_UBB_FACE_ENABLE=True
Const ZC_UBB_IMAGE_ENABLE=True
Const ZC_UBB_MEDIA_ENABLE=True
Const ZC_UBB_FLASH_ENABLE=True
Const ZC_UBB_TYPESET_ENABLE=True
Const ZC_UBB_AUTOLINK_ENABLE=True
Const ZC_UBB_AUTOKEY_ENABLE=False


Const ZC_EMOTICONS_FILENAME="neutral|grin|happy|slim|smile|tongue|wink|surprised|confuse|cool|cry|evilgrin|fat|mad|red|roll|unhappy|waii|yell"
Const ZC_EMOTICONS_FILETYPE="png"
Const ZC_EMOTICONS_FILESIZE=16


Const ZC_UPLOAD_FILETYPE="jpg|gif|png|jpeg|bmp|psd|wmf|ico|rpm|deb|tar|gz|sit|7z|bz2|zip|rar|xml|xsl|svg|svgz|doc|xls|wps|chm|txt|pdf|mp3|avi|mpg|rm|ra|rmvb|mov|wmv|wma|swf|fla|torrent|zpi|zti"
Const ZC_UPLOAD_FILESIZE=10485760
Const ZC_UPLOAD_DIRBYMONTH=True


Const ZC_DISPLAY_MODE_ALL=1
Const ZC_DISPLAY_MODE_INTRO=2
Const ZC_DISPLAY_MODE_HIDE=3
Const ZC_DISPLAY_MODE_LIST=4
Const ZC_DISPLAY_MODE_ONTOP=5
Const ZC_DISPLAY_MODE_SEARCH=6

Const ZC_USERNAME_MAX=20
Const ZC_PASSWORD_MAX=32
Const ZC_EMAIL_MAX=30
Const ZC_HOMEPAGE_MAX=100
Const ZC_CONTENT_MAX=1000
Const ZC_TB_EXCERPT_MAX=250
Const ZC_RECENT_COMMENT_WORD_MAX=16

Const ZC_COMMENT_REVERSE_ORDER_EXPORT=False
Const ZC_GUEST_REVERT_COMMENT_ENABLE=True


Const ZC_CUSTOM_DIRECTORY_ENABLE=False
'{%post%},{%category%},{%user%},{%year%},{%month%},{%day%},{%id%},{%alias%}之间的组合,可以用/分隔
Const ZC_CUSTOM_DIRECTORY_REGEX="{%post%}"
Const ZC_CUSTOM_DIRECTORY_ANONYMOUS=False

Const ZC_GUESTBOOK_CONTENT="欢迎给我留言。"
Const ZC_GUESTBOOK_ID=0

Const ZC_UPDATE_INFO_URL="http://update.rainbowsoft.org/info/"


Const ZC_USING_PLUGIN_LIST="PluginSapper|FileManage"



'--------------------------------------------------------------------

Const ZC_IE_DISPLAY_WAP=False
Const ZC_DISPLAY_COUNT_WAP=2
Const ZC_COMMENT_COUNT_WAP=3
Const ZC_PAGEBAR_COUNT_WAP=5
Const ZC_SINGLE_SIZE_WAP=1000
Const ZC_SINGLE_PAGEBAR_COUNT_WAP=5
Const ZC_COMMENT_PAGEBAR_COUNT_WAP=5
Const ZC_FILENAME_WAP="wap.asp"
Const ZC_WAPCOMMENT_ENABLE=True

'--------------------------------------------------------------------


'{asp html shtml}
Const ZC_STATIC_TYPE="html"

Const ZC_STATIC_DIRECTORY="post"
Const ZC_TEMPLATE_DIRECTORY="template"
Const ZC_UPLOAD_DIRECTORY="upload"

Const ZC_BLOG_LANGUAGE="zh-CN"
'--------------------------------------------------------------------

Const ZC_MAXFLOOR=4

%>
<!-- #include file="c_custom.asp" -->