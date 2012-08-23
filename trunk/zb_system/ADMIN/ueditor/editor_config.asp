<%@ CODEPAGE=65001 %>
<!-- #include file="../../../zb_users/c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_users/plugin/p_config.asp" -->
<%
Response.ContentType="application/x-javascript"
%>
<%
Call ActivePlugin()
For Each sAction_Plugin_UEditor_Config_Begin in Action_Plugin_UEditor_Config_Begin
	If Not IsEmpty(sAction_Plugin_UEditor_Config_Begin) Then Call Execute(sAction_Plugin_UEditor_Config_Begin)
Next


	Dim strUPLOADDIR

	strUPLOADDIR = Replace(ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())),"\","/")

	Dim Path
	Path=BlogHost & ""& strUPLOADDIR &"/"
	dim strJSContent
	strJSContent="(function(){var URL;URL = '"&BlogHost&"zb_system/admin/ueditor/'; window.UEDITOR_CONFIG = {UEDITOR_HOME_URL : URL,imageUrl:URL+""asp/picUp.asp"",imagePath:"""&Path&""" ,imageFieldName:""edtFileLoad"" ,fileUrl:URL+""asp/fileUp.asp"",filePath:"""&Path&""" ,fileFieldName:""edtFileLoad"",catchRemoteImageEnable:false,imageManagerUrl:URL +""asp/imageManager.asp"" ,imageManagerPath:"""&BlogHost&""",wordImageUrl:URL+""asp/picUp.asp"",wordImagePath:"""&Path&""",wordImageFieldName:""edtFileLoad"",getMovieUrl:URL+""asp/getMovie.asp"",toolbars:[ ['fullscreen', 'source', '|', 'undo', 'redo', '|', 'bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript','|',  'forecolor', 'backcolor', 'insertorderedlist', 'insertunorderedlist','|', 'indent', '|', 'justifyleft', 'justifycenter', 'justifyright', 'justifyjustify',  '|',  'removeformat','autotypeset', 'searchreplace'],[ 'fontfamily', 'fontsize', '|', 'emotion','link','insertimage', 'insertvideo', 'attachment','spechars','|', 'map', 'gmap', '|', 'highlightcode','blockquote', 'pasteplain','wordimage','|','inserttable', 'deletetable', '|','insertintro','|','preview']],labelMap:{ 'anchor':'锚点', 'undo':'撤销', 'redo':'重做', 'bold':'加粗', 'indent':'首行缩进', 'italic':'斜体', 'underline':'下划线', 'strikethrough':'删除线', 'subscript':'下标', 'superscript':'上标', 'source':'源代码', 'blockquote':'引用', 'pasteplain':'纯文本粘贴模式', 'selectall':'全选', 'print':'打印', 'preview':'预览', 'horizontal':'分隔线','removeformat':'清除格式', 'unlink':'取消链接', 'insertrow':'前插入行', 'insertcol':'前插入列', 'mergeright':'右合并单元格', 'mergedown':'下合并单元格', 'deleterow':'删除行','wordimage':'从Word复制图片','autotypeset': '自动排版', 'deletecol':'删除列', 'splittorows':'拆分成行', 'splittocols':'拆分成列', 'splittocells':'完全拆分单元格', 'mergecells':'合并多个单元格', 'deletetable':'删除表格', 'insertparagraphbeforetable':'表格前插行', 'cleardoc':'清空文档', 'fontfamily':'字体', 'fontsize':'字号', 'paragraph':'段落格式', 'insertimage':'图片', 'inserttable':'表格', 'link':'超链接', 'emotion':'表情', 'spechars':'特殊字符', 'searchreplace':'查询替换', 'map':'Baidu地图', 'gmap':'Google地图', 'insertvideo':'视频', 'justifyleft':'居左对齐', 'justifyright':'居右对齐', 'justifycenter':'居中对齐', 'justifyjustify':'两端对齐', 'forecolor':'字体颜色', 'backcolor':'背景色', 'insertorderedlist':'有序列表', 'insertunorderedlist':'无序列表', 'fullscreen':'全屏','RowSpacingTop':'段前距', 'RowSpacingBottom':'段后距','highlightcode':'插入代码','imagenone':'默认', 'imageleft':'左浮动', 'imageright':'右浮动','attachment':'附件', 'imagecenter':'居中','insertintro':'摘要分割'},isShow : true,initialContent:'<p></p>' ,iframeCssUrl: URL+'/themes/default/iframe.css',textarea:'editorValue' ,focus:false ,minFrameHeight:350,autoClearEmptyNode : false ,fullscreen : false ,readonly : false ,zIndex : 900 ,imagePopup:true,initialStyle:'body{font-size:14px;line-height: 1.5}',emotionLocalization:false ,enterTag:'p' ,pasteplain:false,insertorderedlist : [['1,2,3...','decimal'],['a,b,c...','lower-alpha'],['i,ii,iii...','lower-roman'],['A,B,C','upper-alpha'],['I,II,III...','upper-roman']],insertunorderedlist : [['○','circle'],['●','disc'],['■','square']],'fontfamily':[['宋体',['宋体', 'SimSun']],['楷体',['楷体', '楷体_GB2312', 'SimKai']],['黑体',['黑体', 'SimHei']],['隶书',['隶书', 'SimLi']],		['微软雅黑',['微软雅黑','MSYaHei']],['andale mono',['andale mono']],['arial',['arial', 'helvetica', 'sans-serif']],['system',['system']],['comic sans ms',['comic sans ms']],['impact',['impact', 'chicago']],['times new roman',['times new roman']]],fontsize:[5 , 10, 11, 12, 14, 16, 18, 20, 24, 36, 48],wordCount:false,maximumWords:1000000000,wordCountMsg:'当前已输入 {#count} 个字符 ',highlightJsUrl:URL+""third-party/SyntaxHighlighter/shCore.js"" ,highlightCssUrl:URL+""third-party/SyntaxHighlighter/shCoreDefault.css"",tabSize:4,tabNode:'&nbsp;',elementPathEnabled : false,removeFormatTags:'b,big,code,del,dfn,em,font,i,ins,kbd,q,samp,small,span,strike,strong,sub,sup,tt,u,var',removeFormatAttributes:'class,style,lang,width,height,align,hspace,valign',maxUndoCount:20,maxInputCount:20,autoHeightEnabled:true,autoFloatEnabled:true,indentValue:'2em',sourceEditor:""codemirror"",codeMirrorJsUrl:URL+""third-party/codemirror2.15/codemirror.js"",codeMirrorCssUrl:URL+""third-party/codemirror2.15/codemirror.css"" };})();"


Call Filter_Plugin_UEditor_Config(strJSContent)

For Each sAction_Plugin_UEditor_Config_End in Action_Plugin_UEditor_Config_End
	If Not IsEmpty(sAction_Plugin_UEditor_Config_End) Then Call Execute(sAction_Plugin_UEditor_Config_End)
Next

	response.write strJSContent

%>