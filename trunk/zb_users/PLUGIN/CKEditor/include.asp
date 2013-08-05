<%
Response_Plugin_Html_Js_Add_CodeHighLight_JS="document.writeln(""<script src='"&BlogHost&"zb_users/plugin/kindeditor/kindeditor/plugins/code/prettify.js' type='text/javascript'></script><link rel='stylesheet' type='text/css' href='"&BlogHost&"zb_users/plugin/kindeditor/kindeditor/plugins/code/prettify.css'/>"");"
Response_Plugin_Html_Js_Add_CodeHighLight_Action="prettyPrint();"

'注册插件
Call RegisterPlugin("CKEditor","ActivePlugin_CKEditor")
'挂口部分
Function ActivePlugin_CKEditor()
	Call Add_Action_Plugin("Action_Plugin_Edit_Form","CKEditor()")
End Function

Sub KindEditor()
	Response_Plugin_Edit_Article_Header="<script src="""&BlogHost & "zb_users/PLUGIN/ckeditor/ckeditor/ckeditor.js""></script>"
	Response_Plugin_Edit_Article_EditorInit="function editor_init(){editor_api.editor.content.get=function(){return this.obj.getData()};editor_api.editor.content.put=function(str){return this.obj.insertHtml(str)};editor_api.editor.content.focus=function(str){return this.obj.focus()};editor_api.editor.intro.get=function(){return this.obj.getData()};editor_api.editor.intro.put=function(str){return this.obj.insertHtml(str)};editor_api.editor.intro.focus=function(str){return this.obj.focus()};$(document).ready(function(){CKEDITOR.replace('editor_txt',{toolbar:[{name:'document',groups:['mode','document','doctools'],items:['Source','-','Preview','Print','-','Templates']},{name:'clipboard',groups:['clipboard','undo'],items:['Cut','Copy','Paste','PasteText','PasteFromWord','-','Undo','Redo']},{name:'editing',groups:['find','selection','spellchecker'],items:['Find','Replace','-','SelectAll']},{name:'paragraph',groups:['list','indent','blocks','align','bidi'],items:['NumberedList','BulletedList','-','Outdent','Indent','-','Blockquote','CreateDiv','-','JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock','-','BidiLtr','BidiRtl']},{name:'links',items:['Link','Unlink','Anchor']},{name:'insert',items:['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar','PageBreak','Iframe']},{name:'tools',items:['Maximize','ShowBlocks']},'/',{name:'styles',items:['Styles','Format','Font','FontSize']},{name:'basicstyles',groups:['basicstyles','cleanup'],items:['Bold','Italic','Underline','Strike','Subscript','Superscript','-','RemoveFormat']},{name:'colors',items:['TextColor','BGColor']},{name:'others',items:['-']},{name:'about',items:['About']}],height:500});CKEDITOR.replace('editor_txt2',{toolbar:[{name:'document',groups:['mode','document','doctools'],items:['Source','-','Preview']},{name:'styles',items:['Format','Font','FontSize']},{name:'colors',items:['TextColor','BGColor']},{name:'basicstyles',groups:['basicstyles','cleanup'],items:['Bold','Italic','Underline','Strike','Subscript','Superscript','-','RemoveFormat']},{name:'links',items:['Link','Unlink']},]});$('#contentready').hide();editor_api.editor.content.obj=CKEDITOR.instances.editor_txt;$('#editor_txt').prev().removeAttr('style');sContent=editor_api.editor.content.getData();$('#introready').hide();editor_api.editor.intro.obj=CKEDITOR.instances.editor_txt2;sIntro=editor_api.editor.intro.getData()})};"
End Sub

%>