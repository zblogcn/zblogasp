<!-- #include file="function.asp"-->
<%

'注册插件
Call RegisterPlugin("ReplaceWord","ActivePlugin_ReplaceWord")
'挂口部分
Function ActivePlugin_ReplaceWord()
	
	
	
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Core","ReplaceWord_")
	Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","ReplaceWord__")
	Call Add_Filter_Plugin("Filter_Plugin_RegPage_Vaild","Replaceword___")
	
End Function

Function ReplaceWord_(obj)
	replaceword.init()
	replaceword.orig=obj.Content
	obj.Content=replaceword.replace()
	replaceword.orig=obj.Intro
	obj.Intro=replaceword.replace()
	replaceword.orig=obj.Title
	obj.Title=replaceword.replace()
End Function

Function ReplaceWord__(obj)
	replaceword.init()
	replaceword.orig=obj.Content
	obj.Content=replaceword.replace()
	replaceword.orig=obj.Author
	obj.Author=replaceword.replace()
End Function

Function Replaceword___(a,b,c,d)
	Stop
	replaceword.init()
	replaceword.orig=a
	replaceword.replace()
	if replaceword.string<>replaceword.orig Then ExportErr "含有敏感词,请更换用户名!"
End Function
%>