
<%'<!-- include file="AntiXSS.asp"-->
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    CommentEditor - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

ZC_UBB_ENABLE=True
ZC_UBB_AUTOLINK_ENABLE=False

Call Add_Response_Plugin("Response_Plugin_Html_js_Add","document.write(""<script type=\""text/javascript\"" src=\"""&_
BlogHost & "zb_users/plugin/commenteditor/commenteditor.pack.js\""></script>"");") 


'注册插件
Call RegisterPlugin("CommentEditor","ActivePlugin_CommentEditor")
'挂口部分
Function ActivePlugin_CommentEditor()
	'Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_TemplateTags","CommentEditor_ExportCommentHTML")
	'Call Add_Action_Plugin("Action_Plugin_BlogReBuild_Comments_Begin","BlogReBuild_Comments=CommentEditor_ReBuild():Exit Function")
End Function

'Function CommentEditor_ExportCommentHTML(t,v)
'	Dim s
'	s=v(7)
'	's=TransferHTML(s,"[anti-html-format]")
'	s=TransferHTML(s,"[nofollow]")
'	s=AntiXSS_run(s)
'	s=Replace(Replace(s,"&lt;!--r","<!--r"),"--&gt;","-->")
'	s=CommentEditor_AntiOther(s)
'	v(7)=s	
'	
''	'
'End Function
'
'Function CommentEditor_AntiOther(s)
'	Dim b
'	Set b=New RegExp
'	b.IgnoreCase=True
'	b.Pattern="<img\b[^<>]*?\bsrc[\s\t\r\n]*=[\s\t\r\n]*[""']?[\s\t\r\n]*([^\s\t\r\n""'<>]*)[^<>]*?/?[\s\t\r\n]*>"
'	b.Global=True
'	Dim c
'	Dim s1
'	s1=s
'	Set c=b.Execute(s)
''	Dim m
	'For Each m In c
	'	If Left(m.SubMatches(0),Len(BlogHost))<>BlogHost Then s1=b.Replace(m.value,"")'
'	Next
'	CommentEditor_AntiOther=s1
'End Function

'Function CommentEditor_ReBuild()

'	Call GetFunction()
'	If CStr(Functions(FunctionMetas.GetValue("comments")).SideBarID)="0" Then 
'
'		Exit Function
'	End If
	
'	Dim objRS
'	Dim objStream
'	Dim objArticle
'
	'Comments
'	Dim strComments

'	Dim s,t
'	Dim i,j

'	j=Functions(FunctionMetas.GetValue("comments")).MaxLi
'	If j=0 Then j=10
'
'	Set objRS=objConn.Execute("SELECT TOP "&j&" [log_ID],[comm_ID],[comm_Content],[comm_PostTime],[comm_AuthorID],[comm_Author] FROM [blog_Comment] WHERE [log_ID]>0 A'ND [comm_IsCheck]=0 ORDER BY [comm_PostTime] DESC,[comm_ID] DESC")
	'If (Not objRS.bof) And (Not objRS.eof) Then
	'	For i=1 to j
	'		Call GetUsersbyUserIDList(objRS("comm_AuthorID"))
'
'			Set objArticle=New TArticle
'
'			If objArticle.LoadInfoByID(objRS("log_ID")) Then
'				t=objArticle.FullUrl
''			End If
'
'			s=objRS("comm_Content")
'			s=Replace(s,vbCrlf,"")
''			s=Left(s,ZC_ARTICLE_EXCERPT_MAX)
	'		s=TransferHTML(s,"[anti-html-format]")
	'		s=TransferHTML(s,"[nohtml]")
	''		
		'	strComments=strComments & "<li style=""text-overflow:ellipsis;""><a href="""& t & "#cmt" & objRS("comm_ID") & """ title=""" & objRS("comm_PostTime") & " post by " & IIf(Users(objRS("comm_AuthorID")).Level=5,objRS("comm_Author"),Users(objRS("comm_AuthorID")).FirstName) & """>"+s+"</a></li>"
'			Set objArticle=Nothing
'			objRS.MoveNext
'			If objRS.eof Then Exit For
'		Next
'	End If
'	objRS.close
'	Set objRS=Nothing

'	strComments=TransferHTML(strComments,"[no-asp]")

'	Functions(FunctionMetas.GetValue("comments")).Content=strComments
'	Functions(FunctionMetas.GetValue("comments")).Post()
'	Functions(FunctionMetas.GetValue("comments")).SaveFile

'	CommentEditor_ReBuild=True

'End Function

%>
