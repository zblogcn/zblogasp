<%
'****************************************
' weixin 子菜单
'****************************************
Function weixin_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("基本设置","微信连接设置")
	aryPath=Array("main.asp","tokenst.asp")
	aryFloat=Array("m-left","m-left")
	aryInNewWindow=Array(False,False)
	For i=0 To Ubound(aryName)
		weixin_SubMenu=weixin_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function


'////=================================================================================////
'去掉html标签
Function RemoveHTML( strText )
    Dim TAGLIST
    TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" &_
              "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" &_
              "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" &_
              "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" &_
              "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" &_
              "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;P;PARAM;PLAINTEXT;" &_
              "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" &_
              "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"
    Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"     
    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock     
    nPos1 = InStr(strText, "<")
    Do While nPos1 > 0
        nPos2 = InStr(nPos1 + 1, strText, ">")
        If nPos2 > 0 Then
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
        strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

            nPos3 = InStr(strTagName, " ")
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)
            End If
If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
            Else
                bSearchForBlock = True
            End If
            
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                bRemove = True
                If bSearchForBlock Then
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            Else
bRemove = False
            End If
            
            If bRemove Then
                strResult = strResult & Left(strText, nPos1 - 1)
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    strResult = strResult & strText
    
    RemoveHTML = strResult
End Function

'首次关注
Function Welcome()
	Welcome = Content & "欢迎关注《未寒博客》！！！"& VBCrLf
	Welcome = Welcome & "您可发送“最新文章”来查看博客最新的10篇文章，或者直接发送关键词来搜索博客中已发表的文章。更多使用帮助请输入英文“help”或者数字“0”来查看。"
End Function

'查询文章
Function Search(Content)
	Dim LTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("SELECT TOP 15  [log_ID],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_FullUrl] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )")
	Do Until LTRS.Eof
		InserNewHtml = InserNewHtml & "<a href=""" & ZC_BLOG_HOST & "ZB_USERS/plugin/weixin/view.asp?wid=" & LTRS("log_ID") & """>" & LTRS("log_ID") & "、" & LTRS("log_Title") & "</a>" & VBCrLf & VBCrLf 'LTRS("log_PostTime") & 
		'InserNewHtml = InserNewHtml & TransferHTML(LTRS("log_Content"),"[nohtml]")
		'Exit Do
		LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	InserNewHtml = Replace(InserNewHtml,"&nbsp;"," ")
	InserNewHtml = Replace(InserNewHtml,"<#ZC_BLOG_HOST#>",BlogHost)
	
	Content = "“" & Content & "”搜索结果：" & VBCrLf
	Search = Content & InserNewHtml & VBCrLf & "  提示：请直接点击文章标题查看博客文章，或者回复标题前的编号直接在微信中查看文字版。"
End Function

'最新文章
Function LastPost()
	Dim LTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("SELECT TOP 5 [log_ID], [log_Title], [log_Intro], [log_Content], [log_PostTime], [log_Type] FROM blog_Article WHERE ((([log_Type])=0)) ORDER BY [log_PostTime] DESC")
	Do Until LTRS.Eof
		InserNewHtml = InserNewHtml & "<item><Title><![CDATA[" & LTRS("log_Title") & "]]></Title><Description><![CDATA[" & TransferHTML(LTRS("log_Intro"),"[nohtml]") & "]]></Description><PicUrl><![CDATA["

		if GetFirstUrl(LTRS("log_Content"))="" then
			InserNewHtml = InserNewHtml & "http://imzhou.com/zb_system/image/logo/zblog.gif"
		else
			InserNewHtml = InserNewHtml & GetFirstUrl(LTRS("log_Content"))
		End if

		InserNewHtml = InserNewHtml & "]]></PicUrl><Url><![CDATA[" & ZC_BLOG_HOST & "ZB_USERS/plugin/weixin/view.asp?wid=" & LTRS("log_ID") & "]]></Url></item>"
		LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	InserNewHtml = Replace(InserNewHtml,"&nbsp;"," ")
	InserNewHtml = Replace(InserNewHtml,"<#ZC_BLOG_HOST#>",BlogHost)

	LastPost = InserNewHtml
End Function

'=======================================================
'函数: 从正文中提取图片路径.
'输入: 文章全文.
'返回: 有图则返回图片路径, 无图返回空.
'=======================================================
Function GetFirstUrl(ByVal strContent)
	'On Error Resume Next
	Dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase=True
	objRegExp.Global=False

	objRegExp.Pattern="(<img[^>]+(src|data-original)[^""]+"")([^""]+)([^>]+>)"

	Dim Match, Matches, Value
	Set Matches=objRegExp.Execute(strContent)
		For Each Match in Matches
			Value=objRegExp.Replace(Match.value,"$3")
		Next
	Set Matches=Nothing

	Set objRegExp=Nothing

	GetFirstUrl=Value

	'Err.Clear
End Function


'Call RsFilter(数量,是否提取图片,提取内容,表名,筛选特性,排列方式,输入框名)
'数据库提取-------------------------------------
Function RsFilter(LTamount,LTImg,LTFilter,LTList,LTWhere,LTType,Lt_textarea)
	dim LTRS,CTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("select top "&LTamount&" "&LTFilter&" from "&LTList&" WHERE "&LTWhere&" order by "&LTType&"")
	Do Until LTRS.Eof
		Set CTRS=objConn.Execute("select * from [blog_Category] WHERE [cate_ID]="&LTRS("log_CateID")&" order by [cate_ID] DESC")
		InserNewHtml = InserNewHtml & "<li><a href="""&LTRS("log_FullUrl")&""" Title=""" & LTRS("log_Title") &"""><span>" & LTRS("log_Title") &"</span><em class=""PostTime"" style=""display:none;"">"&LTRS("log_PostTime")&"</em><em class=""ArtCategory"" style=""display:none;"">"&CTRS("cate_Name")&"</em>"
		
		if LTImg = True then
			if GetFirstImg(LTRS("log_Content"))="" then
				InserNewHtml = InserNewHtml & " "
			else
				InserNewHtml = InserNewHtml & "<img class=""NewArtImg"" style=""display:none;"" src="""&ZC_BLOG_HOST&"zb_users/"&GetFirstImg(LTRS("log_Content"))&"""/>"
			End if
		End if
		InserNewHtml = InserNewHtml & "</a>"
		InserNewHtml = InserNewHtml & "<div class=""ArtInfoList"" style=""display:none;"">"&RemoveHTML(LTRS("log_Intro"))&"</div>"
		InserNewHtml = InserNewHtml & "</li>"
	LTRS.MoveNext
	Loop
	RsFilter = InserNewHtml
	Set LTRS=Nothing
	Set CTRS=Nothing
End Function



'检查最新文章列表====================================
Function CheckNewArticle()
Dim objFunction,objConfig
Set objConfig=New TConfig
objConfig.Load("ListType")
Set objFunction=New TFunction
if CheckFields("fn_FileName","newartfile","blog_Function") = 0 Then
	objFunction.ID=0
	objFunction.Name="最新文章"
	objFunction.FileName="NewArtFile"
	objFunction.HtmlID="divNewArt"
	objFunction.Ftype="ul"
	objFunction.Order=15
	objFunction.MaxLi=0
	objFunction.SidebarID=10000
	else
	objFunction.LoadInfoByID(CheckFields("fn_FileName","newartfile","blog_Function"))
End if
	objFunction.IsSystem=True
	objFunction.Content=RsFilter(objConfig.Read("SetNewArt"),True,"*","blog_Article","[log_Type]=0","Log_PostTime DESC","inpContent")
	objFunction.save
	Call SaveFunctionType()
	Call MakeBlogReBuild_Core()
End Function

%>