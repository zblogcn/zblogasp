<%
'****************************************
' weixin_Search 子菜单
'****************************************
Function weixin_Search_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("首页")
	aryPath=Array("main.asp")
	aryFloat=Array("m-left")
	aryInNewWindow=Array(False)
	For i=0 To Ubound(aryName)
		weixin_Search_SubMenu=weixin_Search_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
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



'首张图片提取-------------------------------------
Function GetFirstImg(str) '取得img 标签内容
	dim tmp,objRegExp,Matches,Match
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True '忽略大小写
	objRegExp.Global = false '全文搜索 !关键!
	objRegExp.Pattern = "<img (.*?)src=(.[^\[^>]*)(.*?)>"
	Set Matches =objRegExp.Execute(str)
	For Each Match in Matches
	tmp=tmp & Match.Value
	Next
	GetFirstImg=GetImgS(tmp)
End Function

Function GetImgS(str)'获取所有图片
	dim objRegExp1,mm,Match1,imgsrc
	Set objRegExp1 = New Regexp
	objRegExp1.IgnoreCase = True '忽略大小写
	objRegExp1.Global = True '全文搜索
	objRegExp1.Pattern = "src\=.+?\.(gif|jpg|png|bmp)"
	set mm=objRegExp1.Execute(str)
	For Each Match1 in mm
	imgsrc=Match1.Value
	'也许存在不能过滤的字符，确保万一
	imgsrc=replace(imgsrc,"""","")
	imgsrc=replace(imgsrc,"src=","")
	imgsrc=replace(imgsrc,"<","")
	imgsrc=replace(imgsrc,">","")
	imgsrc=replace(imgsrc,"img","")
	imgsrc=replace(imgsrc," ","")
	GetImgS=GetImgS&imgsrc'把里面的地址串起来备用
	next
End Function

'检验数据是否存在------------------------------------
Function CheckFields(ParameterName,FieldsName,TableName)
dim cRs
Set cRs=objConn.Execute("SELECT * FROM "&TableName&" Where "&ParameterName&" like '%" & FieldsName & "%'")
if not cRs.eof then
    CheckFields = cRs("fn_ID")
	else
	CheckFields = 0
End if
Set cRs = nothing
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