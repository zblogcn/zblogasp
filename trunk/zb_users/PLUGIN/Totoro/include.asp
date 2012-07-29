<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////


Dim TOTORO_INTERVAL_VALUE
Dim TOTORO_BADWORD_VALUE
Dim TOTORO_HYPERLINK_VALUE
Dim TOTORO_NAME_VALUE
Dim TOTORO_LEVEL_VALUE
Dim TOTORO_SV_THRESHOLD
Dim TOTORO_SV_THRESHOLD2
Dim TOTORO_DEL_DIRECTLY
Dim TOTORO_ConHuoxingwen
Dim TOTORO_BADWORD_LIST
Dim TOTORO_NUMBER_VALUE
Dim TOTORO_REPLACE_KEYWORD
Dim TOTORO_REPLACE_LIST
Dim TOTORO_CHINESESV

Dim Totoro_Config
'Const TOTORO_SV_THRESHOLD = 50
 
'Const TOTORO_LEVEL_VALUE = 100
'Const TOTORO_NAME_VALUE = 45
'Const TOTORO_HYPERLINK_VALUE = 10
'Const TOTORO_BADWORD_VALUE = 50
'Const TOTORO_BADWORD_LIST = "虚拟主机|域名注册|服务器托管|hosting|poker|免费铃声|免费彩信|铃声下载|搜'索引擎营销|数据恢复|彩票软件|手机图片|魔兽金币|交友中心|成人用品|私服|企业黄页|出租|显示屏|投影仪|群发|翻译公司|留学咨询|外挂|硬盘录像机|google排名|注册香港公司|婚庆公司|投影幕|培养箱|花店|一号通|印刷公司|打包机|封口机|管件|砂机|打标机|升降机"
'Const TOTORO_INTERVAL_VALUE = 25
'Const TOTORO_INTERVAL_VALUE = 1
'Const TOTORO_DEL_DIRECTLY = False
'Const TOTORO_ConHuoxingwen = True
'Const TOTORO_NUMBER_VALUE=10

Dim Totoro_SV
Totoro_SV=0

Dim Totoro_SpamCount_Comment

'注册插件
Call RegisterPlugin("Totoro","ActivePlugin_Totoro")


'具体的接口挂接
Function ActivePlugin_Totoro() 

	'挂上接口
	'Filter_Plugin_PostComment_Core
	Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","Totoro_chkComment")
	'Action_Plugin_Admin_Begin
	Call Add_Action_Plugin("Action_Plugin_Admin_Begin","If Request.QueryString(""act"")=""CommentMng"" Then Call Totoro_GetSpamCount_Comment() End If")
	'网站管理加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("Totoro设置",GetCurrentHost() & "zb_users/plugin/totoro/setting.asp","m-left",False))

End Function


Function InstallPlugin_Totoro()
	Set Totoro_Config = New TConfig
	Totoro_Config.Load("Totoro")
	If Totoro_Config.Exists("TOTORO_VERSION")=False Then
		Totoro_Config.Write "TOTORO_VERSION","0.0"
		Totoro_Config.Write "TOTORO_INTERVAL_VALUE",25
		Totoro_Config.Write "TOTORO_BADWORD_VALUE",50
		Totoro_Config.Write "TOTORO_HYPERLINK_VALUE",10
		Totoro_Config.Write "TOTORO_NAME_VALUE",45
		Totoro_Config.Write "TOTORO_LEVEL_VALUE",100
		Totoro_Config.Write "TOTORO_SV_THRESHOLD",50
		Totoro_Config.Write "TOTORO_SV_THRESHOLD2",150

		Totoro_Config.Write "TOTORO_DEL_DIRECTLY","False"
		Totoro_Config.Write "TOTORO_ConHuoxingwen","True"
		Totoro_Config.Write "TOTORO_BADWORD_LIST","虚拟主机|域名注册|服务器托管|host|铃声|彩信|营销|SEO|数据恢复|彩票|手机图片|游戏币|金币|交友中心|成人用品|私服|黄页|出租|求购|显示屏|投影仪|群发|翻译公司|留学咨询|外挂|google排名|婚庆公司|淘宝|皮肤病|不孕不育|性病|怀孕|医院"
		Totoro_Config.Write "TOTORO_NUMBER_VALUE",10
		Totoro_Config.Write "TOTORO_REPLACE_KEYWORD","**"
		Totoro_Config.Write "TOTORO_REPLACE_LIST",""
		Totoro_Config.Write "TOTORO_CHINESESV",50

		Totoro_Config.Save
		Call SetBlogHint_Custom("您是第一次安装Totoro，已经为您导入初始配置。")
	End If
End Function

Function Totoro_Initialize()
	Set Totoro_Config = New TConfig
	Totoro_Config.Load("Totoro")
	TOTORO_INTERVAL_VALUE=CLng(Totoro_Config.Read ("TOTORO_INTERVAL_VALUE"))
	TOTORO_BADWORD_VALUE=CLng(Totoro_Config.Read ("TOTORO_BADWORD_VALUE"))
	TOTORO_HYPERLINK_VALUE=CLng(Totoro_Config.Read ("TOTORO_HYPERLINK_VALUE"))
	TOTORO_NAME_VALUE=CLng(Totoro_Config.Read ("TOTORO_NAME_VALUE"))
	TOTORO_LEVEL_VALUE=CLng(Totoro_Config.Read ("TOTORO_LEVEL_VALUE"))
	TOTORO_SV_THRESHOLD=CLng(Totoro_Config.Read ("TOTORO_SV_THRESHOLD"))
	TOTORO_SV_THRESHOLD2=CLng(Totoro_Config.Read ("TOTORO_SV_THRESHOLD2"))
	TOTORO_DEL_DIRECTLY=Totoro_Config.Read ("TOTORO_DEL_DIRECTLY")
	TOTORO_ConHuoxingwen=Totoro_Config.Read ("TOTORO_ConHuoxingwen")
	TOTORO_BADWORD_LIST=Totoro_Config.Read ("TOTORO_BADWORD_LIST")
	TOTORO_NUMBER_VALUE=CLng(Totoro_Config.Read ("TOTORO_NUMBER_VALUE"))
	TOTORO_REPLACE_KEYWORD=Totoro_Config.Read ("TOTORO_REPLACE_KEYWORD")
	TOTORO_REPLACE_LIST=Totoro_Config.Read ("TOTORO_REPLACE_LIST")
	TOTORO_CHINESESV=Totoro_Config.Read("TOTORO_CHINESESV")
End Function
'*********************************************************
' 目的：    检查评论
'*********************************************************
Function Totoro_chkComment(ByRef objComment)
	Call Totoro_Initialize
	
	If objComment.IsCheck=True Then Exit Function
	If objComment.IsThrow=True Then Exit Function
	
	Dim strTemp
	strTemp=objComment.Content
	If TOTORO_ConHuoxingwen Then
		strTemp=Totoro_FxxxHuoxingwen(strTemp)
		strTemp=Totoro_FromSBCCode(strTemp)
		strTemp=Totoro_GetNum(strTemp)
	End If
	Call Totoro_checkLevel(BlogUser.Level)
	Call Totoro_checkName(objComment.Author)
	Call Totoro_checkHyperLink(strTemp)
	Call Totoro_checkBadWord(strTemp & "&" & objComment.Author & "&" & objComment.HomePage & "&" & objComment.IP & "&" & objComment.Email)
	Call Totoro_checkInterval(Request.ServerVariables("REMOTE_ADDR"),now,true)
	Call Totoro_checkNumLong(strTemp)
	Call Totoro_checkChinese(strTemp)
	objComment.Content=Totoro_replaceWord(objComment.Content)

	If Totoro_SV>=TOTORO_SV_THRESHOLD Then
		ZVA_ErrorMsg(14)="Totoro Ⅲ" & "插件大显神威!" & ZVA_ErrorMsg(14)
		ZVA_ErrorMsg(53)="Totoro Ⅲ" & "插件大显神威!" & ZVA_ErrorMsg(53)
		If Totoro_SV<TOTORO_SV_THRESHOLD2 Or TOTORO_SV_THRESHOLD2=0 Then
			objComment.IsCheck=True
		ElseIf TOTORO_SV_THRESHOLD2<=Totoro_SV Then
			objComment.IsThrow=True
		End If
	End If
	
	Totoro_chkComment=True

End Function
'*********************************************************

Function Totoro_checkChinese(Content)
	Dim a
	a=CheckRegExp(Content,"[\u4e00-\u9fa5]")
	If a=False Then Totoro_SV=Totoro_SV+TOTORO_CHINESESV
End Function

Function Totoro_checkLevel(ByVal level)

	If TOTORO_LEVEL_VALUE=0 Then Exit Function

	If Level=1 Then
	Totoro_SV=Totoro_SV-TOTORO_LEVEL_VALUE*(8)
	ElseIf  Level=2 Then
	Totoro_SV=Totoro_SV-TOTORO_LEVEL_VALUE*(4)
	ElseIf  Level=3 Then
	Totoro_SV=Totoro_SV-TOTORO_LEVEL_VALUE*(2)
	ElseIf  Level=4 Then
	Totoro_SV=Totoro_SV-TOTORO_LEVEL_VALUE*(1)
	ElseIf  Level=5 Then
	Totoro_SV=Totoro_SV-TOTORO_LEVEL_VALUE*(0)
	End If
End Function


Function Totoro_checkName(ByVal name)

	If TOTORO_NAME_VALUE=0 Then Exit Function

	Dim i,s
	s=FilterSQL(name)

	Dim objRS
	Set objRS=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [log_ID]>=0 and [comm_Author] ='" & s & "'")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS(0)
	End If
	Set objRS=Nothing

	If i>0 And i<=10   Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(0)
	If i>10 And i<=20  Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(1)
	If i>20 And i<=50 Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(2)
	If i>50           Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(3)

End Function


Function Totoro_checkBadWord(ByVal content)
	If Totoro_SV+TOTORO_BADWORD_VALUE=0 Then Exit Function
	Dim o
	Set o=New RegExp
	o.Pattern=TOTORO_BADWORD_LIST
	o.Global=True
	o.IgnoreCase=True
	Dim j,k
	j=len(o.replace(content,""))
	k=len(content)
	j=k-j
	Set o=Nothing
	Totoro_SV=Totoro_SV+TOTORO_BADWORD_VALUE*j
End Function

Function Totoro_replaceWord(content)
	Dim o
	Set o=New RegExp
	o.Pattern=TOTORO_REPLACE_LIST&"|"&TOTORO_BADWORD_LIST
	o.Global=True
	o.IgnoreCase=True
	Totoro_replaceWord=o.replace(content,TOTORO_REPLACE_KEYWORD)
	Set o=Nothing
End Function



Function Totoro_checkHyperLink(ByVal content)

	If TOTORO_HYPERLINK_VALUE=0 Then Exit Function

	Dim SRegExp,Matches
	Set SRegExp=New RegExp
	SRegExp.IgnoreCase =True
	SRegExp.Global=True
	SRegExp.Pattern="https:|http:|ftp|www."
	Set Matches = SRegExp.Execute(content)

	If Matches.count=0 Then
		Totoro_SV=Totoro_SV
	ElseIf  Matches.count=1 Then
		Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2-1)
	ElseIf  Matches.count=2 Then
		Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2*2-1)
	ElseIf  Matches.count=3 Then
		Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2*2*2-1)
	ElseIf  Matches.count=4 Then
		Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2*2*2*2-1)
	ElseIf  Matches.count=5 Then
		Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2*2*2*2*2-1)
	Else
		Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2*2*2*2*2*2-1)
	End If

	Set SRegExp=Nothing

End Function


Function Totoro_checkInterval(ByVal ip,ByVal posttime,ByVal iscomment)

	If TOTORO_INTERVAL_VALUE=0 Then Exit Function

	Dim i,j,t,s,m,n
	Dim objRS
	m="SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_IP] ='" & ip & "'"
	n="SELECT [comm_PostTime] FROM [blog_Comment] WHERE [comm_IP] ='" & ip & "'"

	Set objRS=objConn.Execute(m)
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS(0)
	End If
	Set objRS=Nothing
	s=0
	If i>0 Then
		Set objRS=objConn.Execute(n)
			If (Not objRS.bof) And (Not objRS.eof) Then
				Do While Not objRS.eof
					t=objRS("comm_PostTime")
					If DateDiff("h",t,posttime)<TOTORO_INTERVAL_VALUE Then
						j=j+1
						If     DateDiff("n",t,posttime)>((TOTORO_INTERVAL_VALUE*60)\5)*4 Then
							s=s+(TOTORO_INTERVAL_VALUE\5)*1
						ElseIf DateDiff("n",t,posttime)>((TOTORO_INTERVAL_VALUE*60)\5)*3 Then
							s=s+(TOTORO_INTERVAL_VALUE\5)*2
						ElseIf DateDiff("n",t,posttime)>((TOTORO_INTERVAL_VALUE*60)\5)*2 Then
							s=s+(TOTORO_INTERVAL_VALUE\5)*3
						ElseIf DateDiff("n",t,posttime)>((TOTORO_INTERVAL_VALUE*60)\5)*1 Then
							s=s+(TOTORO_INTERVAL_VALUE\5)*4
						ElseIf DateDiff("n",t,posttime)>((TOTORO_INTERVAL_VALUE*60)\5)*0 Then
							s=s+(TOTORO_INTERVAL_VALUE\5)*5
						Else
							s=s+(TOTORO_INTERVAL_VALUE\5)*6
						End If
					End If
					objRS.MoveNext
				Loop
			End If
		Set objRS=Nothing
	End If

	Totoro_SV=Totoro_SV+s

End Function




'*********************************************************
' 目的：    错误退出
'*********************************************************
Function Totoro_ExitError(strInput)
	If IsEmpty(Request.Form("inpAjax"))=False Then
		Call RespondError(vbObjectError+1,strInput)
		Response.End
	End If
	Response.Redirect ZC_BLOG_HOST & "function/c_error.asp?errorid=" & 0 & "&number=" & (vbObjectError+1) & "&description=" & Server.URLEncode(strInput) & "&source="
End Function
'*********************************************************


'*********************************************************
' 目的：    
'*********************************************************
Function Totoro_GetSpamCount_Comment()
	If IsEmpty(objConn)=True Then Exit Function
	Dim objRS1
	Set objRS1=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_isCheck]=1")
	If (Not objRS1.bof) And (Not objRS1.eof) Then
		Totoro_SpamCount_Comment="("&objRS1(0)&"条未审核的评论)"
	End If

	'评论管理加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_CommentMng_SubMenu",MakeSubMenu("审核评论" & Totoro_SpamCount_Comment,GetCurrentHost() & "zb_users/plugin/totoro/setting1.asp","m-left",False))
' & "<scr" & "ipt src=""../plugin/totoro/common.js"" type=""text/javascript""></scr" & "ipt><scr" & "ipt src=""../plugin/totoro/cmmng.js"" type=""text/javascript""></scr" & "ipt>"
End Function
'*********************************************************


Function Totoro_FxxxHuoxingwen(str)
	Dim a,b,d
	d=str
	a=Array("҉|","蕶","ニ|貳","弎","陸","ハ|仈","艽","ā|á|ǎ|à|а|А|α","в|в|В|ъ|Ъ|ы|Ы|ь|Ь|β","с|с|С","Ё|е|Е|ё|Ё|ê|ē|é|ě|è","℉|ｆ","ɡ","н|Н","ī|í|ǐ|ì","ｊ","κ","ι","м|М","ń|п|П|Й|π","0|ō|ó|ǒ|ǒ|о|О|ο|σ|⊙|○|◎","р|Р|ρ","я|Я","\$","т|Т|τ","ū|ú|ǔ|ù|∪|μ|υ","∨|ν","ω","×|х|Х|χ","у|У|γ","э|Э","θ","ф|Ф")
	b=Array("",0,2,3,6,8,9,"a","b","c","e","f","g","h","i","j","k","l","m","n","o","p","r","s","t","u","v","w","x","y",3,8,"中")
	Dim c,i
	set c=new regexp

	c.Global=True
	c.IgnoreCase=True
	For i=0 To ubound(a)
		c.Pattern=a(i)
		d=c.replace(d,b(i))
	Next
	Totoro_FxxxHuoxingwen=d
	set c=nothing
End Function

Function Totoro_FromSBCCode(str)
	Dim a,b,c
	a="ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ１２３４５６７８９０ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ［］；＇／．，＜＞？＂：{}｜＋＿＼＝－）（＊＆＾％＄#＠！￣"
	b="ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890abcdefghijklmnopqrstuvwxyz[];'/.,<>?"":{}|+_\=-)(*&^%$#@!~"
	For c=1 To Len(a)
		str=Replace(str,Mid(a,c,1),Mid(b,c,1))
	Next
	Totoro_FromSBCCode=str
End Function

Function Totoro_GetNum(str)
	Dim a,b,d
	d=str
	a=Array("零|〇"," 一|壹|Ⅰ|⒈|㈠|①|⑴","二|贰|Ⅱ|⒉|㈡|②|⑵","三|叁|Ⅲ|⒊|㈢|③|⑶","四|肆|Ⅳ|⒋|㈣|④|⑷","五|伍|Ⅴ|⒌|⑤|㈤|⑸","六|陆|Ⅵ|⒍|㈥|⑥|⑹","七|柒|Ⅶ|⒎|⑦|㈦|⑺","八|捌|Ⅷ|⒏|㈧|⑧|⑻","九|玖|Ⅸ|⒐|⑨|㈨|⑼")
	b=Array(0,1,2,3,4,5,6,7,8,9)
	Dim c,i
	set c=new regexp
	c.Global=True
	c.IgnoreCase=True
	For i=0 To 9
		c.Pattern=a(i)
		d=c.replace(d,b(i))
	Next
	Totoro_GetNum=d
	set c=nothing

End Function



Function Totoro_CheckNumLong(str)
	Dim a,b,c
	set c=new regexp
	c.global=true
	c.pattern="\d"
	b=str
	b=c.replace(b,"")
	a=len(str)-len(b)
	if a>10 then
		Totoro_SV=Totoro_SV+TOTORO_NUMBER_VALUE*(a-10)
	end if
	Totoro_CheckNumLong=True
	set c=nothing
End Function

%>