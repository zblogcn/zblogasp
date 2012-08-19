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
		Totoro_Config.Write "TOTORO_BADWORD_LIST","虚拟主机|域名注册|服务器托管|host|铃声|彩信|营销|SEO|数据恢复|彩票|手机图片|游戏币|金币|交友中心|成人用品|私服|黄页|出租|求购|显示屏|投影仪|群发|翻译公司|留学咨询|外挂|google排名|婚庆公司|淘宝|皮肤病|不孕不育|性病|怀孕|医院|论坛群发|性感美女|乳腺病|尖锐湿疣|货到付款|汽车配件|推广联盟|劳务派遣|网络兼职|网络赚钱|证件公司|打包机|试验机|打标机|条码|标签纸|灌装机|升降机|升降平台|专业注册|网站建设|域名注册|出租网|六合彩|双色球|手机游戏|手机窃听|手机监听|成人小电影|激情视频|成人用品|成人电影|激情电影|二手电脑|出售枪支|枪支出售|高压电警棒|麻醉枪|麻醉乙醚|左轮手枪|私服|翻译公司|迷幻药|迷药|麻醉剂|迷昏药|催情药|蒙汗药|情趣用品|三唑仑|春药|张韶涵|诚招加盟|诚信经营|注册香港公司|公司注册|职业杀手|奇迹世界|工作服|免费电影|搬家公司|wow"
		Totoro_Config.Write "TOTORO_NUMBER_VALUE",10
		Totoro_Config.Write "TOTORO_REPLACE_KEYWORD","**"
		Totoro_Config.Write "TOTORO_REPLACE_LIST",""
		Totoro_Config.Write "TOTORO_CHINESESV",50

		Totoro_Config.Save
		Call SetBlogHint_Custom("您是第一次安装Totoro，已经为您导入初始配置。")
	End If
End Function

Function Totoro_Initialize()
	InstallPlugin_Totoro
	TOTORO_INTERVAL_VALUE=CLng(Totoro_Config.Read ("TOTORO_INTERVAL_VALUE"))
	TOTORO_BADWORD_VALUE=CLng(Totoro_Config.Read ("TOTORO_BADWORD_VALUE"))
	TOTORO_HYPERLINK_VALUE=CLng(Totoro_Config.Read ("TOTORO_HYPERLINK_VALUE"))
	TOTORO_NAME_VALUE=CLng(Totoro_Config.Read ("TOTORO_NAME_VALUE"))
	TOTORO_LEVEL_VALUE=CLng(Totoro_Config.Read ("TOTORO_LEVEL_VALUE"))
	TOTORO_SV_THRESHOLD=CLng(Totoro_Config.Read ("TOTORO_SV_THRESHOLD"))
	TOTORO_SV_THRESHOLD2=CLng(Totoro_Config.Read ("TOTORO_SV_THRESHOLD2"))
	TOTORO_DEL_DIRECTLY=CBool(Totoro_Config.Read ("TOTORO_DEL_DIRECTLY"))
	TOTORO_ConHuoxingwen=CBool(Totoro_Config.Read ("TOTORO_ConHuoxingwen"))
	TOTORO_BADWORD_LIST=Totoro_Config.Read ("TOTORO_BADWORD_LIST")
	TOTORO_NUMBER_VALUE=CLng(Totoro_Config.Read ("TOTORO_NUMBER_VALUE"))
	TOTORO_REPLACE_KEYWORD=Totoro_Config.Read ("TOTORO_REPLACE_KEYWORD")
	TOTORO_REPLACE_LIST=Totoro_Config.Read ("TOTORO_REPLACE_LIST")
	TOTORO_CHINESESV=Totoro_Config.Read("TOTORO_CHINESESV")
End Function


Function Totoro_Xiou(strContent)
	Dim a,b,c,d,text
	text=strContent
	Set a=New RegExp
	a.Pattern="&#(\d+?);"
	a.Global=True
	Set b=a.Execute(text)
	For Each c In b
		d = CLng(c.Submatches(0))
		If d - 65536 > 0 Then
			d = d - 65536
		End If
		text = Replace(text, c.value, ChrW(d))
	Next
	Totoro_Xiou=text
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
		strTemp=Totoro_Xiou(strTemp)
		strTemp=Totoro_FxxxHuoxingwen(strTemp)
		strTemp=Totoro_FromSBCCode(strTemp)
		strTemp=Totoro_GetNum(strTemp)		
	End If
	Call Totoro_checkLevel(BlogUser.Level)
	Call Totoro_checkName(Request.ServerVariables("REMOTE_ADDR"))
	Call Totoro_checkHyperLink(strTemp)
	Call Totoro_checkBadWord(strTemp & "&" & objComment.Author & "&" & objComment.HomePage & "&" & objComment.IP & "&" & objComment.Email)
	Call Totoro_checkInterval(Request.ServerVariables("REMOTE_ADDR"))
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


Function Totoro_checkName(ByVal ip)

	If TOTORO_NAME_VALUE=0 Then Exit Function

	Dim i,s
	s=FilterSQL(ip)

	Dim objRS
	Set objRS=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [log_ID]>=0 and [comm_IP] ='" & ip & "'")
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
	If TOTORO_REPLACE_LIST="" Then Totoro_replaceWord=content:Exit Function
	Set o=New RegExp
	o.Pattern=TOTORO_REPLACE_LIST
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

	Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2^matches.count-1)

	Set SRegExp=Nothing

End Function


Function Totoro_checkInterval(ByVal ip)

	If TOTORO_INTERVAL_VALUE=0 Then Exit Function
	Dim i,j,t,s,m,n
	Dim objRS
	m="SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_IP] ='" & ip & "'"
	m=m&" AND [comm_PostTime]>"&ZC_SQL_POUND_KEY&DateAdd("h", -1, now)&ZC_SQL_POUND_KEY
	Set objRS=objConn.Execute(m)
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS(0)
	End If
	Set objRS=Nothing
	If i>0 Then
		If i<=10 Then
			s=TOTORO_INTERVAL_VALUE*1/5
		ElseIf i>10 And i<=20 Then
			s=TOTORO_INTERVAL_VALUE*2/5
		Elseif i>20 And i<=30  Then
			s=TOTORO_INTERVAL_VALUE*3/5
		ElseIf i>30 And i<=40  Then
			s=TOTORO_INTERVAL_VALUE*4/5
		ElseIf i>40 And i<=50  Then
			s=TOTORO_INTERVAL_VALUE*5/5
		Else
			s=TOTORO_INTERVAL_VALUE*6/5
		End If
	End If

	Totoro_SV=Totoro_SV+s

End Function




'*********************************************************
' 目的：    
'*********************************************************
Function Totoro_GetSpamCount_Comment()
	If IsEmpty(objConn)=True Then Exit Function
	Dim objRS1
	Set objRS1=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_isCheck]=-1")
	If (Not objRS1.bof) And (Not objRS1.eof) Then
		Totoro_SpamCount_Comment="("&objRS1(0)&"条未审核)"
	End If

	'评论管理加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_CommentMng_SubMenu",MakeSubMenu("审核评论" & Totoro_SpamCount_Comment,GetCurrentHost() & "zb_users/plugin/totoro/setting1.asp","m-left",False)& "<scr" & "ipt src="""&GetCurrentHost&"/zb_users/plugin/totoro/common.js"" type=""text/javascript""></scr" & "ipt><scr" & "ipt src="""&GetCurrentHost&"/zb_users/plugin/totoro/cmmng.js"" type=""text/javascript""></scr" & "ipt>")
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