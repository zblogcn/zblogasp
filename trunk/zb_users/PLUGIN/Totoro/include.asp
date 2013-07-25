<!-- #include file="tran2simp.asp"-->
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
Dim Totoro_DebugData
Dim Totoro_Debug
Dim TOTORO_INTERVAL_VALUE
Dim TOTORO_BADWORD_VALUE
Dim TOTORO_HYPERLINK_VALUE
Dim TOTORO_NAME_VALUE
Dim TOTORO_LEVEL_VALUE
Dim TOTORO_SV_THRESHOLD
Dim TOTORO_SV_THRESHOLD2
Dim TOTORO_DEL_DIRECTLY
Dim TOTORO_CONHUOXINGWEN
Dim TOTORO_BADWORD_LIST
Dim TOTORO_NUMBER_VALUE
Dim TOTORO_REPLACE_KEYWORD
Dim TOTORO_REPLACE_LIST
Dim TOTORO_CHINESESV
Dim TOTORO_KILLIP
Dim TOTORO_TRANTOSIMP
Dim TOTORO_FILTERIP
Dim TOTORO_PM
Dim TOTORO_THROWCOUNT
Dim TOTORO_CHECKCOUNT

Dim TOTORO_CHECKSTR
Dim TOTORO_THROWSTR
Dim TOTORO_KILLIPSTR

Dim Totoro_Config


Dim Totoro_SV
Totoro_SV=0

Dim Totoro_SpamCount_Comment

'注册插件
Call RegisterPlugin("Totoro","ActivePlugin_Totoro")
Sub Totoro_AddDebug(data)
	If Totoro_Debug Then Totoro_DebugData=Totoro_DebugData & vbCrlf & "【" & Now & "】" & data
End Sub

'具体的接口挂接
Function ActivePlugin_Totoro() 

	'挂上接口
	'Filter_Plugin_PostComment_Core
	Call Add_Filter_Plugin("Filter_Plugin_PostComment_Core","Totoro_chkComment")
	'Action_Plugin_Admin_Begin
	Call Add_Filter_Plugin("Filter_Plugin_CommentAduit_Core","Totoro_WriteConfig")
	
	'Call Add_Action_Plugin("Action_Plugin_Admin_Begin","If Request.QueryString(""act"")=""CommentMng"" Then Call Totoro_GetSpamCount_Comment() End If")
	'网站管理加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("TotoroⅢ设置",GetCurrentHost() & "zb_users/plugin/totoro/setting.asp","m-left",False))
	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_CommentMng_SubMenu",MakeSubMenu("TotoroⅢ设置",GetCurrentHost() & "zb_users/plugin/totoro/setting.asp","m-right",False)&"<script type='text/javascript'>$(document).ready(function(){$('#divMain2').before('<div id=\'totoro\'>Totoro is getting data..</div>');$.get('"&GetCurrentHost() & "zb_users/plugin/totoro/getcount.asp',{'rnd':Math.random()},function(txt){$('#totoro').html(txt)})})</script>")


	'Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"TotoroⅢ",GetCurrentHost&"zb_users/plugin/totoro/main.asp","nav_totoro","aTotoro",GetCurrentHost&"zb_users/plugin/totoro/antivirus-alt.png"))

End Function

Function Totoro_WriteConfig(oTop,isA)
	Dim strTmp2,strTmp
	
	If isA=False And oTop.isCheck=True Then
		Totoro_Initialize
		
		strTmp2=Totoro_Config.Read("TOTORO_BADWORD_LIST")
		strTmp=oTop.HomePage & "|" & oTop.Content
	
		Dim objReg,objMatches,Match
		Set objReg = New RegExp
		objReg.IgnoreCase = True
		objReg.Global = True
		objReg.Pattern = "(([\w\d]+\.)+\w{2,})"
		Set objMatches = objReg.Execute(strTmp)
		For Each Match In objMatches
			If CheckRegExp(Match.SubMatches(0),strTmp2)=False Then
				strTmp2=strTmp2 & "|" & Replace(Match.SubMatches(0),".","\.")
				SetBlogHint_Custom "Totoro新增黑词"& Replace(Match.SubMatches(0),".","\.")
			End if
		Next
		Set objReg = Nothing
		Set objMatches = Nothing
		Set Match = Nothing
		If left(strTmp2,1)="|" then strTmp2=Right(strTmp2, Len(strTmp2) - 1)

		Totoro_Config.Write "TOTORO_BADWORD_LIST",strTmp2
		Totoro_Config.Save
	End If
End Function

Function InstallPlugin_Totoro()
	Set Totoro_Config = New TConfig
	Totoro_Config.Load("Totoro")
	If Totoro_Config.Exists("TOTORO_VERSION")=False Then
		Totoro_Config.Write "TOTORO_VERSION","3.0.4"
		Totoro_Config.Write "TOTORO_INTERVAL_VALUE",25
		Totoro_Config.Write "TOTORO_BADWORD_VALUE",50
		Totoro_Config.Write "TOTORO_HYPERLINK_VALUE",10
		Totoro_Config.Write "TOTORO_NAME_VALUE",45
		Totoro_Config.Write "TOTORO_LEVEL_VALUE",100
		Totoro_Config.Write "TOTORO_SV_THRESHOLD",50
		Totoro_Config.Write "TOTORO_SV_THRESHOLD2",150

		Totoro_Config.Write "TOTORO_DEL_DIRECTLY","False"
		Totoro_Config.Write "TOTORO_ConHuoxingwen","True"
		Totoro_Config.Write "TOTORO_BADWORD_LIST","(推广|群发|广告|解密|赌博|包青天|广告|阿凡提|发贴|顶贴|(针孔|隐形|隐蔽式)摄像|干扰|顶帖|发帖|消声|遥控|解码|窃听|身份证生成|拦截|复制|监听|定位|消声|作弊|扩散|侦探|追杀)(机|器|软件|设备|系统)|(求|换|有偿|买|卖|出售)(肾|器官|眼角膜|血)|肾源|(假|毕业)(证|文凭|发票|币)|(手榴|人|麻醉|霰)弹|治疗(肿瘤|乙肝|性病|红斑狼疮)|重亚硒酸钠|(粘氯|原砷)酸|麻醉乙醚|原藜芦碱A|永伏虫|蝇毒|罂粟|银氰化钾|氯胺酮|因毒(硫磷|磷)|异氰酸(甲酯|苯酯)|异硫氰酸烯丙酯|乙酰(亚砷酸铜|替硫脲)|乙烯甲醇|乙酸(亚铊|铊|三乙基锡|三甲基锡|甲氧基乙基汞|汞)|乙硼烷|乙醇腈|乙撑亚胺|乙撑氯醇|伊皮恩|海洛因|一氧(化汞|化二氟)|一氯(乙醛|丙酮)|氧氯化磷|氧化(亚铊|铊|汞|二丁基锡)|烟碱|亚硝酰乙氧|亚硝酸乙酯|亚硒酸氢钠|亚硒酸钠|亚硒酸镁|亚硒酸二钠|亚硒酸|亚砷酸(钠|钾|酐)|冰毒|预测答案|考前预测|押题|代写论文|(提供|司考|级|传送|考中|短信)答案|(待|代|带|替|助)考|(包|顺利|保)过|考后付款|无线耳机|考试作弊|考前密卷|漏题|中特|一肖|报码|(合|香港)彩|彩宝|3D轮盘|liuhecai|一码|(皇家|俄罗斯)轮盘|赌具|特码|盗(号|qq|密码)|盗取(密码|qq)|嗑药|帮招人|社会混|拜大哥|电警棒|帮人怀孕|征兵计划|切腹|VE视觉|电鸡|仿真手枪|做炸弹|ONS|走私|陪聊|h(图|漫|网)|开苞|找(男|女)|口淫|卖身|元一夜|(男|女)奴|双(筒|桶)|看JJ|做台|厕奴|骚女|嫩逼|一夜激情|乱伦|泡友|富(姐|婆)|(足|群|茹)交|阴户|性(服务|伴侣|伙伴|交)|有偿(捐献|服务)|(有|无)码|包养|(犬|兽|幼)交|根浴|援交|小口径|性(虐|爱|息)|刻章|摇头丸|监听王|昏药|侦探设备|性奴|透视眼(睛|镜)|拍肩神|(失忆|催情|迷(幻|昏|奸)?|安定)(药|片|香)|游戏机破解|隐形耳机|银行卡复制设备|一卡多号|信用卡套现|消防[灭火]?枪|香港生子|土炮|胎盘|手机魔卡|容弹量|枪模|铅弹|汽(枪|狗|走表器)|气枪|气狗|伟哥|纽扣摄像机|免电灯|卖QQ号码|麻醉药|康生丹|警徽|记号扑克|激光(汽|气)|红床|狗友|反雷达测速|短信投票业务|电子狗导航手机|弹(种|夹)|(追|讨)债|车用电子狗|避孕|办理(证件|文凭)|斑蝥|暗访包|SIM卡复制器|BB(枪|弹)|雷管|弓弩|(电|长)狗|导爆索|爆炸物|爆破|左棍|婊子|换妻|成人片|淫(靡|水|兽)|阴(毛|蒂|道|唇)|小穴|缩阴|少妇自拍|(三级|色情|激情|黄色|小)(片|电影|视频|交友|电话)|肉棒|(情|奸)杀|裸照|乱伦|口交|禁(网|片)|春宫图|SM用品|自动群发|私家侦探服务|生意宝|商务(快车|短信)|慧聪|供应发票|发票代开|短信群发|短信猫|点金商务|士的宁|士的年|六合采|乐透码|彩票|百乐二呓|百家乐|黄页|出租|求购|留学咨询|外挂|淘宝|群发|货到付款|汽车配件|推广联盟|劳务派遣|网络(兼职|赚钱)|(证件|婚庆|翻译|搬家|追债|债务)公司|手机(游戏|窃听|监听|铃声|图片)|三唑仑|奇迹世界|工作服|wow|论文|铃声|彩(信|铃|票)|显示屏|投影仪|虚拟主机|(域名|专业)注册|营销|服务器托管|网站建设|(google|百度)排名|数据恢复|医院|性病|不孕不育|乳腺病|尖锐湿疣|皮肤病|减肥|瘦|3P|人兽|sex|代孕|打炮|找小姐|刻章|乱伦|中出|楼凤|卖淫|荡妇|群交|幼女|18禁|伦理电影|(催情|蒙汗|蒙汉|春)药|情趣用品|成人.+?(电影|用品)|激情(视频|电影|影院)|爽片|性感美女|交友|怀孕|裸聊|制服诱惑|丝袜|长腿|寂寞女子|免费电影|双色球|福彩|体彩|6合彩|时时彩|双色球|咨询热线|股票|荐股|开股|私服|SF|枪|警棒|警服|麻醉|诚招加盟|诚信经营|杀手|(游戏|金)币|群发|注册.+?公司|公司注册|发票|代开|淘宝|返利|团购|培训|折扣|(打包|试验|打标|破碎|灌装|升降)机|条码|标签纸|升降平台|二手(车|电脑)"
		Totoro_Config.Write "TOTORO_NUMBER_VALUE",10
		Totoro_Config.Write "TOTORO_REPLACE_KEYWORD","**"
		Totoro_Config.Write "TOTORO_REPLACE_LIST",""
		Totoro_Config.Write "TOTORO_CHINESESV",50
		Totoro_Config.Write "TOTORO_KILLIP",3
		Totoro_Config.Write "TOTORO_FILTERIP",""
		Totoro_Config.Write "TOTORO_TRANTOSIMP",True
        Totoro_Config.Write "TOTORO_CHECKSTR","Totoro大显神威！你的评论被怀疑是垃圾评论已经被提交审核。"
        Totoro_Config.Write "TOTORO_THROWSTR","Totoro大显神威！你的评论被怀疑是垃圾评论已经被删除。"
        Totoro_Config.Write "TOTORO_KILLIPSTR","Totoro大显神威！你的IP不合法不允许评论。"
		Totoro_Config.Write "TOTORO_PM",False
		Totoro_Config.Write "TOTORO_THROWCOUNT",0
		Totoro_Config.Write "TOTORO_CHECKCOUNT",0
		Totoro_Config.Save
		'Call SetBlogHint_Custom("您是第一次安装Totoro，已经为您导入初始配置。")
	ElseIf Totoro_Config.Read("TOTORO_VERSION")="0.0" Then
		Totoro_Config.Write "TOTORO_VERSION","3.0.3"
		Totoro_Config.Write "TOTORO_KILLIP",3
		Totoro_Config.Write "TOTORO_FILTERIP",""
		Totoro_Config.Write "TOTORO_TRANTOSIMP",False
        Totoro_Config.Write "TOTORO_CHECKSTR","Totoro大显神威！你的评论被怀疑是垃圾评论已经被提交审核。"
        Totoro_Config.Write "TOTORO_THROWSTR","Totoro大显神威！你的评论被怀疑是垃圾评论已经被删除。"
        Totoro_Config.Write "TOTORO_KILLIPSTR","Totoro大显神威！你的IP不合法不允许评论。"
		Totoro_Config.Write "TOTORO_THROWCOUNT",0
		Totoro_Config.Write "TOTORO_CHECKCOUNT",0
		Totoro_Config.Save
	ElseIf Totoro_Config.Read("TOTORO_VERSION")="3.0.3" Then
		Totoro_Config.Write "TOTORO_VERSION","3.0.4"
		Totoro_Config.Write "TOTORO_PM",False
		Totoro_Config.Write "TOTORO_THROWCOUNT",0
		Totoro_Config.Write "TOTORO_CHECKCOUNT",0
		Totoro_Config.Save
	ElseIf Totoro_Config.Read("TOTORO_VERSION")="3.0.4" Then
		Totoro_Config.Write "TOTORO_VERSION","3.0.5"
		Totoro_Config.Write "TOTORO_CHECKCOUNT",0
		Totoro_Config.Save
	End If
End Function

Function Totoro_Initialize()
	'On Error Resume Next
	InstallPlugin_Totoro
	TOTORO_INTERVAL_VALUE=CLng(Totoro_Config.Read ("TOTORO_INTERVAL_VALUE"))
	If Err.Number<>0 Then Totoro_Config.Remove("TOTORO_VERSION"):Totoro_Config.Save:Call InstallPlugin_Totoro:TOTORO_INTERVAL_VALUE=CLng(Totoro_Config.Read ("TOTORO_INTERVAL_VALUE")):Call SetBlogHint_Custom("Totoro配置出错，已经重新初始化！")
	TOTORO_BADWORD_VALUE=CLng(Totoro_Config.Read ("TOTORO_BADWORD_VALUE"))
	TOTORO_HYPERLINK_VALUE=CLng(Totoro_Config.Read ("TOTORO_HYPERLINK_VALUE"))
	TOTORO_NAME_VALUE=CLng(Totoro_Config.Read ("TOTORO_NAME_VALUE"))
	TOTORO_LEVEL_VALUE=CLng(Totoro_Config.Read ("TOTORO_LEVEL_VALUE"))
	TOTORO_SV_THRESHOLD=CLng(Totoro_Config.Read ("TOTORO_SV_THRESHOLD"))
	TOTORO_SV_THRESHOLD2=CLng(Totoro_Config.Read ("TOTORO_SV_THRESHOLD2"))
	TOTORO_DEL_DIRECTLY=CBool(Totoro_Config.Read ("TOTORO_DEL_DIRECTLY"))
	TOTORO_CONHUOXINGWEN=CBool(Totoro_Config.Read ("TOTORO_ConHuoxingwen"))
	TOTORO_BADWORD_LIST=Totoro_Config.Read ("TOTORO_BADWORD_LIST")
	TOTORO_NUMBER_VALUE=CLng(Totoro_Config.Read ("TOTORO_NUMBER_VALUE"))
	TOTORO_REPLACE_KEYWORD=Totoro_Config.Read ("TOTORO_REPLACE_KEYWORD")
	TOTORO_REPLACE_LIST=Totoro_Config.Read ("TOTORO_REPLACE_LIST")
	TOTORO_CHINESESV=Totoro_Config.Read("TOTORO_CHINESESV")
	TOTORO_KILLIP=CLng(Totoro_Config.Read("TOTORO_KILLIP"))
	TOTORO_FILTERIP=Totoro_Config.Read("TOTORO_FILTERIP")
	TOTORO_TRANTOSIMP=CBool(Totoro_Config.Read("TOTORO_TRANTOSIMP"))
    TOTORO_CHECKSTR=Totoro_Config.Read ("TOTORO_CHECKSTR")
    TOTORO_THROWSTR=Totoro_Config.Read ("TOTORO_THROWSTR")
    TOTORO_KILLIPSTR=Totoro_Config.Read ("TOTORO_KILLIPSTR")
	TOTORO_PM=CBool(Totoro_Config.Read("TOTORO_PM"))
	TOTORO_THROWCOUNT=CLng(Totoro_Config.Read("TOTORO_THROWCOUNT"))
	TOTORO_CHECKCOUNT=CLng(Totoro_Config.Read("TOTORO_CHECKCOUNT"))
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
	objComment.IP=GetReallyIP()
	objComment.Agent=Request.ServerVariables("HTTP_USER_AGENT")
	Call Totoro_cComment(objComment,BlogUser,False)

	Totoro_chkComment=True

End Function
'*********************************************************

Sub Totoro_cComment(objComment,objUser,isDebug)
	Totoro_Debug=isDebug
	
	If objComment.IsCheck=True Then Exit Sub
	If objComment.IsThrow=True Then Exit Sub
	
	If Totoro_FunctionFilterIP(objComment.IP) Then
		ZVA_ErrorMsg(14)=TOTORO_KILLIPSTR
		Totoro_AddDebug ZVA_ErrorMsg(14)
		objComment.IsThrow=True
		Exit Sub
	End iF
	Totoro_AddDebug "IP("&objcomment.ip&")不在范围内，进入下一步测试"
	
	Dim strTemp
	strTemp=objComment.Content
	If Totoro_PM Then
		strTemp=Totoro_FilterPMPlusHtmlTag(strTemp)
	End If
	If TOTORO_TRANTOSIMP Then
		strTemp=Totoro_FunctionTranToSimp(strTemp)
	End If
	If TOTORO_ConHuoxingwen Then
		
		strTemp=Totoro_Xiou(strTemp)
		strTemp=Totoro_FxxxHuoxingwen(strTemp)
		strTemp=Totoro_FromSBCCode(strTemp)
		strTemp=Totoro_GetNum(strTemp)		
		
	End If
		

	Totoro_AddDebug "待处理评论：" & vbcrlf & strTemp
	Call Totoro_checkLevel(objUser.Level)
	Totoro_AddDebug "用户级别测试完毕。SV为：" & Totoro_SV
	Call Totoro_checkName(objComment.IP)
	Totoro_AddDebug "访客熟悉度测试完毕。SV为：" & Totoro_SV
	Call Totoro_checkHyperLink(strTemp)
	Totoro_AddDebug "超链接测试完毕。SV为：" & Totoro_SV
	Call Totoro_checkBadWord(strTemp & "&" & objComment.Author & "&" & objComment.HomePage & "&" & objComment.IP & "&" & objComment.Email)
	Totoro_AddDebug "黑词测试完毕。SV为：" & Totoro_SV
	Call Totoro_checkInterval(GetReallyIP())
	Totoro_AddDebug "发表频率测试完毕。SV为：" & Totoro_SV
	Call Totoro_checkNumLong(strTemp)
	Totoro_AddDebug "数字长度测试完毕。SV为：" & Totoro_SV
	Call Totoro_checkChinese(strTemp)
	Totoro_AddDebug "中文测试完毕。SV为：" & Totoro_SV
	objComment.Content=Totoro_replaceWord(objComment.Content)
	objComment.Author=Totoro_replaceWord(objComment.Author)
	'Response.AddHeader "Totoro_SV",Totoro_SV
	'Response.AddHeader "Content",strTemp
	Totoro_AddDebug "敏感词替换完毕"
	Dim o
	
	If Totoro_SV>=TOTORO_SV_THRESHOLD Then
		ZVA_ErrorMsg(14)=TOTORO_THROWSTR
		ZVA_ErrorMsg(53)=TOTORO_CHECKSTR
		
		If Totoro_SV<TOTORO_SV_THRESHOLD2 Or TOTORO_SV_THRESHOLD2=0 Then
			objComment.IsCheck=True
			Totoro_AddDebug "该评论进入审核列表"
			If isDebug=False Then
				TOTORO_CHECKCOUNT=TOTORO_CHECKCOUNT+1
				Totoro_Config.Write "TOTORO_CHECKCOUNT",TOTORO_CHECKCOUNT
				Totoro_Config.Save
				o=Totoro_FunctionKillIP(objComment,False)
			End If
		ElseIf TOTORO_SV_THRESHOLD2<=Totoro_SV Then
			If isDebug=False Then 
				TOTORO_THROWCOUNT=TOTORO_THROWCOUNT+1
				Totoro_Config.Write "TOTORO_THROWCOUNT",TOTORO_THROWCOUNT
				Totoro_Config.Save
				objComment.IsThrow=True
				o=Totoro_FunctionKillIP(objComment,True)
			Else
				Totoro_AddDebug "该评论被拦截"
			End If
		
		End If
	Else
		Totoro_AddDebug "判断为正常评论"
		Totoro_AddDebug "最终输出评论：" &vbcrlf & objComment.Content
		Totoro_AddDebug "最终用户名：" & vbcrlf & objComment.Author
	End If
End Sub

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

Function Totoro_FilterPMPlusHtmlTag(str)
	Dim a,strT
	strT=TransferHTML(str,"[nohtml]")
	a="~!@#$%^&*()_+|-=\{}[];':""<>?/.,！＃￥…（）—、【】｛｝；：‘’“”《》，。、？"&Chr(9)
	Dim i
	For i=1 To Len(a)
		strT=Replace(strT,Mid(a,i,1),"")
	Next
	Totoro_FilterPMPlusHtmlTag=strT
End Function

Function Totoro_checkName(ByVal ip)

	If TOTORO_NAME_VALUE=0 Then Exit Function

	Dim i,s
	s=FilterSQL(ip)
	i=0
	Dim objRS
	Set objRS=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [log_ID]>=0 and [comm_IP] ='" & ip & "' and [comm_isCheck]=0")
	If (Not objRS.bof) And (Not objRS.eof) Then
		i=objRS(0)
	End If
	Set objRS=Nothing

	If i=0 Then Totoro_SV=Totoro_SV
	If i>0 And i<=10   Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(0)
	If i>10 And i<=20  Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(1)
	If i>20 And i<=50 Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(2)
	If i>50           Then Totoro_SV=Totoro_SV-10-TOTORO_NAME_VALUE*(3)

End Function


Function Totoro_checkBadWord(ByVal content)
	On Error Resume Next
	If Totoro_SV+TOTORO_BADWORD_VALUE=0 Then Exit Function
	Dim o
	Set o=New RegExp
	o.Pattern=vbsunescape(TOTORO_BADWORD_LIST)
	o.Global=True
	o.IgnoreCase=True
	dim matches
	set matches=o.execute(content)
	If Err.Number=0 Then
		Totoro_SV=Totoro_SV+TOTORO_BADWORD_VALUE*matches.count
	Else
		Totoro_AddDebug "黑词列表错误"
	End If
	Set o=Nothing
End Function

Function Totoro_replaceWord(content)
	On Error Resume Next
	Dim o
	If TOTORO_REPLACE_LIST="" Then Totoro_replaceWord=content:Exit Function
	Set o=New RegExp
	o.Pattern=TOTORO_REPLACE_LIST
	o.Global=True
	o.IgnoreCase=True
	If Err.Number=0 Then
		Totoro_replaceWord=o.replace(content,TOTORO_REPLACE_KEYWORD)
	Else
		Totoro_AddDebug "敏感词列表错误"
	End If
	Set o=Nothing
End Function



Function Totoro_checkHyperLink(ByVal content)

	If TOTORO_HYPERLINK_VALUE=0 Then Exit Function

	Dim SRegExp,Matches
	Set SRegExp=New RegExp
	SRegExp.IgnoreCase =True
	SRegExp.Global=True
	SRegExp.Pattern="https?://(?!www|ftp)|ftp|www."
	Set Matches = SRegExp.Execute(content)

	Totoro_SV=Totoro_SV+TOTORO_HYPERLINK_VALUE*(2^matches.count-1)

	Set SRegExp=Nothing

End Function


Function Totoro_checkInterval(ByVal ip)
	On Error Resume Next
	If TOTORO_INTERVAL_VALUE=0 Then Exit Function
	Dim i,j,t,s,m,n
	Dim objRS
	i=0
	m="SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_IP] ='" & ip & "'"
	
	m=m&" AND [comm_PostTime]>"&ZC_SQL_POUND_KEY& FormatDateTime( DateAdd("h", -1, now) ) &ZC_SQL_POUND_KEY
	Set objRS=objConn.Execute(m)
	If Err.Number=0 Then
		If (Not objRS.bof) And (Not objRS.eof) Then
			i=objRS(0)
		End If
	Else
		i=0
		Totoro_AddDebug "时间格式可能有误"
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
	Set objRS1=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_isCheck]=-1 Or [comm_isCheck]=1")
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
	a=Array("҉|","蕶","ニ|貳","弎","陸","ハ|仈","艽","ā|á|ǎ|à|а|А|α","в|в|В|ъ|Ъ|ы|Ы|ь|Ь|β","с|с|С","Ё|е|Е|ё|Ё|ê|ē|é|ě|è","℉|ｆ","ɡ","н|Н","ī|í|ǐ|ì","ｊ","κ","ι","м|М","ń|п|П|Й|π","ō|ó|ǒ|ǒ|о|О|ο|σ|⊙|○|◎","р|Р|ρ","я|Я","\$","т|Т|τ","ū|ú|ǔ|ù|∪|μ|υ","∨|ν","ω","×|х|Х|χ","у|У|γ","э|Э","θ","ф|Ф")
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
	a=Array("零|〇"," 一|壹|Ⅰ|⒈|㈠|①|⑴","二|贰|Ⅱ|⒉|㈡|②|⑵","三|叁|Ⅲ|⒊|㈢|③|⑶","四|肆|Ⅳ|⒋|㈣|④|⑷","五|伍|Ⅴ|⒌|⑤|㈤|⑸","六|陆|Ⅵ|⒍|㈥|⑥|⑹","七|柒|Ⅶ|⒎|⑦|㈦|⑺","八|捌|Ⅷ|⒏|㈧|⑧|⑻","九|玖|Ⅸ|⒐|⑨|㈨|⑼","十|拾|㈩|⑩|⒑|Ⅺ|⑽")
	b=Array(0,1,2,3,4,5,6,7,8,9,10)
	Dim c,i
	set c=new regexp
	c.Global=True
	c.IgnoreCase=True
	For i=0 To 10
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

Function Totoro_FunctionFilterIP(userip)
'COPY FROM ANTISPAM (HTTP://WWW.WILLIAMLONG.INFO)
	Dim IPlock
	Dim locklist
	Dim i, StrUserIP, StrKillIP
	IPlock = False
	locklist = Trim(TOTORO_FILTERIP)
	If locklist = "" Then Exit Function
	StrUserIP = userip
	locklist = Split(locklist, "|")
	If StrUserIP = "" Then Exit Function
	StrUserIP = Split(userip, ".")
	If UBound(StrUserIP) <> 3 Then Exit Function
	For i = 0 To UBound(locklist)
		locklist(i) = Trim(locklist(i))
	    If locklist(i) <> "" Then
			StrKillIP = Split(locklist(i), ".")
			If UBound(StrKillIP) <> 3 Then Exit For
			IPlock = True
			If (StrUserIP(0) <> StrKillIP(0)) And InStr(StrKillIP(0), "*") = 0 Then IPlock = False
			If (StrUserIP(1) <> StrKillIP(1)) And InStr(StrKillIP(1), "*") = 0 Then IPlock = False
			If (StrUserIP(2) <> StrKillIP(2)) And InStr(StrKillIP(2), "*") = 0 Then IPlock = False
			If (StrUserIP(3) <> StrKillIP(3)) And InStr(StrKillIP(3), "*") = 0 Then IPlock = False
			If IPlock Then Exit For
	    End If
	Next
	Totoro_FunctionFilterIP = IPlock	
End Function
	
Function Totoro_FunctionKillIP(obj,ist)
	If TOTORO_KILLIP=0 Then Exit Function
	Dim objRs,strSQL,strSQL2
	If ZC_MSSQL_ENABLE Then
		strSQL2=" [comm_PostTime]>'"&DateAdd("d",-1,now)&"'"
	Else
		strSQL2=" [comm_PostTime]>#"&DateAdd("d",-1,now)&"#"
	End If
	strSQL="SELECT COUNT ([comm_ID]) FROM [blog_Comment] WHERE [comm_IP]='"&obj.IP&"' AND"&strSQL2
	Set objRs=objConn.Execute(strSQL)
	Dim j
	If Not objRs.Eof Then
		j=objRs(0)
	End If
	If j>TOTORO_KILLIP Or ist=True Then
			If ist=False Then
				TOTORO_FILTERIP=IIf(TOTORO_FILTERIP="",obj.ip,TOTORO_FILTERIP&"|"&obj.ip)
				Totoro_Config.Write "TOTORO_FILTERIP",TOTORO_FILTERIP
				Totoro_Config.Save
			End If
			Call Totoro_DelSpam(obj.IP,ist)
	End If
	Totoro_FunctionKillIP=j
End Function
	
Function Totoro_DelSpam(IP,isTh)
	Dim objRs,strSQL,strSQL2
	If isTh=False Then
		If ZC_MSSQL_ENABLE Then
			strSQL2=" AND [comm_PostTime]>'"&DateAdd("d",-1,now)&"'"
		Else
			strSQL2=" AND [comm_PostTime]>#"&DateAdd("d",-1,now)&"#"
		End If
	End If
	strSQL="UPDATE [blog_Comment] SET [comm_isCheck]=1 WHERE [comm_IP]='"&IP&"'"&strSQL2
	Set objRs=objConn.Execute(strSQL)
	strSQL="SELECT [log_ID] FROM [blog_Comment] WHERE [comm_IP]='"&IP&"'"&strSQL2
	Set objRs=objConn.Execute(strSQL)
	Do Until objRs.Eof
		Call BuildArticle(objRs("log_ID"),False,True)
		objRs.MoveNext
	Loop
	BlogReBuild_Comments
	Call ClearGlobeCache
	Call LoadGlobeCache
End Function



%>