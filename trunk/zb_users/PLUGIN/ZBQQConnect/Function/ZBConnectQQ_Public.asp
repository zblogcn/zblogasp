
<%
'**************************************
' 西欧字符转换
'**************************************
Function ZBQQConnect_ReplaceXO(ByVal strContent)
	Dim a,b,c,d
	a=Split("ltZSXgtZSXampZSXquotZSXcopyZSXregZSXtimesZSXdivideZSXnbspZSXyenZSXordfZSXmacrZSXacuteZSXsup1ZSXfrac34ZSXatildeZSXegraveZSXiacuteZSXograveZSXtimesZSXuumlZSXaacuteZSXaeligZSXeumlZSXethZSXotildeZSXuacuteZSXyumlZSXiexclZSXbrvbarZSXlaquoZSXdegZSXmicroZSXordmZSXiquestZSXaumlZSXeacuteZSXicircZSXoacuteZSXoslashZSXyacuteZSXacircZSXccedilZSXigraveZSXntildeZSXoumlZSXucircZSXcurrenZSXcopyZSXregZSXsup3ZSXcedilZSXfrac12ZSXacircZSXccedilZSXigraveZSXntildeZSXoumlZSXucircZSXagraveZSXaringZSXecircZSXiumlZSXocircZSXugraveZSXthornZSXcentZSXsectZSXnotZSXplusmnZSXparaZSXraquoZSXagraveZSXaringZSXecircZSXiumlZSXocircZSXugraveZSXthornZSXatildeZSXegraveZSXiacuteZSXograveZSXdivideZSXuumlZSXpoundZSXumlZSXshyZSXsup2ZSXmiddotZSXfrac14ZSXaacuteZSXaeligZSXeumlZSXethZSXotildeZSXuacuteZSXszligZSXaumlZSXeacuteZSXicircZSXoacuteZSXoslashZSXyacuteZSXoeligZSXoeligZSXtildeZSXzwjZSXlsquoZSXbdquoZSXrsaquoZSXscaronZSXlrmZSXrsquoZSXdaggerZSXeuroZSXscaronZSXrlmZSXsbquoZSXdaggerZSXyumlZSXthinspZSXndashZSXldquoZSXpermilZSXcircZSXzwnjZSXmdashZSXrdquoZSXlsaquoZSXfnofZSXepsilonZSXkappaZSXomicronZSXupsilonZSXalphaZSXzetaZSXlambdaZSXpiZSXupsilonZSXthetasymZSXprimeZSXimageZSXuarrZSXlarrZSXforallZSXisinZSXminusZSXangZSXintZSXneZSXsupZSXotimesZSXlfloorZSXspadesZSXalphaZSXzetaZSXlambdaZSXpiZSXphiZSXbetaZSXetaZSXmuZSXrhoZSXphiZSXupsihZSXprimeZSXrealZSXrarrZSXuarrZSXpartZSXnotinZSXlowastZSXandZSXthere4ZSXequivZSXnsubZSXperpZSXrfloorZSXclubsZSXbetaZSXetaZSXmuZSXrhoZSXchiZSXgammaZSXthetaZSXnuZSXsigmafZSXchiZSXpivZSXolineZSXtradeZSXdarrZSXrarrZSXexistZSXniZSXradicZSXorZSXsimZSXleZSXsubeZSXsdotZSXlangZSXheartsZSXgammaZSXthetaZSXnuZSXsigmaZSXpsiZSXdeltaZSXiotaZSXxiZSXsigmaZSXpsiZSXbullZSXfraslZSXalefsymZSXharrZSXdarrZSXemptyZSXprodZSXpropZSXcapZSXcongZSXgeZSXsupeZSXlceilZSXrangZSXdiamsZSXdeltaZSXiotaZSXxiZSXtauZSXomegaZSXepsilonZSXkappaZSXomicronZSXtauZSXomegaZSXhellipZSXweierpZSXlarrZSXcrarrZSXharrZSXnablaZSXsumZSXinfinZSXcupZSXasympZSXsubZSXoplusZSXrceilZSXloz","ZSX")
	b=Split("<SOFT>SOFT&SOFT""SOFT©SOFT®SOFT×SOFT÷SOFT"&vbcrlf&"SOFT¥SOFTªSOFT¯SOFT´SOFT¹SOFT¾SOFTÃSOFTÈSOFTÍSOFTÒSOFT×SOFTÜSOFTáSOFTæSOFTëSOFTðSOFTõSOFTúSOFTÿSOFT¡SOFT¦SOFT«SOFT°SOFTµSOFTºSOFT¿SOFTÄSOFTÉSOFTÎSOFTÓSOFTØSOFTÝSOFTâSOFTçSOFTìSOFTñSOFTöSOFTûSOFT¤SOFT©SOFT®SOFT³SOFT¸SOFT½SOFTÂSOFTÇSOFTÌSOFTÑSOFTÖSOFTÛSOFTàSOFTåSOFTêSOFTïSOFTôSOFTùSOFTþSOFT¢SOFT§SOFT¬SOFT±SOFT¶SOFT»SOFTÀSOFTÅSOFTÊSOFTÏSOFTÔSOFTÙSOFTÞSOFTãSOFTèSOFTíSOFTòSOFT÷SOFTüSOFT£SOFT¨SOFT"&vbcrlf&"SOFT²SOFT·SOFT¼SOFTÁSOFTÆSOFTËSOFTÐSOFTÕSOFTÚSOFTßSOFTäSOFTéSOFTîSOFTóSOFTøSOFTýSOFTŒSOFTœSOFT˜SOFT‍SOFT‘SOFT„SOFT›SOFTŠSOFT‎SOFT’SOFT†SOFT€SOFTšSOFT‏SOFT‚SOFT‡SOFTŸSOFT SOFT–SOFT“SOFT‰SOFTˆSOFT SOFT—SOFT”SOFT‹SOFTƒSOFTΕSOFTΚSOFTΟSOFTΥSOFTαSOFTζSOFTλSOFTπSOFTυSOFT?SOFT′SOFTℑSOFT↑SOFT⇐SOFT∀SOFT∈SOFT−SOFT∠SOFT∫SOFT≠SOFT⊃SOFT⊗SOFT?SOFT♠SOFTΑSOFTΖSOFTΛSOFTΠSOFTΦSOFTβSOFTηSOFTμSOFTρSOFTφSOFT?SOFT″SOFTℜSOFT→SOFT⇑SOFT∂SOFT∉SOFT∗SOFT∧SOFT∴SOFT≡SOFT⊄SOFT⊥SOFT?SOFT♣SOFTΒSOFTΗSOFTΜSOFTΡSOFTΧSOFTγSOFTθSOFTνSOFTςSOFTχSOFT?SOFT‾SOFT™SOFT↓SOFT⇒SOFT∃SOFT∋SOFT√SOFT∨SOFT∼SOFT≤SOFT⊆SOFT⋅SOFT?SOFT♥SOFTΓSOFTΘSOFTΝSOFTΣSOFTΨSOFTδSOFTιSOFTξSOFTσSOFTψSOFT•SOFT⁄SOFTℵSOFT↔SOFT⇓SOFT∅SOFT∏SOFT∝SOFT∩SOFT∝SOFT≥SOFT⊇SOFT?SOFT?SOFT♦SOFTΔSOFTΙSOFTΞSOFTΤSOFTΩSOFTεSOFTκSOFTοSOFTτSOFTωSOFT…SOFT℘SOFT←SOFT↵SOFT⇔SOFT∇SOFT∑SOFT∞SOFT∪SOFT≈SOFT⊂SOFT⊕SOFT?SOFT◊","SOFT")
	For c=0 To Ubound(a)
		strContent=Replace(strContent,"&"&a(c)&";",b(c))
	Next
	a=""
	Set a=New RegExp
	a.Pattern="&#(\d+?);"
	a.Global=True
	Set b=a.Execute(strContent)
	For Each c In b
		d = CLng(c.Submatches(0))
		If d - 65536 > 0 Then
			d = d - 65536
		End If
		strContent = Replace(strContent, c.value, ChrW(d))
	Next
    Set b=Nothing
	Set a=Nothing
	
	ZBQQConnect_ReplaceXO=strContent
End Function
'**************************************
' 输出顶部栏
'**************************************

Function ZBQQConnect_SBar(Btype)
	dim b(4,3),i,j,k
	b(1,1)="m-left"
	b(1,2)="main.asp"
	b(1,3)="首页"
	If BlogUser.Level<5 Then
		b(2,1)="m-left"
		b(2,2)="m.asp"
		b(2,3)="绑定管理"
		If BlogUser.Level=1 Then
			b(3,1)="m-left"
			b(3,2)="setting.asp"
			b(3,3)="插件配置"
		End If
		b(4,1)="m-left"
		b(4,2)="usersetting.asp"
		b(4,3)="用户配置"
	End If
	For i=1 to 4
		if b(i,1)<>"" then
			if btype=i then
				k=k&"<a href=""" & b(i,2) & """><span class=""" & b(i,1) & " m-now"">" & b(i,3) & "</span></a>"
			else
				k=k&"<a href=""" & b(i,2) & """><span class=""" & b(i,1) & """>" & b(i,3) & "</span></a>"
			end if
		end if
	Next
	k=k&"<script type=""text/javascript"">ActiveLeftMenu(""aQQConnect"");</script>"
	ZBQQConnect_SBar=k
End Function
'******************************
'检查配置并初始化程序
'******************************
Function ZBQQConnect_Initialize
	Set ZBQQConnect_Config=New TConfig
	ZBQQConnect_Config.Load "ZBQQConnect"
	If ZBQQConnect_Config.Exists("-。-")=False Then ZBQQConnect_First
	ZBQQConnect_notfoundpic="~"
	ZBQQConnect_strLong=30
	ZBQQConnect_DefaultToZone=CBool(ZBQQConnect_Config.Read("a"))
	ZBQQConnect_DefaultTot=CBool(ZBQQConnect_Config.Read("b"))
	ZBQQConnect_PicSendToWb=CBool(ZBQQConnect_Config.Read("c"))
	ZBQQConnect_OpenComment=CBool(ZBQQConnect_Config.Read("d"))
	ZBQQConnect_CommentToZone=CBool(ZBQQConnect_Config.Read("e"))
	ZBQQConnect_CommentToT=CBool(ZBQQConnect_Config.Read("f"))
	ZBQQConnect_CommentToOwner=CBool(ZBQQConnect_Config.Read("g"))
	ZBQQConnect_allowQQLogin=CBool(ZBQQConnect_Config.Read("h"))
	ZBQQConnect_allowQQReg=CBool(ZBQQConnect_Config.Read("i"))
	ZBQQConnect_HeadMode=CInt(ZBQQConnect_Config.Read("a1"))
	'ZBQQConnect_Head=ZBQQConnect_Config.Read("Gravatar")
	ZBQQConnect_Content=ZBQQConnect_Config.Read("content")
	ZBQQConnect_WBKey=ZBQQConnect_Config.Read("WBKEY")
	ZBQQConnect_WBSecret=ZBQQConnect_Config.Read("WBAPPSecret")
	ZBQQConnect_CommentTemplate=ZBQQConnect_Config.Read("pl")
	Set ZBQQConnect_Net=New ZBQQConnect_NetWork
	set ZBQQConnect_class=new ZBQQConnect
	Set ZBQQConnect_DB=New ZBConnectQQ_DB
	ZBQQConnect_class.app_key=ZBQQConnect_Config.Read("AppID")    '设置appkey
	ZBQQConnect_class.app_secret=ZBQQConnect_Config.Read("KEY")  '设置app_secret
	ZBQQConnect_class.callbackurl=GetCurrentHost&"ZB_USERS/PLUGIN/ZBQQConnect/callback.asp"  '设置回调地址
	ZBQQConnect_class.fakeQQConnect.app_key=ZBQQConnect_WBKey
	ZBQQConnect_class.fakeQQConnect.app_secret=ZBQQConnect_WBSecret
	ZBQQConnect_class.fakeQQConnect.Token=ZBQQConnect_Config.Read("WBToken")
	ZBQQConnect_class.fakeQQConnect.Secret=ZBQQConnect_Config.Read("WBSecret")
	ZBQQConnect_class.fakeQQConnect.UserID=ZBQQConnect_Config.Read("WBName")
End Function
'******************************
'第一次使用导入配置
'******************************
Sub ZBQQConnect_First()
		dim i
		ZBQQConnect_Config.Write "-。-","1.0"
		For i=97 To 105
			ZBQQConnect_Config.Write Chr(i),iIf(chr(i)<>"g",True,False)
		Next
		ZBQQConnect_Config.Write "a1","0"
		'ZBQQConnect_Config.Write "Gravatar","http://www.gravatar.com/avatar/<#EmailMD5#>?s=40&d=<#ZC_BLOG_HOST#>%2FZB%5FSYSTEM%2Fimage%2Fadmin%2Favatar%2Epng"
		ZBQQConnect_Config.Write "content","更新了文章：《%t》，%u"
		ZBQQConnect_Config.Write "WBKEY","2e21c7b056f341b080d4d3691f3d50fb"
		ZBQQConnect_Config.Write "WBAPPSecret","1b84a3016c132a6839d082605b854bbe"
		ZBQQConnect_Config.Write "pl","@%a 评论 %c"
		ZBQQConnect_Config.Save
End Sub
'******************************
'绑定现有帐号
'******************************
Sub ZBQQConnect_RegSave(UID)
	If Not IsEmpty(Request.Form("QQOpenID")) Then
		ZBQQConnect_Initialize
		ZBQQConnect_DB.OpenID=Request.Form("QQOpenID")
		If ZBQQConnect_DB.LoadInfo(4)=True Then
			ZBQQConnect_DB.objUser.LoadInfoById UID
			ZBQQConnect_DB.Email=ZBQQConnect_DB.objUser.Email
			ZBQQConnect_DB.Bind
		End If
	End If
End Sub
'******************************
'  得到文章内第一张图片
'******************************
Function ZBQQConnect_LoadPicture(ByVal str)
	Dim objRegExp,Match,Matches,tmp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True
	objRegExp.Pattern="<img.*src\s*=\s*[\""|\']?\s*([^>\""\'\s]*)" 
	Set Matches=objRegExp.Execute(str)
	For Each Match in Matches 
		tmp=objRegExp.Replace(Match.Value,"$1") 
		Exit For
	Next
	set objregexp=nothing
	If Instr(tmp,"http")<0 And tmp<>"" Then tmp=GetCurrentHost & "/" & tmp
	ZBQQConnect_LoadPicture=tmp
End Function
%>