<%
''*****************************************************
'   ZSXSOFT ASP QQConnect类
''*****************************************************
Class ZBQQConnect

'*******************************************************************************
'** 定义变量                                                                    **
'*******************************************************************************
Private strOauthToken,strPostUrl,strOauthSessionUserID,strAccToken,strOpenID,intErrorCount,strOauthTokenSecret
Private strMadeUpUrl
Private strHttptype
Private strAppKey,strAppSecret
Private intRepeatMax,aryMutliContent
Private strContentWithReplaceEnter,strContentWithOutEncode,strOauthCallbackUrl,strWithOutOauthSignature
Private tempid,objJSON,bolDebugMsg
Private objXmlhttp,strUserAgent
Private strPrototype,strOauthVersion,strSessionClsID

'*******************************************************************************
'** 初始化                                                                **
'*******************************************************************************
Sub Class_Initialize()
	strSessionClsID=ZC_BLOG_CLSID
	strHttptype = "GET&"
	intErrorCount=0
	intRepeatMax=3
	set objJSON=ZBQQConnect_toobject("{}")
	
	strOpenID=Session(strSessionClsID&"ZBQQConnect_strOpenID")
	strAccToken=Session(strSessionClsID&"ZBQQConnect_strAccessToken")
	set objXmlhttp=server.CreateObject("msxml2.serverXmlhttp")
	version="2.0"
	strPrototype="http://"
	strUserAgent="ZBQQConnect ByZSXSOFT"
	'需要读取数据库等可在这里插入代码
End Sub
'*******************************************************************************
'** 回收资源                            
'*******************************************************************************
Sub Class_Terminate()
	set objXmlhttp=nothing
End Sub
'****************************************************************************
'** 设置CallBackUrl
'****************************************************************************
Public Property Let callbackurl(url)
		strOauthCallbackUrl=strUrlEnCode(url)
End Property
Public Property Get callbackurl
		callbackurl=strOauthCallbackUrl
End Property
'****************************************************************************
'** Debug模式
'****************************************************************************
Public Property Let debug(bool)
		bolDebugMsg=bool
End Property
Public Property Get debug
		debug=bolDebugMsg
End Property
'****************************************************************************
'** Oauth_Version设置
'****************************************************************************
Public Property Let Version(Str)
		strOauthVersion=str
		strPrototype="https://"
End Property
Public Property Get Version
		Version=strOauthVersion
End Property
'****************************************************************************
'** 设置AppKey
'****************************************************************************
Public Property Let app_key(Str)
		strAppKey=str
End Property
Public Property Get app_key
		app_key=strAppKey
End Property
'****************************************************************************
'** 设置AppSecret
'****************************************************************************
Public Property Let app_secret(Str)
		strAppSecret=str
End Property
Public Property Get app_secret
		app_secret=strAppSecret
End Property
'****************************************************************************
'** 设置User-Agent
'****************************************************************************
Public Property Let UserAgent(Str)
		strUserAgent=str
End Property
Public Property Get UserAgent
		UserAgent=strUserAgent
End Property
'****************************************************************************
'** 设置最大重试次数
'****************************************************************************
Public Property Let Repeat(inte)
		intErrorCount=inte
End Property
Public Property Get Repeat
		Repeat=intErrorCount
End Property
'****************************************************************************
'** 设置HttpType
'****************************************************************************
Public Property Let httptype(str)
		strHttptype=str
End Property
Public Property Get httptype
		httptype=strHttptype
End Property
'****************************************************************************
'** 设置唯一CLSID （防止Session冲突）
'****************************************************************************
Public Property Let ClsID(str)
		strSessionClsID=str
End Property
Public Property Get ClsID
		ClsID=strSessionClsID
End Property
'****************************************************************************
'** 关于
'****************************************************************************
Public Function About
	Set About=ZBQQConnect_Toobject("{'author':'ZSXSOFT','url':'http://www.zsxsoft.com','version':'ZBQQConnect 1.0'}")
End Function
'****************************************************************************
'** 得到是否登陆
'****************************************************************************
'Public Property Get  logined()
'		if strOpenID="" or  request.QueryString("typ")="logout" then
'			logined=false
'		else
'			logined=true
'		end if
'end Property
'****************************************************************************
'** 得到OpenID
'****************************************************************************
Public Property Let OpenID(str)
		strOpenID=str
End Property
Public Property Get OpenID
		OpenID=strOpenID
End Property
'****************************************************************************
'** 得到AccessToken
'****************************************************************************
Public Property Let AccessToken(str)
		strAccToken=str
End Property
Public Property Get AccessToken
		AccessToken=strAccToken
End Property

'****************************************************************************
'** 注销
'****************************************************************************
Public Sub logout()
	'Session(strSessionClsID&"ZBQQConnect_strAccessToken")=""
	'Session(strSessionClsID&"ZBQQConnect_strOpenID")=""
	''数据库代码在这里加入
	ZBQQConnect_DB.OpenID=strOpenID
	ZBQQConnect_DB.LoadInfo 4
	ZBQQConnect_DB.Del 
End Sub
'*******************************************************************************
'** 运行插件                                                                **
'*******************************************************************************
Function API(url,json,httptype)
	If Right(httptype,1)<>"&" Then httptype=httptype&"&"
	strHttpType=httptype
	if strHttptype="GET&" then
		API=gethttp(MakeOauthUrl(url,"sdk_custom",json))
	elseif strHttptype="POST&" then
		API=posthttp(MakeOauthUrl(url,"sdk_custom",json))
	end if
	If bolDebugMsg=true then response.write "<font color='darkyellow'>返回结果：" &run & "</font>"
	tempid=""
	aryMutliContent=""
	strHttptype = "GET&"
End Function


'*******************************************************************************
'** 组合Url                                                                    **
'*******************************************************************************
Function MakeOauthUrl(ByRef oauth_url,ip,content)
	MakeOauthUrl=MakeOauth2Url(oauth_url,ip,content)
End Function
'*******************************************************************************
'** 得到验证URL                                                         **
'*******************************************************************************
Function Authorize()
	Dim a,b,c
	a="https://graph.qq.com/oauth2.0/authorize"
	Set b=ZBQQConnect_toObject("{}")
	Call ZBQQConnect_addobj(b,"response_type","code")
	Call ZBQQConnect_addobj(b,"client_id",strAppKey)
	Call ZBQQConnect_addobj(b,"redirect_uri",strOauthCallBackUrl)
	Call ZBQQConnect_addobj(b,"scope","get_user_info,add_share,get_info,add_idol")
	Call ZBQQConnect_addobj(b,"state","zsxsoft")

	c=ZBQQConnect_toStr(b)
	Authorize=a&"?"&c'GetHttp(a&"?"&c)
End Function
'*******************************************************************************
'** CallBack                                                         **
'*******************************************************************************
Function CallBack()
	Dim a,b,c,d
	a="https://graph.qq.com/oauth2.0/token"
	Set b=ZBQQConnect_toObject("{}")
	Call ZBQQConnect_addobj(b,"grant_type","authorization_code")
	Call ZBQQConnect_addobj(b,"client_id",strAppKey)
	Call ZBQQConnect_addobj(b,"client_secret",strAppSecret)
	Call ZBQQConnect_addobj(b,"code",Request.QueryString("code"))
	Call ZBQQConnect_addobj(b,"state","zsxsoft")
	Call ZBQQConnect_addobj(b,"redirect_uri",strOauthCallBackUrl)
	c=ZBQQConnect_toStr(b)
	d=gethttp(a&"?"&c)
	CallBack=Split(Split(d,"=")(1),"&")(0)
	Session(strSessionClsID&"ZBQQConnect_strAccessToken")=CallBack
	
End Function
'*******************************************************************************
'** 得到OpenID                                                         **
'*******************************************************************************
Function GetOpenId(AccessToken)
	Dim a,b,c,d
	a="https://graph.qq.com/oauth2.0/me"
	b="?access_token="&AccessToken
	OpenId=Split(Split(gethttp(a&b),"""openid"":""")(1),"""")(0)
	Session(strSessionClsID&"ZBQQConnect_strOpenID")=OpenId
	strOpenID=Session(strSessionClsID&"ZBQQConnect_strOpenID")
	strAccToken=Session(strSessionClsID&"ZBQQConnect_strAccessToken")
End Function
'*******************************************************************************
'** 组合Url(Oauth 2.0)                                                         **
'*******************************************************************************
Function MakeOauth2Url(ByRef oauth_url,ip,content)
	dim iscustom
	if ip="sdk_custom" then iscustom=true
	Call ZBQQConnect_addobj(objJSON,"oauth_consumer_key",strAppKey) '设置APPKEY
	Call ZBQQConnect_addobj(objJSON,"access_token",strAccToken)
	Call ZBQQConnect_addobj(objJSON,"openid",strOpenID)
	Call ZBQQConnect_addobj(objJSON,"oauth_version","2.a")
	Call ZBQQConnect_addobj(objJSON,"clientip",getIP)
	Call ZBQQConnect_addobj(objJSON,"scope","all")
	if iscustom<>true then 
		MakeAPIPar oauth_url,ip,content
	else
		set objJSON=ZBQQConnect_toObject(ZBQQConnect_JSONExtendBasic(objJSON,content))
	End If
	oauth_url=replace(oauth_url,"<strPrototype>",strPrototype)
	strMadeUpUrl=ZBQQConnect_toStr(objJSON)
	if bolDebugMsg=true then response.write "<div class='ZBQQConnect_Debug'><font color='black'>最终生成：" & oauth_url&"?"&strMadeUpUrl & "</font></div>"
	MakeOauth2Url=oauth_url&"?"&strMadeUpUrl
	strPostUrl = oauth_url
End Function
'*******************************************************************************
'** Encode地址                                                                **
'*******************************************************************************
Function strUrlEnCode(byVal strUrl)
	'GB2312用户请把下面两行的注释去掉
	'Session.CodePage=65001
	strUrlEnCode = Server.URLEncode(strUrl)
	strUrlEnCode = Replace(strUrlEnCode,"%5F","_")
	strUrlEnCode = Replace(strUrlEnCode,"%2E",".")
	strUrlEnCode = Replace(strUrlEnCode,"%2D","-")
	strUrlEnCode = Replace(strUrlEnCode,"+","%20")
	'Session.CodePage=936
End Function

'*******************************************************************************
'** GetHttp                                                                   **
'*******************************************************************************
Function gethttp(gethttp_url)
    on error resume next
	dim a
	a=gethttp_url
	If intErrorCount>intRepeatMax Then 
		gethttp=False
		exit function
	End If 
	If Instr(a,"<strPrototype>") then a=replace(a,"<strPrototype>",strPrototype)
	objXmlhttp.SetTimeOuts 10000, 10000, 10000, 10000 
	objXmlhttp.Open "GET",a,False
	objXmlhttp.SetRequestHeader "User-Agent",strUserAgent
	objXmlhttp.Send

    if err.number=0 then
		gethttp = ZBQQConnect_BytesToBstr(objXmlhttp.responseBody,"utf-8")
	else
		gethttp = gethttp(gethttp_url)
	end if
End Function

'*******************************************************************************
'** PostHttp                                                                  **
'*******************************************************************************
Function posthttp(posthttp_url)
    'on error resume next
	If intErrorCount>intRepeatMax Then 
		posthttp=False
		exit function
	End If 
	dim a
	a=strPostUrl
	If Instr(a,"<strPrototype>") then a=replace(a,"<strPrototype>",strPrototype)
	objXmlhttp.SetTimeOuts 10000, 10000, 10000, 10000 
	objXmlhttp.Open "POST",strPostUrl,False
	objXmlhttp.SetRequestHeader "User-Agent",strUserAgent
	objXmlhttp.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
	objXmlhttp.Send strMadeUpUrl
    if err.number=0 then
		posthttp = ZBQQConnect_BytesToBstr(objXmlhttp.responseBody,"utf-8")
	else
		posthttp = posthttp(posthttp_url)
	end if
	strMadeUpUrl=""
End Function 


Function getIP
	dim a,b
	a=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	b=Request.ServerVariables("REMOTE_ADDR")
	if b="" then getIP=a else getIP=b
End Function
End Class
'*******************************************************************************
'** 源代码处理                                                                **
'*******************************************************************************
Function ZBQQConnect_BytesToBstr(body,Cset)
	dim objstream
	set objstream=createobject("adodb.stream")
	objstream.Type = 1
	objstream.Mode =3
	objstream.Open
	objstream.Write body
	objstream.Position = 0
	objstream.Type = 2
	objstream.Charset = Cset
	ZBQQConnect_BytesToBstr = objstream.ReadText
	objstream.Close
	set objstream=nothing
End Function
%>