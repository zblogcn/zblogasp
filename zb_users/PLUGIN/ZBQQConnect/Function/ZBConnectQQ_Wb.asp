<%
''*****************************************************
'   ZSXSOFT 腾讯微博SDK类
'   设置APPKEY：class.strAppKey=""
'   设置APPSECRET：class.strAppSecret=""
'   设置callback地址：class.callbackurl=""
'   得到是否登陆：class.logined=true 或 false
'   其他参见Readme.txt
''*****************************************************
Class ZBQQConnect_Wb

'*******************************************************************************
'** 定义变量                                                                    **
'*******************************************************************************
Private strOauth1BaseString,strOauthToken,strPostUrl,strOauthSessionUserID,strOauthSessionToken,strOauthSessionSecret,intErrorCount,strOauthTokenSecret
Private strMadeUpUrl
Private strHttptype
Private Oauth1_RequestToken_url ,Oauth1_authorize_url ,Oauth1_accesstoken_url 
Private Oauth2_authorize_url,Oauth2_accesstoken_url
Private strAppKey,strAppSecret
Private intRepeatMax,aryMutliContent
Private strContentWithReplaceEnter,strContentWithOutEncode,strOauthCallbackUrl,strWithOutOauthSignature
Private strPictureAddressInServer,tempid,objJSON,bolDebugMsg
Private objXmlhttp,strUserAgent
Private strPrototype,strOauthVersion,ZC_BLOG_CLSID

'*******************************************************************************
'** 初始化                                                                **
'*******************************************************************************
Sub Class_Initialize()
	ZC_BLOG_CLSID=""
	'*************在这里配置API地址*********************************************
	'这里是Oauth1的地址
	Oauth1_RequestToken_url = "https://open.t.qq.com/cgi-bin/request_token"
	Oauth1_authorize_url = "https://open.t.qq.com/cgi-bin/authorize"
	Oauth1_accesstoken_url = "https://open.t.qq.com/cgi-bin/access_token"
	'这里是Oauth2的地址
	Oauth2_authorize_url = "https://open.t.qq.com/cgi-bin/oauth2/authorize"
	Oauth2_accesstoken_url = "https://open.t.qq.com/cgi-bin/oauth2/access_token"
	'下面是初始配置
	strHttptype = "GET&"
	intErrorCount=0
	intRepeatMax=3
	set objJSON=ZBQQConnect_toobject("{}")
	set objXmlhttp=server.CreateObject("msxml2.serverXmlhttp")
	version="1.0"
	strPrototype="http://"
	strUserAgent="ZBQQConnect ByZSXSOFT"
	'需要读取数据库等可在这里插入代码
End Sub
'****************************************************************************
'** - -.
'****************************************************************************
Public Property Let UserID(str)
		strOauthSessionUserID=str
End Property
Public Property Get UserID
		UserID=strOauthSessionUserID
End Property
'****************************************************************************
'** - -.
'****************************************************************************
Public Property Let Secret(str)
		strOauthSessionSecret=str
End Property
Public Property Get Secret
		Secret=strOauthSessionSecret
End Property
'****************************************************************************
'** - -.
'****************************************************************************
Public Property Let Token(str)
		strOauthSessionToken=str
End Property
Public Property Get Token
		Token=strOauthSessionToken
End Property
'*******************************************************************************
'** 回收资源                            
'*******************************************************************************
Sub Class_Terminate()
	set objXmlhttp=nothing
End Sub
'*******************************************************************************
'** 得到用户名
'*******************************************************************************
Public Property Get Username
	Username = strOauthSessionUserID
End Property
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
		if strOauthVersion="1.0" then strPrototype="http://" else strPrototype="https://"
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
'** 关于
'****************************************************************************
Public Function About
	Set About=ZBQQConnect_Toobject("{'author':'ZSXSOFT','url':'http://www.zsxsoft.com','version':'ZBQQConnect 2.3'}")
End Function
'****************************************************************************
'** 得到是否登陆
'****************************************************************************
Public Property Get  logined()
	if strOauthSessionUserID = "" Or strOauthSessionUserID = "true" or  request.QueryString("typ")="wblogout" then
		logined=false
	else
		logined=true
	end if
end Property
'****************************************************************************
'** 注销
'****************************************************************************
Public Sub logout()
	strOauthSessionUserID=""
	strOauthSessionToken=""
	strOauthSessionSecret=""
	''数据库代码在这里加入
End Sub
'*******************************************************************************
'** RunAPI                                                                **
'*******************************************************************************
Function API(url,json,httptype)
	If Right(httptype,1)<>"&" Then httptype=httptype&"&"
	strHttpType=httptype
	API=Run(256,url,"","",json)
End Function


Function Run(type0,content,ip,pic,id)
	select case int(type0)
	case 1
		If strOauthVersion="1.0" then 
			Call get_oauth_http(MakeOauthUrl(Oauth1_RequestToken_url,empty,empty))
			Run=Oauth1_authorize_url&"?"&"oauth_token="&Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthToken")
		Else
			Run=Oauth2_authorize_url&"?client_id="&strAppKey&"&response_type=code&redirect_uri=" & strOauthCallbackUrl
		End If
	case 11
		If strOauthVersion="1.0" then 
			Run=get_oauth_http(MakeOauthUrl(Oauth1_accesstoken_url,empty,empty))
		Else
			Run=get_oauth_http(Oauth2_accesstoken_url&"?client_id="&strAppKey&"&client_secret="&strAppSecret&"&grant_type=authorization_code&code="&server.URLEncode(request.QueryString("code"))&"&redirect_uri="&strOauthCallbackUrl)
		end if
	case 256
		if strHttptype="GET&" then
			Run=gethttp(MakeOauthUrl(content,"sdk_custom",id))
		elseif strHttptype="POST&" then
			Run=posthttp(MakeOauthUrl(content,"sdk_custom",id))
		End If
	end select
	If bolDebugMsg=true then response.write "<font color='darkyellow'>返回结果：" &run & "</font>"
	tempid=""
	aryMutliContent=""
	strHttptype = "GET&"
End Function
'*******************************************************************************
'** 组合Url                                                                    **
'*******************************************************************************
Function MakeOauthUrl(ByRef oauth_url,ip,content)
	If strOauthVersion="1.0" Then
		MakeOauthUrl=MakeOauth1Url(oauth_url,ip,content)
	Else
		MakeOauthUrl=MakeOauth2Url(oauth_url,ip,content)
	End If
End Function
'*******************************************************************************
'** 组合Url(Oauth 2.0)                                                         **
'*******************************************************************************
Function MakeOauth2Url(ByRef oauth_url,ip,content)
	dim iscustom
	if ip="sdk_custom" then iscustom=true
	Call ZBQQConnect_addobj(objJSON,"oauth_consumer_key",strAppKey) '设置APPKEY
	Call ZBQQConnect_addobj(objJSON,"access_token",Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthToken"))
	Call ZBQQConnect_addobj(objJSON,"openid",Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthTokenSecret"))
	Call ZBQQConnect_addobj(objJSON,"oauth_version","2.a")
	Call ZBQQConnect_addobj(objJSON,"clientip",ZBQQConnect_getIP)
	Call ZBQQConnect_addobj(objJSON,"scope","all")
	if iscustom<>true then 
		MakeAPIPar oauth_url,ip,content
	else
		set objJSON=ZBQQConnect_toObject(ZBQQConnect_JSONExtendBasic(objJSON,content))
	End If
	oauth_url=replace(oauth_url,"<strPrototype>",strPrototype)
	strMadeUpUrl=ZBQQConnect_toStr(objJSON)
	strMadeUpUrl=ZBQQConnect_toStr(objJSON)
	if bolDebugMsg=true then response.write "<div class='ZBQQConnect_Debug'><font color='black'>最终生成：" & oauth_url&"?"&strMadeUpUrl & "</font></div>"
	MakeOauth2Url=oauth_url&"?"&strMadeUpUrl
	strPostUrl = oauth_url
End Function
'*******************************************************************************
'** 组合Url(Oauth 1.0)                                                         **
'*******************************************************************************
Function MakeOauth1Url(ByRef oauth_url,ip,content)
	dim iscustom
	if ip="sdk_custom" then iscustom=true

	Call ZBQQConnect_addobj(objJSON,"oauth_nonce",makePassword(12))   '添加随机码
	Call ZBQQConnect_addobj(objJSON,"oauth_timestamp",DateDiff("s","01/01/1970 08:00:00",Now()))  '添加时间戳
	Call ZBQQConnect_addobj(objJSON,"oauth_version","1.0")   '设置oauth版本
	Call ZBQQConnect_addobj(objJSON,"oauth_consumer_key",strAppKey) '设置APPKEY
	Call ZBQQConnect_addobj(objJSON,"oauth_signature_method","HMAC-SHA1") '设置加密方法
		If oauth_url<>Oauth1_RequestToken_url Then
			strOauthCallbackUrl = ""  
			if iscustom=true then
				strOauthTokenSecret=Secret
				Call ZBQQConnect_addobj(objJSON,"oauth_token",Token) '设置token
				set objJSON=ZBQQConnect_toObject(ZBQQConnect_JSONExtendBasic(objJSON,content))
			end if
			If oauth_url=Oauth1_accesstoken_url Then
				Call ZBQQConnect_addobj(objJSON,"oauth_token",Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthToken")) '设置token
				strOauthTokenSecret = Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthTokenSecret") 
				Call ZBQQConnect_addobj(objJSON,"oauth_verifier",Request.QueryString("oauth_verifier")) '设置verifier
			'Else
			'	MakeAPIPar oauth_url,ip,content
			End If
		Else
			Call ZBQQConnect_addobj(objJSON,"oauth_callback",strOauthCallbackUrl) '回调地址
		End If

	oauth_url=replace(oauth_url,"<strPrototype>",strPrototype)
	strMadeUpUrl=ZBQQConnect_toStr(objJSON)
	if bolDebugMsg=true then response.write "<div class='ZBQQConnect_Debug'><font color='red'>不包含signature：" & strMadeUpUrl & "</font></br>"
	strOauth1BaseString=strHttptype & strUrlEnCode(oauth_url) & "&" & strUrlEnCode(strMadeUpUrl)
	if bolDebugMsg=true then response.write "<font color='blue'>BaseString：" & strOauth1BaseString & "</font></br><font color='green'>密钥：" & strAppSecret&"&"&strOauthTokenSecret & "</font><br/>"
	strWithOutOauthSignature=strUrlEnCode(ZBQQConnect_b64_hmac_sha1(strAppSecret&"&"&strOauthTokenSecret,strOauth1BaseString))
	Call ZBQQConnect_addobj(objJSON,"oauth_signature",strWithOutOauthSignature)
	strMadeUpUrl=ZBQQConnect_toStr(objJSON)
	if bolDebugMsg=true then response.write "<font color='black'>最终生成：" & strMadeUpUrl & "</font></div>"
	MakeOauth1Url=oauth_url&"?"&strMadeUpUrl
	strPostUrl = oauth_url
End Function
'*******************************************************************************
'** 制造随机串                                                                **
'*******************************************************************************
Function makePassword(byVal maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
		whatsNext = Int((1 - 0 + 1) * Rnd + 0)
		If whatsNext = 0 Then
			upper = 122
			lower = 100
		Else
			upper = 57
			lower = 48
		End If
		strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	makePassword = strNewPass
End function
'*******************************************************************************
'** Encode地址                                                                **
'*******************************************************************************
Public Function strUrlEnCode(byVal strUrl)
	'Session.CodePage=65001
	strUrlEnCode = Server.URLEncode(strUrl)
	strUrlEnCode = Replace(strUrlEnCode,"%5F","_")
	strUrlEnCode = Replace(strUrlEnCode,"%2E",".")
	strUrlEnCode = Replace(strUrlEnCode,"%2D","-")
	strUrlEnCode = Replace(strUrlEnCode,"+","%20")
	'Session.CodePage=936
End Function

'*******************************************************************************
'** 获取Oauth信息                                                             **
'*******************************************************************************
function get_oauth_http(byval oauthUrl)
	objXmlhttp.Open "GET",oauthUrl,False,"",""
	objXmlhttp.Send
	Dim ary_responseText
	ary_responseText = Split(Replace(objXmlhttp.responseText,"=","&"),"&")
	if instr(objXmlhttp.responseText,"=") then
		
		if strOauthVersion="1.0" then
			Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthToken") = ary_responseText(1)
			strOauthSessionToken=ary_responseText(1)
			Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthTokenSecret") = ary_responseText(3)
			strOauthSessionSecret=ary_responseText(3)
			if ubound(ary_responseText)>=5 then
				Session(ZC_BLOG_CLSID&"ZBQQConnect_strUserName") = ary_responseText(5)
				strOauthSessionUserID=ary_responseText(5)
			end if
		else
			Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthToken")=ary_responseText(1)
			Session(ZC_BLOG_CLSID&"ZBQQConnect_strOauthTokenSecret")=request.QueryString("openid")
			Session(ZC_BLOG_CLSID&"ZBQQConnect_strUserName")= ary_responseText(7)
			
		
		end if
	end if
	'ary_responseText(1)  这个是oauthtoken
	'ary_responseText(3)  这个是tokensecret
	'ary_responseText(5)  这个是名字
	'需要录入数据库等可在这里插入代码
	get_oauth_http=objXmlhttp.responseText
End function
'*******************************************************************************
'** GetHttp                                                                   **
'*******************************************************************************
Public Function gethttp(gethttp_url)
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
Public Function posthttp(posthttp_url)
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
	objXmlhttp.Send strMadeUpUrl
    if err.number=0 then
		posthttp = ZBQQConnect_BytesToBstr(objXmlhttp.responseBody,"utf-8")
	else
		posthttp = posthttp(posthttp_url)
	end if
	strMadeUpUrl=""
End Function 


Function ZBQQConnect_getIP
	dim a,b
	a=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	b=Request.ServerVariables("REMOTE_ADDR")
	if b="" then ZBQQConnect_getIP=a else ZBQQConnect_getIP=b
End Function
End Class
%>