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
Private strPrototype,strOauthVersion,ZC_BLOG_CLSID
Public fakeQQConnect
Public debugMsg

Sub AddDebug(Content)
	debugMsg=debugMsg &"<br/>【"& Now & "】"&Content
End Sub
'*******************************************************************************
'** 初始化                                                                **
'*******************************************************************************
Sub Class_Initialize()
	ZC_BLOG_CLSID=ZC_BLOG_CLSID
	strHttptype = "GET&"
	intErrorCount=0
	intRepeatMax=3
	set fakeQQConnect=New ZBQQConnect_Wb
	set objJSON=ZBQQConnect_json.toobject("{}")
	strOpenID=Session(ZC_BLOG_CLSID&"ZBQQConnect_strOpenID")
	strAccToken=Session(ZC_BLOG_CLSID&"ZBQQConnect_strAccessToken")
	version="2.0"
	strPrototype="http://"

	'需要读取数据库等可在这里插入代码
End Sub
'*******************************************************************************
'** 回收资源                            
'*******************************************************************************
Sub Class_Terminate()
End Sub
'****************************************************************************
'** 设置CallBackUrl
'****************************************************************************
Public Property Let callbackurl(url)
		strOauthCallbackUrl=fakeQQConnect.strUrlEnCode(url)
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
	Set About=ZBQQConnect_json.Toobject("{'author':'ZSXSOFT','url':'http://www.zsxsoft.com','version':'ZBQQConnect 1.0'}")
End Function

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
	ZBQQConnect_DB.OpenID=strOpenID
	ZBQQConnect_DB.LoadInfo 4
	ZBQQConnect_DB.Del 
'	Response.Cookies("QQOPENID")=""
'	Response.Cookies("QQAccessToken")=""
End Sub
'*******************************************************************************
'** RunAPI                                                                **
'*******************************************************************************
Function API(url,json,httptype)
'	Response.Cookies("QQOPENID")=strOpenID
'	Response.Cookies("QQOPENID").Expires = DateAdd("d", 90, now)
'	Response.Cookies("QQOPENID").Path="/"
'	Response.Cookies("QQAccessToken")=strAccToken
'	Response.Cookies("QQAccessToken").Expires = DateAdd("d", 90, now)
'	Response.Cookies("QQAccessToken").Path="/"
	AddDebug "RunAPI"
	AddDebug "OpenID="&strOpenID
	AddDebug "AccessToken="&strAccToken
	AddDebug "URL="&url
	AddDebug "PrivateJSON="&json
	Dim MUrl
	If Right(httptype,1)<>"&" Then httptype=httptype&"&"
	strHttpType=httptype
	AddDebug "httptype="&strHttpType
	MUrl=MakeOauthUrl(url,"sdk_custom",json)
	AddDebug "URL2="&MUrl
	AddDebug "PostData="&strMadeUpUrl
	if strHttptype="GET&" then
		API=ZBQQConnect_Net.GetHttp(Murl)
	elseif strHttptype="POST&" then
		API=ZBQQConnect_Net.PostHttp(url,strMadeUpUrl)
	end if
	AddDebug "Result="&API
	tempid=""
	aryMutliContent=""
	strHttptype = "GET&"
	set objjson=ZBQQConnect_json.toObject("{}")
End Function


'*******************************************************************************
'** 组合Url                                                                    **
'*******************************************************************************
Function MakeOauthUrl(ByRef oauth_url,ip,content)
	MakeOauthUrl=MakeOauth2Url(oauth_url,ip,content)
End Function
'*******************************************************************************
'** 分享                                                         **
'*******************************************************************************
Function Share(title,url,comment,summary,images,nswb)
	Dim b
	Set b=ZBQQConnect_json.toObject("{}")
	Call ZBQQConnect_json.addobj(b,"title",title)
	Call ZBQQConnect_json.addobj(b,"url",url)
	Call ZBQQConnect_json.addobj(b,"comment",comment)
	Call ZBQQConnect_json.addobj(b,"summary",summary)
	If images="~" Then images=""
	Call ZBQQConnect_json.addobj(b,"images",images)
	Call ZBQQConnect_json.addobj(b,"nswb",nswb)
	b=ZBQQConnect_json.toJSON(b)
	Share=API("https://graph.qq.com/share/add_share",b,"POST&")
End Function
'*******************************************************************************
'** 发微博                                                         **
'*******************************************************************************
Function t(content)
	Dim b
	Set b=ZBQQConnect_json.toObject("{}")
	Call ZBQQConnect_json.addobj(b,"format","json")
	Call ZBQQConnect_json.addobj(b,"content",content)
	Call ZBQQConnect_json.addobj(b,"clientip",fakeQQConnect.getIP)
	b=ZBQQConnect_json.toJSON(b)
	t=API("https://graph.qq.com/t/add_t",b,"POST&")
End Function
'*******************************************************************************
'** 得到验证URL                                                         **
'*******************************************************************************
Function Authorize()
	Dim a,b,c
	a="https://graph.qq.com/oauth2.0/authorize"
	set b=ZBQQConnect_json.toObject("{}")
	Call ZBQQConnect_json.addobj(b,"response_type","code")
	Call ZBQQConnect_json.addobj(b,"client_id",strAppKey)
	Call ZBQQConnect_json.addobj(b,"redirect_uri",strOauthCallBackUrl)
	Call ZBQQConnect_json.addobj(b,"scope","get_user_info,add_share,get_info,add_idol,add_t")
	Call ZBQQConnect_json.addobj(b,"state","zsxsoft")

	c=ZBQQConnect_json.toStr(b)
	Authorize=a&"?"&c'ZBQQConnect_Net.GetHttp(a&"?"&c)
End Function
'*******************************************************************************
'** CallBack                                                         **
'*******************************************************************************
Function CallBack()
	Dim a,b,c,d
	a="https://graph.qq.com/oauth2.0/token"
	Set b=ZBQQConnect_json.toObject("{}")
	Call ZBQQConnect_json.addobj(b,"grant_type","authorization_code")
	Call ZBQQConnect_json.addobj(b,"client_id",strAppKey)
	Call ZBQQConnect_json.addobj(b,"client_secret",strAppSecret)
	Call ZBQQConnect_json.addobj(b,"code",Request.QueryString("code"))
	Call ZBQQConnect_json.addobj(b,"state","zsxsoft")
	Call ZBQQConnect_json.addobj(b,"redirect_uri",strOauthCallBackUrl)
	c=ZBQQConnect_json.toStr(b)
	d=ZBQQConnect_Net.GetHttp(a&"?"&c)
	
	CallBack=Split(Split(d,"=")(1),"&")(0)
	Session(ZC_BLOG_CLSID&"ZBQQConnect_strAccessToken")=CallBack
	
End Function
'*******************************************************************************
'** 得到OpenID                                                         **
'*******************************************************************************
Function GetOpenId(AccessToken)
	Dim a,b,c,d
	a="https://graph.qq.com/oauth2.0/me"
	b="?access_token="&AccessToken
	OpenId=Split(Split(ZBQQConnect_Net.GetHttp(a&b),"""openid"":""")(1),"""")(0)
	Session(ZC_BLOG_CLSID&"ZBQQConnect_strOpenID")=OpenId
	strOpenID=Session(ZC_BLOG_CLSID&"ZBQQConnect_strOpenID")
	strAccToken=Session(ZC_BLOG_CLSID&"ZBQQConnect_strAccessToken")
End Function
'*******************************************************************************
'** 组合Url(Oauth 2.0)                                                         **
'*******************************************************************************
Function MakeOauth2Url(ByRef oauth_url,ip,content)
	dim iscustom
	Call ZBQQConnect_json.addobj(objJSON,"oauth_consumer_key",strAppKey) '设置APPKEY
	Call ZBQQConnect_json.addobj(objJSON,"access_token",strAccToken)
	Call ZBQQConnect_json.addobj(objJSON,"openid",strOpenID)
	set objJSON=ZBQQConnect_json.toObject(ZBQQConnect_JSONExtendBasic(objJSON,content))
	oauth_url=replace(oauth_url,"<strPrototype>",strPrototype)
	strMadeUpUrl=ZBQQConnect_json.toStr(objJSON)
	if bolDebugMsg=true then response.write "<div class='ZBQQConnect_Debug'><font color='black'>最终生成：" & oauth_url&"?"&strMadeUpUrl & "</font></div>"
	MakeOauth2Url=oauth_url&"?"&strMadeUpUrl
	strPostUrl = oauth_url
End Function

End Class


%>