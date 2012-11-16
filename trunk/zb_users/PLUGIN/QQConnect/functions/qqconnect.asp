<%
''*****************************************************
'   ZSXSOFT ASP QQConnect类
''*****************************************************
Class qqconnect_connect
	
	'*******************************************************************************
	'** 定义变量                                                                    **
	'*******************************************************************************
	Private strOauthToken,strPostUrl,strOauthSessionUserID,strOauthTokenSecret
	Private strMadeUpUrl
	Private strHttptype
	Private aryMutliContent
	Private strContentWithReplaceEnter,strContentWithOutEncode,strOauthCallbackUrl,strWithOutOauthSignature
	Private tempid,objJSON,bolDebugMsg
	Private strPrototype,strOauthVersion,ZC_BLOG_CLSID
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
		set objJSON=qqconnect_json.toobject("{}")
		'qqconnect.config.qqconnect.openid=Session(ZC_BLOG_CLSID&"qqconnect_connect_qqconnect.config.qqconnect.openid")
		'qqconnect.config.qqconnect.accesstoken=Session(ZC_BLOG_CLSID&"qqconnect_connect_strAccessToken")
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
			strOauthCallbackUrl=qqconnect_encodeurl(url)
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
		Set About=qqconnect_json.Toobject("{'author':'ZSXSOFT','url':'http://www.zsxsoft.com','version':'qqconnect_connect 1.0'}")
	End Function
	
	'****************************************************************************
	'** 得到OpenID
	'****************************************************************************
	Public Property Let OpenID(str)
			qqconnect.config.qqconnect.openid=str
	End Property
	Public Property Get OpenID
			OpenID=qqconnect.config.qqconnect.openid
	End Property
	'****************************************************************************
	'** 得到AccessToken
	'****************************************************************************
	Public Property Let AccessToken(str)
			qqconnect.config.qqconnect.accesstoken=str
	End Property
	Public Property Get AccessToken
			AccessToken=qqconnect.config.qqconnect.accesstoken
	End Property
	
	'****************************************************************************
	'** 注销
	'****************************************************************************
	Public Sub logout()
		qqconnect_connect_DB.OpenID=qqconnect.config.qqconnect.openid
		qqconnect_connect_DB.LoadInfo 4
		qqconnect_connect_DB.Del 
	'	Response.Cookies("QQOPENID")=""
	'	Response.Cookies("QQAccessToken")=""
	End Sub
	'*******************************************************************************
	'** RunAPI                                                                **
	'*******************************************************************************
	Function API(url,json,httptype)
	'	Response.Cookies("QQOPENID")=qqconnect.config.qqconnect.openid
	'	Response.Cookies("QQOPENID").Expires = DateAdd("d", 90, now)
	'	Response.Cookies("QQOPENID").Path="/"
	'	Response.Cookies("QQAccessToken")=qqconnect.config.qqconnect.accesstoken
	'	Response.Cookies("QQAccessToken").Expires = DateAdd("d", 90, now)
	'	Response.Cookies("QQAccessToken").Path="/"
		AddDebug "RunAPI"
		AddDebug "OpenID="&qqconnect.config.qqconnect.openid
		AddDebug "AccessToken="&qqconnect.config.qqconnect.accesstoken
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
			API=qqconnect.n.GetHttp(Murl)
		elseif strHttptype="POST&" then
			API=qqconnect.n.PostHttp(url,strMadeUpUrl)
		end if
		AddDebug "Result="&API
		tempid=""
		aryMutliContent=""
		strHttptype = "GET&"
		set objjson=qqconnect_json.toObject("{}")
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
		Set b=qqconnect_json.toObject("{}")
		Call qqconnect_json.addobj(b,"title",title)
		Call qqconnect_json.addobj(b,"url",url)
		Call qqconnect_json.addobj(b,"comment",comment)
		Call qqconnect_json.addobj(b,"summary",summary)
		If images="~" Then images=""
		Call qqconnect_json.addobj(b,"images",images)
		Call qqconnect_json.addobj(b,"nswb",nswb)
		b=qqconnect_json.toJSON(b)
		Share=API("https://graph.qq.com/share/add_share",b,"POST&")
	End Function
	'*******************************************************************************
	'** 发微博                                                         **
	'*******************************************************************************
	Function t(content)
		Dim b
		Set b=qqconnect_json.toObject("{}")
		Call qqconnect_json.addobj(b,"format","json")
		Call qqconnect_json.addobj(b,"content",content)
		Call qqconnect_json.addobj(b,"clientip",qqconnect_getip())
		b=qqconnect_json.toJSON(b)
		t=API("https://graph.qq.com/t/add_t",b,"POST&")
	End Function
	'*******************************************************************************
	'** 得到验证URL                                                         **
	'*******************************************************************************
	Function Authorize()
		Dim a,b,c
		a="https://graph.qq.com/oauth2.0/authorize"
		set b=qqconnect_json.toObject("{}")
		Call qqconnect_json.addobj(b,"response_type","code")
		Call qqconnect_json.addobj(b,"client_id",qqconnect.config.qqconnect.appid)
		Call qqconnect_json.addobj(b,"redirect_uri",strOauthCallBackUrl)
		Call qqconnect_json.addobj(b,"scope","get_user_info,add_share,get_info,add_idol,add_t")
		Call qqconnect_json.addobj(b,"state","zsxsoft")
	
		c=qqconnect_json.toStr(b)
		Authorize=a&"?"&c'qqconnect.n.GetHttp(a&"?"&c)
	End Function
	'*******************************************************************************
	'** CallBack                                                         **
	'*******************************************************************************
	Function CallBack()
		Dim a,b,c,d
		a="https://graph.qq.com/oauth2.0/token"
		Set b=qqconnect_json.toObject("{}")
		Call qqconnect_json.addobj(b,"grant_type","authorization_code")
		Call qqconnect_json.addobj(b,"client_id",qqconnect.config.qqconnect.appid)
		Call qqconnect_json.addobj(b,"client_secret",qqconnect.config.qqconnect.appsecret)
		Call qqconnect_json.addobj(b,"code",Request.QueryString("code"))
		Call qqconnect_json.addobj(b,"state","zsxsoft")
		Call qqconnect_json.addobj(b,"redirect_uri",strOauthCallBackUrl)
		c=qqconnect_json.toStr(b)
		d=qqconnect.n.GetHttp(a&"?"&c)
		
		CallBack=Split(Split(d,"=")(1),"&")(0)
		Session(ZC_BLOG_CLSID&"qqconnect_connect_strAccessToken")=CallBack
		
	End Function
	'*******************************************************************************
	'** 得到OpenID                                                         **
	'*******************************************************************************
	Function GetOpenId(AccessToken)
		Dim a,b,c,d
		a="https://graph.qq.com/oauth2.0/me"
		b="?access_token="&AccessToken
		OpenId=Split(Split(qqconnect.n.GetHttp(a&b),"""openid"":""")(1),"""")(0)
		Session(ZC_BLOG_CLSID&"qqconnect_connect_qqconnect.config.qqconnect.openid")=OpenId
		qqconnect.config.qqconnect.openid=Session(ZC_BLOG_CLSID&"qqconnect_connect_qqconnect.config.qqconnect.openid")
		qqconnect.config.qqconnect.accesstoken=Session(ZC_BLOG_CLSID&"qqconnect_connect_strAccessToken")
	End Function
	'*******************************************************************************
	'** 组合Url(Oauth 2.0)                                                         **
	'*******************************************************************************
	Function MakeOauth2Url(ByRef oauth_url,ip,content)
		dim iscustom
		Call qqconnect_json.addobj(objJSON,"oauth_consumer_key",qqconnect.config.qqconnect.appid) '设置APPKEY
		Call qqconnect_json.addobj(objJSON,"access_token",qqconnect.config.qqconnect.accesstoken)
		Call qqconnect_json.addobj(objJSON,"openid",qqconnect.config.qqconnect.openid)
		set objJSON=qqconnect_json.e(objJSON,qqconnect_json.toObject(content))
		oauth_url=replace(oauth_url,"<strPrototype>",strPrototype)
		strMadeUpUrl=qqconnect_json.toStr(objJSON)
		if bolDebugMsg=true then response.write "<div class='qqconnect_connect_Debug'><font color='black'>最终生成：" & oauth_url&"?"&strMadeUpUrl & "</font></div>"
		MakeOauth2Url=oauth_url&"?"&strMadeUpUrl
		strPostUrl = oauth_url
	End Function

End Class


%>