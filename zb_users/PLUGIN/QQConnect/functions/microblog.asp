<%
''*****************************************************
'   ZSXSOFT 腾讯微博SDK类
'   设置callback地址：class.callbackurl=""
'   其他参见http://www.zsxsoft.com/archives/202.html
''*****************************************************
Class qqconnect_weibo

'*******************************************************************************
'** 定义变量                                                                    **
'*******************************************************************************
Private strOauth1BaseString,strOauthToken,strAccessToken,strOauthTokenSecret
Private strHttptype
Private Oauth1_RequestToken_url ,Oauth1_authorize_url ,Oauth1_accesstoken_url 
Private aryMutliContent
Private strContentWithReplaceEnter,strContentWithOutEncode,strOauthCallbackUrl,strWithOutOauthSignature
Private strPictureAddressInServer,tempid,objJSON,bolDebugMsg
Private objXmlhttp,strUserAgent
Public strPostUrl,strMadeUpUrl

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
	'下面是初始配置
	strHttptype = "GET&"
	set objJSON=qqconnect_json.toobject("{}")
	strUserAgent="qqconnect By ZSXSOFT"
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
'***************************************************************
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
	Set About=qqconnect_json.toobject("{'author':'ZSXSOFT','url':'http://www.zsxsoft.com','version':'qqconnect 2.3'}")
End Function
'*******************************************************************************
'** RunAPI                                                                **
'*******************************************************************************
Function t(content,pic)
	Dim b
	Set b=qqconnect_json.toobject("{}")
	Call qqconnect_json.addObj(b,"content",content)
	Call qqconnect_json.addObj(b,"clientip",qqconnect_getip())
	Call qqconnect_json.addObj(b,"format","json")
	Call qqconnect_json.addObj(b,"syncflag",1)
	If pic<>"~" and  pic<>"" Then
		Call qqconnect_json.addObj(b,"pic_url",pic)
	
		t=API("http://open.t.qq.com/api/t/add_pic_url",qqconnect_json.TOJSON(b),"POST&")
	Else
		t=API("http://open.t.qq.com/api/t/add",qqconnect_json.TOJSON(b),"POST&")
	End If
End Function
Function r(content,id)
	Dim b
	Set b=qqconnect_json.toobject("{}")
	Call qqconnect_json.addObj(b,"content",content)
	Call qqconnect_json.addObj(b,"clientip",qqconnect_getip())
	Call qqconnect_json.addObj(b,"format","json")
	Call qqconnect_json.addObj(b,"syncflag",1)
	Call qqconnect_json.addObj(b,"reid",id)
	r=API("http://open.t.qq.com/api/t/re_add",qqconnect_json.TOJSON(b),"POST&")
End Function

Function API(url,json,httptype)
	If Right(httptype,1)<>"&" Then httptype=httptype&"&"
	strHttpType=httptype
	API=Run(256,url,"","",json)
	strMadeUpUrl=""
End Function


Function Run(type0,content,ip,pic,id)
	select case int(type0)
	case 1
		Call get_oauth_http(MakeOauthUrl(Oauth1_RequestToken_url,empty,empty))
		Run=Oauth1_authorize_url&"?"&"oauth_token="&Session(ZC_BLOG_CLSID&"qqconnect_strOauthToken")
	case 11
		Run=get_oauth_http(MakeOauthUrl(Oauth1_accesstoken_url,empty,empty))
	case 256
		if strHttptype="GET&" then
			Run=qqconnect.n.gethttp(MakeOauthUrl(content,"sdk_custom",id))
		elseif strHttptype="POST&" then
			Call MakeOauthUrl(content,"sdk_custom",id)
			Run=qqconnect.n.posthttp(strPostUrl,strMadeUpUrl)
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
	MakeOauthUrl=MakeOauth1Url(oauth_url,ip,content)
End Function

'*******************************************************************************
'** 组合Url(Oauth 1.0)                                                         **
'*******************************************************************************
Function MakeOauth1Url(ByRef oauth_url,ip,content)
	dim iscustom

	Call qqconnect_json.addObj(objJSON,"oauth_nonce",makePassword(12))   '添加随机码
	Call qqconnect_json.addObj(objJSON,"oauth_timestamp",DateDiff("s","01/01/1970 08:00:00",Now()))  '添加时间戳
	Call qqconnect_json.addObj(objJSON,"oauth_version","1.0")   '设置oauth版本
	Call qqconnect_json.addObj(objJSON,"oauth_consumer_key",qqconnect.config.weibo.appkey) '设置qqconnect.config.weibo.appkey
	Call qqconnect_json.addObj(objJSON,"oauth_signature_method","HMAC-SHA1") '设置加密方法
		If oauth_url<>Oauth1_RequestToken_url Then
			strOauthCallbackUrl = ""  
			strOauthTokenSecret=qqconnect.config.weibo.secret
			Call qqconnect_json.addObj(objJSON,"oauth_token",qqconnect.config.weibo.token) '设置token
			set objJSON=qqconnect_json.e(objJSON,qqconnect_json.toObject(content))
			If oauth_url=Oauth1_accesstoken_url Then
				Call qqconnect_json.addObj(objJSON,"oauth_token",Session(ZC_BLOG_CLSID&"qqconnect_strOauthToken")) '设置token
				strOauthTokenSecret = Session(ZC_BLOG_CLSID&"qqconnect_strOauthTokenSecret") 
				Call qqconnect_json.addObj(objJSON,"oauth_verifier",Request.QueryString("oauth_verifier")) '设置verifier
			End If
		Else
			Call qqconnect_json.addObj(objJSON,"oauth_callback",strOauthCallbackUrl) '回调地址
		End If

	strMadeUpUrl=qqconnect_json.toStr(objJSON)
	if bolDebugMsg=true then response.write "<div class='qqconnect_Debug'><font color='red'>不包含signature：" & strMadeUpUrl & "</font></br>"
	strOauth1BaseString=strHttptype & qqconnect_encodeurl(oauth_url) & "&" & qqconnect_encodeurl(strMadeUpUrl)
	if bolDebugMsg=true then response.write "<font color='blue'>BaseString：" & strOauth1BaseString & "</font></br><font color='green'>密钥：" & qqconnect.config.weibo.appsecret&"&"&strOauthTokenSecret & "</font><br/>"
	strWithOutOauthSignature=qqconnect_encodeurl(qqconnect_b64_hmac_sha1(qqconnect.config.weibo.appsecret&"&"&strOauthTokenSecret,strOauth1BaseString))
	Call qqconnect_json.addObj(objJSON,"oauth_signature",strWithOutOauthSignature)
	strMadeUpUrl=qqconnect_json.toStr(objJSON)
	if bolDebugMsg=true then response.write "<font color='black'>最终生成：" & oauth_url&"?"&strMadeUpUrl & "</font></div>"
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
'** 获取Oauth信息                                                             **
'*******************************************************************************
function get_oauth_http(byval oauthUrl)
	Dim strResponse
	strResponse=qqconnect.n.GetHttp(oauthUrl)
	Dim ary_responseText
	ary_responseText = Split(Replace(strResponse,"=","&"),"&")
	if instr(strResponse,"=") then
		
		Session(ZC_BLOG_CLSID&"qqconnect_strOauthToken") = ary_responseText(1)
		qqconnect.config.weibo.token=ary_responseText(1)
		Session(ZC_BLOG_CLSID&"qqconnect_strOauthTokenSecret") = ary_responseText(3)
		qqconnect.config.weibo.secret=ary_responseText(3)
		if ubound(ary_responseText)>=5 then
			Session(ZC_BLOG_CLSID&"qqconnect_strUserName") = ary_responseText(5)
			strAccessToken=ary_responseText(5)
		end if

		
	end if
	'ary_responseText(1)  这个是oauthtoken
	'ary_responseText(3)  这个是tokensecret
	'ary_responseText(5)  这个是名字
	get_oauth_http=strResponse
End function


End Class

%>
