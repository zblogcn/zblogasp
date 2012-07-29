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
Public debugMsg

Sub AddDebug(Content)
	debugMsg=debugMsg &"<br/>【"& Now & "】"&Content
End Sub
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
	Response.Cookies("QQOPENID")=""
	Response.Cookies("QQAccessToken")=""
End Sub
'*******************************************************************************
'** RunAPI                                                                **
'*******************************************************************************
Function API(url,json,httptype)
	Response.Cookies("QQOPENID")=strOpenID
	Response.Cookies("QQOPENID").Expires = DateAdd("d", 90, now)
	Response.Cookies("QQOPENID").Path="/"
	Response.Cookies("QQAccessToken")=strAccToken
	Response.Cookies("QQAccessToken").Expires = DateAdd("d", 90, now)
	Response.Cookies("QQAccessToken").Path="/"
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
		API=gethttp(Murl)
	elseif strHttptype="POST&" then
		API=posthttp(Murl)
	end if
	AddDebug "Result="&API
	tempid=""
	aryMutliContent=""
	strHttptype = "GET&"
	set objjson=zbqqconnect_toobject("{}")
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
	Set b=ZBQQConnect_toObject("{}")
	Call ZBQQConnect_addobj(b,"title",title)
	Call ZBQQConnect_addobj(b,"url",url)
	Call ZBQQConnect_addobj(b,"comment",comment)
	Call ZBQQConnect_addobj(b,"summary",summary)
	If images="~" Then images=""
	Call ZBQQConnect_addobj(b,"images",images)
	Call ZBQQConnect_addobj(b,"nswb",nswb)
	b=ZBQQConnect_toJSON(b)
	Share=API("https://graph.qq.com/share/add_share",b,"POST&")
End Function
'*******************************************************************************
'** 发微博                                                         **
'*******************************************************************************
Function t(content)
	Dim b
	Set b=ZBQQConnect_toObject("{}")
	Call ZBQQConnect_addobj(b,"format","json")
	Call ZBQQConnect_addobj(b,"content",content)
	Call ZBQQConnect_addobj(b,"clientip",getip)
	b=ZBQQConnect_toJSON(b)
	t=API("https://graph.qq.com/t/add_t",b,"POST&")
End Function
'*******************************************************************************
'** 得到验证URL                                                         **
'*******************************************************************************
Function Authorize()
	Dim a,b,c
	a="https://graph.qq.com/oauth2.0/authorize"
	set b=ZBQQConnect_ToObject("{}")
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
	Call ZBQQConnect_addobj(objJSON,"oauth_consumer_key",strAppKey) '设置APPKEY
	Call ZBQQConnect_addobj(objJSON,"access_token",strAccToken)
	Call ZBQQConnect_addobj(objJSON,"openid",strOpenID)
	set objJSON=ZBQQConnect_toObject(ZBQQConnect_JSONExtendBasic(objJSON,content))
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
	strUrlEnCode = Server.URLEncode(strUrl)
	strUrlEnCode = Replace(strUrlEnCode,"%5F","_")
	strUrlEnCode = Replace(strUrlEnCode,"%2E",".")
	strUrlEnCode = Replace(strUrlEnCode,"%2D","-")
	strUrlEnCode = Replace(strUrlEnCode,"+","%20")
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
Function ZBQQConnect_ReplaceXO(ByVal strContent):on error resume next:strContent = Replace(strContent,"&#8194;"," "):strContent = Replace(strContent,"&#8195;"," "):strContent = Replace(strContent,"&#160;"," "):strContent = Replace(strContent,"&lt;","<"):strContent = Replace(strContent,"&gt;",">"):strContent = Replace(strContent,"&amp;","&"):strContent = Replace(strContent,"&quot;",""""):strContent = Replace(strContent,"&copy;","©"):strContent = Replace(strContent,"&reg;","®"):strContent = Replace(strContent,"™","™"):strContent = Replace(strContent,"&times;","×"):strContent = Replace(strContent,"&divide;","÷"):strContent = Replace(strContent,"&nbsp;",""):strContent = Replace(strContent,"&yen;","¥"):strContent = Replace(strContent,"&ordf;","ª"):strContent = Replace(strContent,"&macr;","¯"):strContent = Replace(strContent,"&acute;","´"):strContent = Replace(strContent,"&sup1;","¹"):strContent = Replace(strContent,"&frac34;","¾"):strContent = Replace(strContent,"&Atilde;","Ã"):strContent = Replace(strContent,"&Egrave;","È"):strContent = Replace(strContent,"&Iacute;","Í"):strContent = Replace(strContent,"&Ograve;","Ò"):strContent = Replace(strContent,"&times;","×"):strContent = Replace(strContent,"&Uuml;","Ü"):strContent = Replace(strContent,"&aacute;","á"):strContent = Replace(strContent,"&aelig;","æ"):strContent = Replace(strContent,"&euml;","ë"):strContent = Replace(strContent,"&eth;","ð"):strContent = Replace(strContent,"&otilde;","õ"):strContent = Replace(strContent,"&uacute;","ú"):strContent = Replace(strContent,"&yuml;","ÿ"):strContent = Replace(strContent,"&iexcl;","¡"):strContent = Replace(strContent,"&brvbar;","¦"):strContent = Replace(strContent,"&laquo;","«"):strContent = Replace(strContent,"&deg;","°"):strContent = Replace(strContent,"&micro;","µ"):strContent = Replace(strContent,"&ordm;","º"):strContent = Replace(strContent,"&iquest;","¿"):strContent = Replace(strContent,"&Auml;","Ä"):strContent = Replace(strContent,"&Eacute;","É"):strContent = Replace(strContent,"&Icirc;","Î"):strContent = Replace(strContent,"&Oacute;","Ó"):strContent = Replace(strContent,"&Oslash;","Ø"):strContent = Replace(strContent,"&Yacute;","Ý"):strContent = Replace(strContent,"&acirc;","â"):strContent = Replace(strContent,"&ccedil;","ç"):strContent = Replace(strContent,"&igrave;","ì"):strContent = Replace(strContent,"&ntilde;","ñ"):strContent = Replace(strContent,"&ouml;","ö"):strContent = Replace(strContent,"&ucirc;","û"):strContent = Replace(strContent,"&curren;","¤"):strContent = Replace(strContent,"&copy;","©"):strContent = Replace(strContent,"&reg;","®"):strContent = Replace(strContent,"&sup3;","³"):strContent = Replace(strContent,"&cedil;","¸"):strContent = Replace(strContent,"&frac12;","½"):strContent = Replace(strContent,"&Acirc;","Â"):strContent = Replace(strContent,"&Ccedil;","Ç"):strContent = Replace(strContent,"&Igrave;","Ì"):strContent = Replace(strContent,"&Ntilde;","Ñ"):strContent = Replace(strContent,"&Ouml;","Ö"):strContent = Replace(strContent,"&Ucirc;","Û"):strContent = Replace(strContent,"&agrave;","à"):strContent = Replace(strContent,"&aring;","å"):strContent = Replace(strContent,"&ecirc;","ê"):strContent = Replace(strContent,"&iuml;","ï"):strContent = Replace(strContent,"&ocirc;","ô"):strContent = Replace(strContent,"&ugrave;","ù"):strContent = Replace(strContent,"&thorn;","þ"):strContent = Replace(strContent,"&cent;","¢"):strContent = Replace(strContent,"&sect;","§"):strContent = Replace(strContent,"&not;","¬"):strContent = Replace(strContent,"&plusmn;","±"):strContent = Replace(strContent,"&para;","¶"):strContent = Replace(strContent,"&raquo;","»"):strContent = Replace(strContent,"&Agrave;","À"):strContent = Replace(strContent,"&Aring;","Å"):strContent = Replace(strContent,"&Ecirc;","Ê"):strContent = Replace(strContent,"&Iuml;","Ï"):strContent = Replace(strContent,"&Ocirc;","Ô"):strContent = Replace(strContent,"&Ugrave;","Ù"):strContent = Replace(strContent,"&THORN;","Þ"):strContent = Replace(strContent,"&atilde;","ã"):strContent = Replace(strContent,"&egrave;","è"):strContent = Replace(strContent,"&iacute;","í"):strContent = Replace(strContent,"&ograve;","ò"):strContent = Replace(strContent,"&divide;","÷"):strContent = Replace(strContent,"&uuml;","ü"):strContent = Replace(strContent,"&pound;","£"):strContent = Replace(strContent,"&uml;","¨"):strContent = Replace(strContent,"&shy;",""):strContent = Replace(strContent,"&sup2;","²"):strContent = Replace(strContent,"&middot;","·"):strContent = Replace(strContent,"&frac14;","¼"):strContent = Replace(strContent,"&Aacute;","Á"):strContent = Replace(strContent,"&AElig;","Æ"):strContent = Replace(strContent,"&Euml;","Ë"):strContent = Replace(strContent,"&ETH;","Ð"):strContent = Replace(strContent,"&Otilde;","Õ"):strContent = Replace(strContent,"&Uacute;","Ú"):strContent = Replace(strContent,"&szlig;","ß"):strContent = Replace(strContent,"&auml;","ä"):strContent = Replace(strContent,"&eacute;","é"):strContent = Replace(strContent,"&icirc;","î"):strContent = Replace(strContent,"&oacute;","ó"):strContent = Replace(strContent,"&oslash;","ø"):strContent = Replace(strContent,"&yacute;","ý"):strContent = Replace(strContent,"&OElig;","Œ"):strContent = Replace(strContent,"&oelig;","œ"):strContent = Replace(strContent,"&tilde;","˜"):strContent = Replace(strContent,"&zwj;","‍ "):strContent = Replace(strContent,"&lsquo;","‘"):strContent = Replace(strContent,"&bdquo;","„"):strContent = Replace(strContent,"&rsaquo;","›"):strContent = Replace(strContent,"&Scaron;","Š"):strContent = Replace(strContent,"&lrm;","‎ "):strContent = Replace(strContent,"&rsquo;","’"):strContent = Replace(strContent,"&dagger;","†"):strContent = Replace(strContent,"&euro;","€"):strContent = Replace(strContent,"&scaron;","š"):strContent = Replace(strContent,"&rlm;","‏  "):strContent = Replace(strContent,"&sbquo;","‚"):strContent = Replace(strContent,"&Dagger;","‡"):strContent = Replace(strContent,"&Yuml;","Ÿ"):strContent = Replace(strContent,"&thinsp;"," "):strContent = Replace(strContent,"&ndash;","–"):strContent = Replace(strContent,"&ldquo;","“"):strContent = Replace(strContent,"&permil;","‰"):strContent = Replace(strContent,"&circ;","ˆ"):strContent = Replace(strContent,"&zwnj;","‌ "):strContent = Replace(strContent,"&mdash;","—"):strContent = Replace(strContent,"&rdquo;","”"):strContent = Replace(strContent,"&lsaquo;","‹"):strContent = Replace(strContent,"&fnof;","ƒ"):strContent = Replace(strContent,"&Epsilon;","Ε"):strContent = Replace(strContent,"&Kappa;","Κ"):strContent = Replace(strContent,"&Omicron;","Ο"):strContent = Replace(strContent,"&Upsilon;","Υ"):strContent = Replace(strContent,"&alpha;","α"):strContent = Replace(strContent,"&zeta;","ζ"):strContent = Replace(strContent,"&lambda;","λ"):strContent = Replace(strContent,"&pi;","π"):strContent = Replace(strContent,"&upsilon;","υ"):strContent = Replace(strContent,"&thetasym;","?"):strContent = Replace(strContent,"&prime;","′"):strContent = Replace(strContent,"&image;","ℑ"):strContent = Replace(strContent,"&uarr;","↑"):strContent = Replace(strContent,"&lArr;","⇐"):strContent = Replace(strContent,"&forall;","∀"):strContent = Replace(strContent,"&isin;","∈"):strContent = Replace(strContent,"&minus;","−"):strContent = Replace(strContent,"&ang;","∠"):strContent = Replace(strContent,"&int;","∫"):strContent = Replace(strContent,"&ne;","≠"):strContent = Replace(strContent,"&sup;","⊃"):strContent = Replace(strContent,"&otimes;","⊗"):strContent = Replace(strContent,"&lfloor;","?"):strContent = Replace(strContent,"&spades;","♠"):strContent = Replace(strContent,"&Alpha;","Α"):strContent = Replace(strContent,"&Zeta;","Ζ"):strContent = Replace(strContent,"&Lambda;","Λ"):strContent = Replace(strContent,"&Pi;","Π"):strContent = Replace(strContent,"&Phi;","Φ"):strContent = Replace(strContent,"&beta;","β"):strContent = Replace(strContent,"&eta;","η"):strContent = Replace(strContent,"&mu;","μ"):strContent = Replace(strContent,"&rho;","ρ"):strContent = Replace(strContent,"&phi;","φ"):strContent = Replace(strContent,"&upsih;","?"):strContent = Replace(strContent,"&Prime;","″"):strContent = Replace(strContent,"&real;","ℜ"):strContent = Replace(strContent,"&rarr;","→"):strContent = Replace(strContent,"&uArr;","⇑"):strContent = Replace(strContent,"&part;","∂"):strContent = Replace(strContent,"&notin;","∉"):strContent = Replace(strContent,"&lowast;","∗"):strContent = Replace(strContent,"&and;","∧"):strContent = Replace(strContent,"&there4;","∴"):strContent = Replace(strContent,"&equiv;","≡"):strContent = Replace(strContent,"&nsub;","⊄"):strContent = Replace(strContent,"&perp;","⊥"):strContent = Replace(strContent,"&rfloor;","?"):strContent = Replace(strContent,"&clubs;","♣"):strContent = Replace(strContent,"&Beta;","Β"):strContent = Replace(strContent,"&Eta;","Η"):strContent = Replace(strContent,"&Mu;","Μ"):strContent = Replace(strContent,"&Rho;","Ρ"):strContent = Replace(strContent,"&Chi;","Χ"):strContent = Replace(strContent,"&gamma;","γ"):strContent = Replace(strContent,"&theta;","θ"):strContent = Replace(strContent,"&nu;","ν"):strContent = Replace(strContent,"&sigmaf;","ς"):strContent = Replace(strContent,"&chi;","χ"):strContent = Replace(strContent,"&piv;","?"):strContent = Replace(strContent,"&oline;","‾"):strContent = Replace(strContent,"&trade;","™"):strContent = Replace(strContent,"&darr;","↓"):strContent = Replace(strContent,"&rArr;","⇒"):strContent = Replace(strContent,"&exist;","∃"):strContent = Replace(strContent,"&ni;","∋"):strContent = Replace(strContent,"&radic;","√"):strContent = Replace(strContent,"&or;","∨"):strContent = Replace(strContent,"&sim;","∼"):strContent = Replace(strContent,"&le;","≤"):strContent = Replace(strContent,"&sube;","⊆"):strContent = Replace(strContent,"&sdot;","⋅"):strContent = Replace(strContent,"&lang;","?"):strContent = Replace(strContent,"&hearts;","♥"):strContent = Replace(strContent,"&Gamma;","Γ"):strContent = Replace(strContent,"&Theta;","Θ"):strContent = Replace(strContent,"&Nu;","Ν"):strContent = Replace(strContent,"&Sigma;","Σ"):strContent = Replace(strContent,"&Psi;","Ψ"):strContent = Replace(strContent,"&delta;","δ"):strContent = Replace(strContent,"&iota;","ι"):strContent = Replace(strContent,"&xi;","ξ"):strContent = Replace(strContent,"&sigma;","σ"):strContent = Replace(strContent,"&psi;","ψ"):strContent = Replace(strContent,"&bull;","•"):strContent = Replace(strContent,"&frasl;","⁄"):strContent = Replace(strContent,"&alefsym;","ℵ"):strContent = Replace(strContent,"&harr;","↔"):strContent = Replace(strContent,"&dArr;","⇓"):strContent = Replace(strContent,"&empty;","∅"):strContent = Replace(strContent,"&prod;","∏"):strContent = Replace(strContent,"&prop;","∝"):strContent = Replace(strContent,"&cap;","∩"):strContent = Replace(strContent,"&cong;","∝"):strContent = Replace(strContent,"&ge;","≥"):strContent = Replace(strContent,"&supe;","⊇"):strContent = Replace(strContent,"&lceil;","?"):strContent = Replace(strContent,"&rang;","?"):strContent = Replace(strContent,"&diams;","♦"):strContent = Replace(strContent,"&Delta;","Δ"):strContent = Replace(strContent,"&Iota;","Ι"):strContent = Replace(strContent,"&Xi;","Ξ"):strContent = Replace(strContent,"&Tau;","Τ"):strContent = Replace(strContent,"&Omega;","Ω"):strContent = Replace(strContent,"&epsilon;","ε"):strContent = Replace(strContent,"&kappa;","κ"):strContent = Replace(strContent,"&omicron;","ο"):strContent = Replace(strContent,"&tau;","τ"):strContent = Replace(strContent,"&omega;","ω"):strContent = Replace(strContent,"&hellip;","…"):strContent = Replace(strContent,"&weierp;","℘"):strContent = Replace(strContent,"&larr;","←"):strContent = Replace(strContent,"&crarr;","↵"):strContent = Replace(strContent,"&hArr;","⇔"):strContent = Replace(strContent,"&nabla;","∇"):strContent = Replace(strContent,"&sum;","∑"):strContent = Replace(strContent,"&infin;","∞"):strContent = Replace(strContent,"&cup;","∪"):strContent = Replace(strContent,"&asymp;","≈"):strContent = Replace(strContent,"&sub;","⊂"):strContent = Replace(strContent,"&oplus;","⊕"):strContent = Replace(strContent,"&rceil;","?"):strContent = Replace(strContent,"&loz;","◊"):strContent = Replace(strContent,"&#60;","<"):strContent = Replace(strContent,"&#62;",">"):strContent = Replace(strContent,"&#38;","&"):strContent = Replace(strContent,"&#34;",""""):strContent = Replace(strContent,"&#169;","©"):strContent = Replace(strContent,"&#174;","®"):strContent = Replace(strContent,"&#8482;","™"):strContent = Replace(strContent,"&#215;","×"):strContent = Replace(strContent,"&#247;","÷"):strContent = Replace(strContent,"&#160;",""):strContent = Replace(strContent,"&#165;","¥"):strContent = Replace(strContent,"&#170;","ª"):strContent = Replace(strContent,"&#175;","¯"):strContent = Replace(strContent,"&#180;","´"):strContent = Replace(strContent,"&#185;","¹"):strContent = Replace(strContent,"&#190;","¾"):strContent = Replace(strContent,"&#195;","Ã"):strContent = Replace(strContent,"&#200;","È"):strContent = Replace(strContent,"&#205;","Í"):strContent = Replace(strContent,"&#210;","Ò"):strContent = Replace(strContent,"&#215;","×"):strContent = Replace(strContent,"&#220;","Ü"):strContent = Replace(strContent,"&#225;","á"):strContent = Replace(strContent,"&#230;","æ"):strContent = Replace(strContent,"&#235;","ë"):strContent = Replace(strContent,"&#240;","ð"):strContent = Replace(strContent,"&#245;","õ"):strContent = Replace(strContent,"&#250;","ú"):strContent = Replace(strContent,"&#161;","¡"):strContent = Replace(strContent,"&#166;","¦"):strContent = Replace(strContent,"&#171;","«"):strContent = Replace(strContent,"&#176;","°"):strContent = Replace(strContent,"&#181;","µ"):strContent = Replace(strContent,"&#186;","º"):strContent = Replace(strContent,"&#191;","¿"):strContent = Replace(strContent,"&#196;","Ä"):strContent = Replace(strContent,"&#201;","É"):strContent = Replace(strContent,"&#206;","Î"):strContent = Replace(strContent,"&#211;","Ó"):strContent = Replace(strContent,"&#216;","Ø"):strContent = Replace(strContent,"&#221;","Ý"):strContent = Replace(strContent,"&#226;","â"):strContent = Replace(strContent,"&#231;","ç"):strContent = Replace(strContent,"&#236;","ì"):strContent = Replace(strContent,"&#241;","ñ"):strContent = Replace(strContent,"&#246;","ö"):strContent = Replace(strContent,"&#251;","û"):strContent = Replace(strContent,"&#164;","¤"):strContent = Replace(strContent,"&#169;","©"):strContent = Replace(strContent,"&#174;","®"):strContent = Replace(strContent,"&#179;","³"):strContent = Replace(strContent,"&#184;","¸"):strContent = Replace(strContent,"&#189;","½"):strContent = Replace(strContent,"&#194;","Â"):strContent = Replace(strContent,"&#199;","Ç"):strContent = Replace(strContent,"&#204;","Ì"):strContent = Replace(strContent,"&#209;","Ñ"):strContent = Replace(strContent,"&#214;","Ö"):strContent = Replace(strContent,"&#219;","Û"):strContent = Replace(strContent,"&#224;","à"):strContent = Replace(strContent,"&#229;","å"):strContent = Replace(strContent,"&#234;","ê"):strContent = Replace(strContent,"&#239;","ï"):strContent = Replace(strContent,"&#244;","ô"):strContent = Replace(strContent,"&#249;","ù"):strContent = Replace(strContent,"&#254;","þ"):strContent = Replace(strContent,"&#162;","¢"):strContent = Replace(strContent,"&#167;","§"):strContent = Replace(strContent,"&#172;","¬"):strContent = Replace(strContent,"&#177;","±"):strContent = Replace(strContent,"&#182;","¶"):strContent = Replace(strContent,"&#187;","»"):strContent = Replace(strContent,"&#192;","À"):strContent = Replace(strContent,"&#197;","Å"):strContent = Replace(strContent,"&#202;","Ê"):strContent = Replace(strContent,"&#207;","Ï"):strContent = Replace(strContent,"&#212;","Ô"):strContent = Replace(strContent,"&#217;","Ù"):strContent = Replace(strContent,"&#222;","Þ"):strContent = Replace(strContent,"&#227;","ã"):strContent = Replace(strContent,"&#232;","è"):strContent = Replace(strContent,"&#237;","í"):strContent = Replace(strContent,"&#242;","ò"):strContent = Replace(strContent,"&#247;","÷"):strContent = Replace(strContent,"&#252;","ü"):strContent = Replace(strContent,"&#163;","£"):strContent = Replace(strContent,"&#168;","¨"):strContent = Replace(strContent,"&#173;",""):strContent = Replace(strContent,"&#178;","²"):strContent = Replace(strContent,"&#183;","·"):strContent = Replace(strContent,"&#188;","¼"):strContent = Replace(strContent,"&#193;","Á"):strContent = Replace(strContent,"&#198;","Æ"):strContent = Replace(strContent,"&#203;","Ë"):strContent = Replace(strContent,"&#208;","Ð"):strContent = Replace(strContent,"&#213;","Õ"):strContent = Replace(strContent,"&#218;","Ú"):strContent = Replace(strContent,"&#223;","ß"):strContent = Replace(strContent,"&#228;","ä"):strContent = Replace(strContent,"&#233;","é"):strContent = Replace(strContent,"&#238;","î"):strContent = Replace(strContent,"&#243;","ó"):strContent = Replace(strContent,"&#248;","ø"):strContent = Replace(strContent,"&#253;","ý"):strContent = Replace(strContent,"&#338;","Œ"):strContent = Replace(strContent,"&#339;","œ"):strContent = Replace(strContent,"&#732;","˜"):strContent = Replace(strContent,"&#8205;","‍ "):strContent = Replace(strContent,"&#8216;","‘"):strContent = Replace(strContent,"&#8222;","„"):strContent = Replace(strContent,"&#8250;","›"):strContent = Replace(strContent,"&#352;","Š"):strContent = Replace(strContent,"&#8206;","‎ "):strContent = Replace(strContent,"&#8217;","’"):strContent = Replace(strContent,"&#8224;","†"):strContent = Replace(strContent,"&#8364;","€"):strContent = Replace(strContent,"&#353;","š"):strContent = Replace(strContent,"&#8207;","‏  "):strContent = Replace(strContent,"&#8218;","‚"):strContent = Replace(strContent,"&#8225;","‡"):strContent = Replace(strContent,"&#376;","Ÿ"):strContent = Replace(strContent,"&#8201;"," "):strContent = Replace(strContent,"&#8211;","–"):strContent = Replace(strContent,"&#8220;","“"):strContent = Replace(strContent,"&#8240;","‰"):strContent = Replace(strContent,"&#710;","ˆ"):strContent = Replace(strContent,"&#8204;","‌ "):strContent = Replace(strContent,"&#8212;","—"):strContent = Replace(strContent,"&#8221;","”"):strContent = Replace(strContent,"&#8249;","‹"):strContent = Replace(strContent,"&#402;","ƒ"):strContent = Replace(strContent,"&#917;","Ε"):strContent = Replace(strContent,"&#922;","Κ"):strContent = Replace(strContent,"&#927;","Ο"):strContent = Replace(strContent,"&#933;","Υ"):strContent = Replace(strContent,"&#945;","α"):strContent = Replace(strContent,"&#950;","ζ"):strContent = Replace(strContent,"&#955;","λ"):strContent = Replace(strContent,"&#960;","π"):strContent = Replace(strContent,"&#965;","υ"):strContent = Replace(strContent,"&#977;","?"):strContent = Replace(strContent,"&#8242;","′"):strContent = Replace(strContent,"&#8465;","ℑ"):strContent = Replace(strContent,"&#8593;","↑"):strContent = Replace(strContent,"&#8656;","⇐"):strContent = Replace(strContent,"&#8704;","∀"):strContent = Replace(strContent,"&#8712;","∈"):strContent = Replace(strContent,"&#8722;","−"):strContent = Replace(strContent,"&#8736;","∠"):strContent = Replace(strContent,"&#8747;","∫"):strContent = Replace(strContent,"&#8800;","≠"):strContent = Replace(strContent,"&#8835;","⊃"):strContent = Replace(strContent,"&#8855;","⊗"):strContent = Replace(strContent,"&#8970;","?"):strContent = Replace(strContent,"&#9824;","♠"):strContent = Replace(strContent,"&#913;","Α"):strContent = Replace(strContent,"&#918;","Ζ"):strContent = Replace(strContent,"&#923;","Λ"):strContent = Replace(strContent,"&#928;","Π"):strContent = Replace(strContent,"&#934;","Φ"):strContent = Replace(strContent,"&#946;","β"):strContent = Replace(strContent,"&#951;","η"):strContent = Replace(strContent,"&#956;","μ"):strContent = Replace(strContent,"&#961;","ρ"):strContent = Replace(strContent,"&#966;","φ"):strContent = Replace(strContent,"&#978;","?"):strContent = Replace(strContent,"&#8243;","″"):strContent = Replace(strContent,"&#8476;","ℜ"):strContent = Replace(strContent,"&#8594;","→"):strContent = Replace(strContent,"&#8657;","⇑"):strContent = Replace(strContent,"&#8706;","∂"):strContent = Replace(strContent,"&#8713;","∉"):strContent = Replace(strContent,"&#8727;","∗"):strContent = Replace(strContent,"&#8743;","∧"):strContent = Replace(strContent,"&#8756;","∴"):strContent = Replace(strContent,"&#8801;","≡"):strContent = Replace(strContent,"&#8836;","⊄"):strContent = Replace(strContent,"&#8869;","⊥"):strContent = Replace(strContent,"&#8971;","?"):strContent = Replace(strContent,"&#9827;","♣"):strContent = Replace(strContent,"&#914;","Β"):strContent = Replace(strContent,"&#919;","Η"):strContent = Replace(strContent,"&#924;","Μ"):strContent = Replace(strContent,"&#929;","Ρ"):strContent = Replace(strContent,"&#935;","Χ"):strContent = Replace(strContent,"&#947;","γ"):strContent = Replace(strContent,"&#952;","θ"):strContent = Replace(strContent,"&#957;","ν"):strContent = Replace(strContent,"&#962;","ς"):strContent = Replace(strContent,"&#967;","χ"):strContent = Replace(strContent,"&#982;","?"):strContent = Replace(strContent,"&#8254;","‾"):strContent = Replace(strContent,"&#8482;","™"):strContent = Replace(strContent,"&#8595;","↓"):strContent = Replace(strContent,"&#8658;","⇒"):strContent = Replace(strContent,"&#8707;","∃"):strContent = Replace(strContent,"&#8715;","∋"):strContent = Replace(strContent,"&#8730;","√"):strContent = Replace(strContent,"&#8744;","∨"):strContent = Replace(strContent,"&#8764;","∼"):strContent = Replace(strContent,"&#8804;","≤"):strContent = Replace(strContent,"&#8838;","⊆"):strContent = Replace(strContent,"&#8901;","⋅"):strContent = Replace(strContent,"&#9001;","?"):strContent = Replace(strContent,"&#9829;","♥"):strContent = Replace(strContent,"&#915;","Γ"):strContent = Replace(strContent,"&#920;","Θ"):strContent = Replace(strContent,"&#925;","Ν"):strContent = Replace(strContent,"&#931;","Σ"):strContent = Replace(strContent,"&#936;","Ψ"):strContent = Replace(strContent,"&#948;","δ"):strContent = Replace(strContent,"&#953;","ι"):strContent = Replace(strContent,"&#958;","ξ"):strContent = Replace(strContent,"&#963;","σ"):strContent = Replace(strContent,"&#968;","ψ"):strContent = Replace(strContent,"&#8226;","•"):strContent = Replace(strContent,"&#8260;","⁄"):strContent = Replace(strContent,"&#8501;","ℵ"):strContent = Replace(strContent,"&#8596;","↔"):strContent = Replace(strContent,"&#8659;","⇓"):strContent = Replace(strContent,"&#8709;","∅"):strContent = Replace(strContent,"&#8719;","∏"):strContent = Replace(strContent,"&#8733;","∝"):strContent = Replace(strContent,"&#8745;","∩"):strContent = Replace(strContent,"&#8773;","∝"):strContent = Replace(strContent,"&#8805;","≥"):strContent = Replace(strContent,"&#8839;","⊇"):strContent = Replace(strContent,"&#8968;","?"):strContent = Replace(strContent,"&#9002;","?"):strContent = Replace(strContent,"&#9830;","♦"):strContent = Replace(strContent,"&#916;","Δ"):strContent = Replace(strContent,"&#921;","Ι"):strContent = Replace(strContent,"&#926;","Ξ"):strContent = Replace(strContent,"&#932;","Τ"):strContent = Replace(strContent,"&#937;","Ω"):strContent = Replace(strContent,"&#949;","ε"):strContent = Replace(strContent,"&#954;","κ"):strContent = Replace(strContent,"&#959;","ο"):strContent = Replace(strContent,"&#964;","τ"):strContent = Replace(strContent,"&#969;","ω"):strContent = Replace(strContent,"&#8230;","…"):strContent = Replace(strContent,"&#8472;","℘"):strContent = Replace(strContent,"&#8592;","←"):strContent = Replace(strContent,"&#8629;","↵"):strContent = Replace(strContent,"&#8660;","⇔"):strContent = Replace(strContent,"&#8711;","∇"):strContent = Replace(strContent,"&#8721;","∑"):strContent = Replace(strContent,"&#8734;","∞"):strContent = Replace(strContent,"&#8746;","∪"):strContent = Replace(strContent,"&#8776;","≈"):strContent = Replace(strContent,"&#8834;","⊂"):strContent = Replace(strContent,"&#8853;","⊕"):strContent = Replace(strContent,"&#8969;","?"):strContent = Replace(strContent,"&#9674;","◊"):strContent = Replace(strContent,"&#12460;","ガ"):strContent = Replace(strContent,"&#12462;","ギ"):strContent = Replace(strContent,"&#12450;","ア"):strContent = Replace(strContent,"&#12466;","ゲ"):strContent = Replace(strContent,"&#12468;","ゴ"):strContent = Replace(strContent,"&#12470;","ザ"):strContent = Replace(strContent,"&#12472;","ジ"):strContent = Replace(strContent,"&#12474;","ズ"):strContent = Replace(strContent,"&#12476;","ゼ"):strContent = Replace(strContent,"&#12478;","ゾ"):strContent = Replace(strContent,"&#12480;","ダ"):strContent = Replace(strContent,"&#12482;","ヂ"):strContent = Replace(strContent,"&#12485;","ヅ"):strContent = Replace(strContent,"&#12487;","デ"):strContent = Replace(strContent,"&#12489;","ド"):strContent = Replace(strContent,"&#12496;","バ"):strContent = Replace(strContent,"&#12497;","パ"):strContent = Replace(strContent,"&#12499;","ビ"):strContent = Replace(strContent,"&#12500;","ピ"):strContent = Replace(strContent,"&#12502;","ブ"):strContent = Replace(strContent,"&#12502;","ブ"):strContent = Replace(strContent,"&#12503;","プ"):strContent = Replace(strContent,"&#12505;","ベ"):strContent = Replace(strContent,"&#12506;","ペ"):strContent = Replace(strContent,"&#12508;","ボ"):strContent = Replace(strContent,"&#12509;","ポ"):strContent = Replace(strContent,"&#12532;","ヴ"):ZBQQConnect_ReplaceXO=strContent::End Function

Function ZBQQConnect_SBar(Btype)
	dim b(3,3),i,j,k
	b(1,1)="m-left"
	b(1,2)="main.asp"
	b(1,3)="首页"
	b(2,1)="m-left"
	b(2,2)="m.asp"
	b(2,3)="绑定管理"
	b(3,1)="m-left"
	b(3,2)="setting.asp"
	b(3,3)="插件配置"
	For i=1 to 3
		if btype=i then
			k=k&"<span class=""" & b(i,1) & " m-now""><a href=""" & b(i,2) & """>" & b(i,3) & "</a></span>"
		else
			k=k&"<span class=""" & b(i,1) & """><a href=""" & b(i,2) & """>" & b(i,3) & "</a></span>"
		end if
	Next
	k=k&"<script type=""text/javascript"">ActiveLeftMenu(""aPlugInMng"");</script>"
	ZBQQConnect_SBar=k
End Function
%>