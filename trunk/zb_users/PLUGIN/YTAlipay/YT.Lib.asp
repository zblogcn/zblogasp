<!--#include file="YT.MD5.asp" -->
<%
Class YT_Alipay
	Public sPartner
	Public sKey
	Public sSeller_Email
	Public sInput_Charset
	Public sStatus
	Private sNotify_Url
	Private sReturn_Url
	Private sPost_Url
	Private sResult_Url
	Private sSign_Type
	Private sTransAction_Type
	Private sPageSize
	Private oTConfig
	Private Para()
	Private objArticle,YTModelXML,YT_ALIPAY_TYPE
	Public Property Let Partner(s)
		sPartner = s
	End Property
	Public Property Let Key(s)
		sKey = s
	End Property
	Public Property Let Seller_Email(s)
		sSeller_Email = s
	End Property
	
	Private Sub Class_Initialize()
		sPageSize = 18
		Set oTConfig = new TConfig
			With oTConfig
				.Load "YTAlipay"
				sPartner = .Read("PARTNER")
				sKey = .Read("KEY")
				sSeller_Email = .Read("SELLER_EMAIL")
				sInput_Charset = .Read("INPUT_CHARSET")
				sNotify_Url = .Read("NOTIFY_URL")
				sReturn_Url = .Read("RETURN_URL")
				sPost_Url = .Read("POST_URL")
				sResult_Url = .Read("RESULT_URL")
				sSign_Type = .Read("SIGN_TYPE")
			End With
		Redim Para(-1)
		AddPara "partner",sPartner
		AddPara "_input_charset",sInput_Charset
		AddPara "notify_url",sNotify_Url
		AddPara "return_url",sReturn_Url
		AddPara "payment_type",1
		AddPara "seller_email",sSeller_Email
		Set objArticle = new TArticle
		Set YTModelXML = new YT_Model_XML
		sTransAction_Type = Array("CREATE_PARTNER_TRADE_BY_BUYER","CREATE_DIRECT_PAY_BY_USER","TRADE_CREATE_BY_BUYER")
		sStatus = Array(Array("WAIT_BUYER_PAY","WAIT_SELLER_SEND_GOODS","WAIT_BUYER_CONFIRM_GOODS","TRADE_FINISHED","TRADE_CLOSED","WAIT_SELLER_AGREE","SELLER_REFUSE_BUYER","WAIT_BUYER_RETURN_GOODS","WAIT_SELLER_CONFIRM_GOODS","REFUND_SUCCESS","REFUND_CLOSED"),Array("WAIT_BUYER_PAY","TRADE_FINISHED","TRADE_SUCCESS","TRADE_CLOSED","REFUND_SUCCESS","REFUND_CLOSED"))
	End Sub
	Private Sub Class_Terminate()
		Set oTConfig = Nothing
		Set objArticle = Nothing
		Set YTModelXML = Nothing
	End Sub
	Sub Save()
		With oTConfig
			.Write "PARTNER",sPartner
			.Write "KEY",sKey
			.Write "SELLER_EMAIL",sSeller_Email
			.Write "POST_URL","https://www.alipay.com/cooperate/gateway.do?"
			.Write "RESULT_URL","http://notify.alipay.com/trade/notify_query.do?"
			.Write "INPUT_CHARSET","utf-8"
			.Write "SIGN_TYPE","MD5"
			.Write "RETURN_URL",ZC_BLOG_HOST&"Alipay/Return/"
			.Write "NOTIFY_URL",ZC_BLOG_HOST&"Alipay/Notify/"
			.Save
		End With
	End Sub
	Sub Install()
		Dim x,j,f,d,s
		Set x=CreateObject("Microsoft.XMLDOM")
			x.load(BlogPath&"ZB_USERS/PLUGIN/YTAlipay/alipay")
			If x.readyState=4 Then
				Set j=x.selectNodes("//folder/path")
					Set f=CreateObject("Scripting.FileSystemObject")
						For Each d In j
							If f.FolderExists(Server.MapPath("/"&d.Text))=False Then f.CreateFolder(Server.MapPath("/"&d.Text))
						Next
					Set f=Nothing
				Set j=Nothing
				Set j=x.selectNodes("//file/path")
					For Each d In j
						Set s=CreateObject("ADODB.Stream")
							s.Type=1
							s.Open
							s.Write d.nextSibling.nodeTypedvalue
							s.SaveToFile Server.MapPath("/"&d.Text),2
							s.Close
						Set s=Nothing
					Next
				Set j=Nothing
			End If
		Set x=Nothing
	End Sub
	Sub UnInstall()
		Dim x,j,f,d
		Set x=Server.CreateObject("Microsoft.XMLDOM")
			x.load(BlogPath&"ZB_USERS/PLUGIN/YTAlipay/alipay")
			If x.readyState=4 Then
				Set j=x.selectNodes("//folder/path")
					Set f=CreateObject("Scripting.FileSystemObject")
						For Each d In j
							If f.FolderExists(Server.MapPath("/"&d.Text))=True Then f.DeleteFolder(Server.MapPath("/"&d.Text))
						Next
					Set f=Nothing
				Set j=Nothing
			End If
		Set x=Nothing
	End Sub
	Sub AddPara(ByVal Item,ByVal Value)
		Dim i:i = UBound(Para) + 1
		Redim Preserve Para(i)
		Para(i) = Item&"="&Value
	End Sub
	Function GetAlipayType(ID)
		GetAlipayType=Empty
		If objArticle.LoadInfoByID(ID) Then
			Dim Node,Json,Object,i,s
			Set Node = YTModelXML.GetModel(objArticle.CateID)
				If Not Node Is Nothing Then
					Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,objArticle.ID)
					If isEmpty(Json) Then Exit Function
					Set Object = YT.eval(Json)
						If GetFieldValue(Object,"YT_Money") Then
							s = GetFieldValue(Object,"YT_Service")
							If Not isEmpty(s) Then
								For i=LBound(sTransAction_Type) To UBound(sTransAction_Type)
									If UCase(sTransAction_Type(i)) = UCase(s) Then
										GetAlipayType = i
										Exit For
									End If
								Next
							End If
						End If
					Set Object = Nothing
				End If
			Set Node = Nothing
		End If
	End Function
	Function GetParaValue(ByVal Item)
		Dim i,iPos,nLen
		For i = LBound(Para) To UBound(Para)
			iPos = Instr(Para(i),"=")
			nLen = Len(Para(i))
			If LCase(Left(Para(i),iPos-1)) = LCase(Item) Then
				GetParaValue = Right(Para(i),nLen-iPos)
				Exit For
			End If
		Next
	End Function
	Function Buy()
		If oTConfig.Exists("Key") Then
			Dim id
			For Each id in Split(Request.Form("YT_ID"),",")
				Call CheckParameter(id,"int",0)
				YT_ALIPAY_TYPE=GetAlipayType(id)
				If Not IsEmpty(YT_ALIPAY_TYPE) Then
					Select Case YT_ALIPAY_TYPE
						Case 0:Buy=Create_partner_trade_by_buyer
						Case 1:Buy=Create_direct_pay_by_user
						Case Else:Buy=Create_partner_trade_by_buyer
					End Select
				End If
				Exit For
			Next
		Else
			Response.Write("支付宝未进行配置")
		End If
	End Function
	'即时交易函数
	Function Create_direct_pay_by_user()
		Dim body,paymethod,id
		Dim Json,Object,Field,Node
		id = Request.Form("YT_ID")
		Call CheckParameter(id,"int",0)
			If Not objArticle.LoadInfoByID(id) Then Exit Function
			Set Node = YTModelXML.GetModel(objArticle.CateID)
				If Node Is Nothing Then Exit Function
				Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,objArticle.ID)
				Set Object = YT.eval(Json)
					AddPara "service",LCase(sTransAction_Type(YT_ALIPAY_TYPE))
					AddPara "total_fee",FormatCurrency(GetFieldValue(Object,"YT_Money"),2)
				Set Object = Nothing
			Set Node = Nothing
			body = FilterSQL(Request.Form("YT_Body"))
			paymethod = FilterSQL(Request.Form("YT_Paymethod"))
			AddPara "show_url",ZC_BLOG_HOST&"zb_system/view.asp?id="&id
			AddPara "subject",objArticle.Title
			AddPara "body",body								'备注
			AddPara "out_trade_no",GetDateTime
		If paymethod = "directPay" Then
			AddPara "paymethod","directPay"
			AddPara "defaultbank",""
		Else
			AddPara "paymethod","bankPay"
			AddPara "defaultbank",paymethod					'支付方式
		End If
		Dim Sql,FieldName,FieldValue,j(),k,x,jsonBody
			Redim j(-1)
			Call ArrayPreserve(j,"简介",body)
			For Each k In GetRequestPost
				x=Split(k,"=")
				If inStr(x(0),"YT_")=0 Then Call ArrayPreserve(j,FilterSQL(x(0)),FilterSQL(x(1)))
			Next
			
			Call ArrayPreserve(j,"交易类型",YT_ALIPAY_TYPE)
			jsonBody="["&Join(j,",")&"]"
			FieldName = Array("[OrderID]","[OrderName]","[Body]","[Status]","[Service]","[Time]","[log_ID]")
			FieldValue = Array(Chr(39)&GetDateTime&Chr(39),Chr(39)&objArticle.Title&Chr(39),Chr(39)&jsonBody&Chr(39),0,YT_ALIPAY_TYPE,Chr(39)&Now()&Chr(39),id)
			Sql = "INSERT INTO [YT_Alipay]("
			Sql = Sql & Join(FieldName,",") & ") VALUES ("&Join(FieldValue,",")&")"
			objConn.Execute(Sql)
		Create_direct_pay_by_user = BuildFormHtml(Para, sKey, sSign_Type, sInput_Charset, sPost_Url, "POST")
	End Function
	Function ArrayPreserve(Byref a,Byval t,Byval v)
		Dim j
		j=UBound(a)+1
		Redim Preserve a(j)
		a(j)="{Text:"&Chr(34)&YT.escape(t)&Chr(34)&",Value:"&Chr(34)&YT.escape(v)&Chr(34)&"}"
	End Function
	'担保交易函数
	Function Create_partner_trade_by_buyer()
		Dim body,quantity
		dim logistics_type,logistics_payment
		Dim receive_name,receive_address,receive_mobile,receive_phone,receive_zip
		Dim id,Json,Object,Node
		Dim i,arrID,arrQuantity,service,price,logistics_fee,subject
			arrQuantity = Request.Form("YT_Quantity")
			arrQuantity = Split(arrQuantity,",")
			arrID = Request.Form("YT_ID")
			arrID = Split(arrID,",")
			price = 0
			For i = LBound(arrID) To UBound(arrID)
				id=arrID(i)
				Call CheckParameter(id,"int",0)
				quantity = arrQuantity(i)
				Call CheckParameter(quantity,"int",1)
				If Not objArticle.LoadInfoByID(id) Then Exit Function
				Set Node = YTModelXML.GetModel(objArticle.CateID)
					If Node Is Nothing Then Exit Function
					Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,objArticle.ID)
					Set Object = YT.eval(Json)
						service = LCase(sTransAction_Type(YT_ALIPAY_TYPE))
						logistics_type = GetFieldValue(Object,"YT_Logistics_Type")
						price = price + (quantity*FormatCurrency(GetFieldValue(Object,"YT_Money"),2))
						logistics_fee = GetFieldValue(Object,"YT_Logistics_Fee")
						logistics_payment = GetFieldValue(Object,"YT_Logistics_Payment")
						subject = subject & objArticle.Title
					Set Object = Nothing
				Set Node = Nothing
			Next
			AddPara "service",service
			AddPara "logistics_type",logistics_type
			AddPara "price",price
			AddPara "logistics_fee",logistics_fee
			AddPara "logistics_payment",logistics_payment
			AddPara "subject",subject
			
			body = FilterSQL(Request.Form("YT_Body"))
			receive_name = FilterSQL(Request.Form("YT_Receive_Name"))
			receive_address = FilterSQL(Request.Form("YT_Receive_Address"))
			receive_mobile = FilterSQL(Request.Form("YT_Receive_Mobile"))
			receive_phone = FilterSQL(Request.Form("YT_Receive_Phone"))
			receive_zip = FilterSQL(Request.Form("YT_Receive_Zip"))
			AddPara "body",body									'备注
			AddPara "out_trade_no",GetDateTime
			AddPara "quantity",1								'商品数量
			AddPara "receive_name",receive_name					'姓名
			AddPara "receive_address",receive_address			'地址
			AddPara "receive_mobile",receive_mobile				'手机
			AddPara "receive_phone",receive_phone				'电话
			AddPara "receive_zip",receive_zip					'邮编
			AddPara "show_url",ZC_BLOG_HOST&"zb_system/view.asp?id="&id
		Dim Sql,FieldName,FieldValue,j(),k,x,jsonBody
			Redim j(-1)
			Call ArrayPreserve(j,"物流费用",GetParaValue("logistics_fee"))
			Call ArrayPreserve(j,"物流支付类型",logistics_payment)
			Call ArrayPreserve(j,"物流类型",logistics_type)
			Call ArrayPreserve(j,"商品数量",quantity)
			Call ArrayPreserve(j,"姓名",receive_name)
			Call ArrayPreserve(j,"地址",receive_address)
			Call ArrayPreserve(j,"手机",receive_mobile)
			Call ArrayPreserve(j,"电话",receive_phone)
			Call ArrayPreserve(j,"邮编",receive_zip)
			For Each k In GetRequestPost
				x=Split(k,"=")
				If inStr(x(0),"YT_")=0 Then Call ArrayPreserve(j,FilterSQL(x(0)),FilterSQL(x(1)))
			Next
			Call ArrayPreserve(j,"交易类型",YT_ALIPAY_TYPE)
			jsonBody="["&Join(j,",")&"]"
			'jsonBody=Replace(jsonBody,Chr(34),Chr(34)&Chr(34))
			FieldName = Array("[OrderID]","[OrderName]","[Body]","[Status]","[Service]","[Time]","[log_ID]")
			FieldValue = Array(Chr(39)&GetDateTime&Chr(39),Chr(39)&objArticle.Title&Chr(39),Chr(39)&jsonBody&Chr(39),0,YT_ALIPAY_TYPE,Chr(39)&Now()&Chr(39),id)
			Sql = "INSERT INTO [YT_Alipay]("
			Sql = Sql & Join(FieldName,",") & ")VALUES ("&Join(FieldValue,",")&")"
			objConn.Execute(Sql)
		Create_partner_trade_by_buyer = BuildFormHtml(Para, sKey, sSign_Type, sInput_Charset, sPost_Url, "POST")
	End Function
	Function GetStatus(Byval arrObject,Byval Key)
		Dim i
		If isArray(arrObject) Then
			For i = LBound(arrObject) To UBound(arrObject)
				If UCase(arrObject(i)) = UCase(Key) Then
					GetStatus = i
					Exit Function
				End If
			Next
		End If
		GetStatus = 0
	End Function
	Function GetJsonList(byVal intPage)
		Dim objRS,OrderID,OrderName,Status,log_ID,Time
		Dim intPageCount,i,j,jsonText,Field,Fields,R()
			OrderID = FilterSQL(Request.Form("OrderID"))
			OrderName = FilterSQL(Request.Form("OrderName"))
			Status = Request.Form("Status")
			log_ID = Request.Form("log_ID")
			Time = Request.Form("Time")
		Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			objRS.Source="SELECT * FROM [YT_Alipay] WHERE (1=1)"
			If OrderID <> "" Then objRS.Source=objRS.Source & "AND([OrderID]='"&OrderID&"')"
			If OrderName <> "" Then objRS.Source=objRS.Source & "AND([OrderName] LIKE '%"&OrderName&"%')"
			If Not IsEmpty(Status) And IsNumeric(Status) Then objRS.Source=objRS.Source & "AND([Status]="&Status&")"
			If Not IsEmpty(log_ID) And IsNumeric(log_ID) Then objRS.Source=objRS.Source & "AND([log_ID]="&log_ID&")"
			If Not IsEmpty(Time) And IsDate(Time) Then objRS.Source=objRS.Source & "AND([Time]='"&Time&"')"
			objRS.Source=objRS.Source & "ORDER BY [Time] DESC,[ID] DESC"
			objRS.Open()
			If (Not objRS.bof) And (Not objRS.eof) Then
				objRS.PageSize = sPageSize
				intPageCount=objRS.PageCount
				objRS.AbsolutePage = intPage
				For i = 1 To objRS.PageSize
					Redim R(-1)
					jsonText = jsonText & "{"
					For Each Field In objRS.Fields
						j = Ubound(R) + 1
						ReDim Preserve R(j)
						If isDate(objRS(Field.Name)) Then
							R(j) = Field.Name&":"&Chr(34)&objRS(Field.Name)&Chr(34)
						Else
							R(j) = Field.Name&":"&Chr(34)&YT.escape(objRS(Field.Name))&Chr(34)
						End If
					Next
					jsonText = jsonText & Join(R,",") & "},"
					objRS.MoveNext
					If objRS.EOF Then Exit For
				Next
				If InStr(jsonText,",") > 0 Then 
					jsonText = Left(jsonText,Len(jsonText) - 1)
					GetJsonList = "{intPage:"&intPage&",pageCount:"&intPageCount&",objRow:["&jsonText&"]}"
				End If
			End If
			objRS.Close()
		Set objRS = Nothing
	End Function
	Function GetFieldValue(Byref Object,FieldName)
		Dim Field
		For Each Field In Object.YTARRAY
			If UCase(Field) = UCase(FieldName) Then
				Execute("GetFieldValue=Object."&Field)
				Exit Function
			End If
		Next
	End Function
	Sub DelAlipayOrder(ID)
		Dim Sql
		Sql = "DELETE FROM [YT_Alipay] WHERE [ID] = "& ID
		objConn.Execute(Sql)
	End Sub
	Private Function BuildRequestPara(sParaTemp, key, sign_type, input_charset)
		Dim mysign,sPara,sParaSort,nCount
		sPara = FilterPara(sParaTemp)
		sParaSort = SortPara(sPara)
		mysign = BuildMysign(sParaSort, key, sign_type, input_charset)
		nCount = ubound(sParaSort)
		Redim Preserve sParaSort(nCount+1)
		sParaSort(nCount+1) = "sign="&mysign
		Redim Preserve sParaSort(nCount+2)
		sParaSort(nCount+2) = "sign_type="&sign_type
		BuildRequestPara = sParaSort
	End Function

	Function FilterPara(sPara)
		Dim sParaFilter(),nCount,j,i,pos,nLen
		Dim itemName,itemValue
		nCount = ubound(sPara)
		j = 0
		For i = 0 To nCount
			pos = Instr(sPara(i),"=")
			nLen = Len(sPara(i))
			itemName = left(sPara(i),pos-1)
			itemValue = right(sPara(i),nLen-pos)
			
			If itemName <> "sign" And itemName <> "sign_type" And itemValue <> "" and isnull(itemValue) = false Then
				Redim Preserve sParaFilter(j)
				sParaFilter(j) = sPara(i)
				j = j + 1
			End If
		Next
		
		FilterPara = sParaFilter
	End Function

	Function SortPara(sPara)
		Dim nCount,i,minmax,minmaxSlot
		Dim mark,j,temp
		nCount = ubound(sPara)
		For i = nCount To 0 Step -1
			minmax = sPara( 0 )
			minmaxSlot = 0
			For j = 1 To i
				mark = (sPara( j ) > minmax)
				If mark Then 
					minmax = sPara( j )
					minmaxSlot = j
				End If
			Next
			If minmaxSlot <> i Then 
				temp = sPara( minmaxSlot )
				sPara( minmaxSlot ) = sPara( i )
				sPara( i ) = temp
			End If
		Next
		SortPara = sPara
	End Function
	
	Function BuildMysign(sPara, key, sign_type,input_charset)
		Dim prestr,mysign,nLen
		prestr = CreateLinkstring(sPara)
		prestr = prestr & key
		mysign = Sign(prestr,sign_type,input_charset)
	
		BuildMysign = mysign
	End Function

	Function CreateLinkstring(sPara)
		Dim nCount,i
		nCount = ubound(sPara)
		Dim prestr
		For i = 0 To nCount
			If i = nCount Then
				prestr = prestr & sPara(i)
			Else
				prestr = prestr & sPara(i) & "&"
			End if
		Next
		
		CreateLinkstring = prestr
	End Function

	Function Sign(prestr,sign_type,input_charset)
		Dim sResult
		If sign_type = "MD5" Then
			sResult = new Alipay_Md5.MD5(prestr,input_charset)
		Else 
			sResult = ""
		End If
		Sign = sResult
	End Function

	Function GetDateTime()
		Dim sTime,sResult
		sTime=now()
		sResult	= year(sTime)&right("0" & month(sTime),2)&right("0" & day(sTime),2)&right("0" & hour(sTime),2)&right("0" & minute(sTime),2)&right("0" & second(sTime),2)
		GetDateTime = sResult
	End Function

	function notify_verify()
		Dim responseTxt,sGetArray
		responseTxt = get_http()
		
		sGetArray = GetRequestPost()
	
		if IsArray(sGetArray) then
			Dim sArray,sort_para,mysign

			sArray = FilterPara(sGetArray)
			sort_para = SortPara(sArray)
			mysign  = BuildMysign(sort_para,sKey,sSign_Type,sInput_Charset)

			if mysign = request.Form("sign") and responseTxt = "true" then
				notify_verify = true
			else
				notify_verify = false
			end if
		else
			notify_verify = false
		end if
	end function
	
	function return_verify()
		Dim responseTxt,sGetArray
		responseTxt = get_http()
	
		sGetArray = GetRequestGet()
	
		if IsArray(sGetArray) then
			Dim sArray,sort_para,mysign

			sArray = FilterPara(sGetArray)
			sort_para = SortPara(sArray)
			mysign  = BuildMysign(sort_para,sKey,sSign_Type,sInput_Charset)

			if mysign = request.QueryString("sign") and responseTxt = "true" then
				return_verify = true
			else
				return_verify = false
			end if
		else
			return_verify = false
		end if
	end function

	function GetRequestGet()
		dim sArray()
		Dim i,varItem
		i = 0
		For Each varItem in Request.QueryString
			Redim Preserve sArray(i)
			sArray(i) = varItem&"="&Request(varItem) 
			i = i + 1
		Next 
		
		if i = 0 then
			GetRequestGet = ""
		else
			GetRequestGet = sArray
		end if
		
	end function

	function GetRequestPost()
		dim sArray()
		Dim i,varItem
		i = 0
		For Each varItem in Request.Form
			Redim Preserve sArray(i)
			sArray(i) = varItem&"="&Request.Form(varItem) 
			i = i + 1
		Next 
		
		if i = 0 then
			GetRequestPost = ""
		else
			GetRequestPost = sArray
		end if
	end function
	
	function get_http()
		Dim gateway,Retrieval,ResponseTxt
		gateway = sResult_Url &"partner=" & sPartner & "&notify_id=" & request("notify_id")
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		Retrieval.open "GET", gateway, False, "", ""
		Retrieval.send()
		ResponseTxt = Retrieval.ResponseText
		Set Retrieval = Nothing
		get_http = ResponseTxt
	end function

	Function Send_goods_confirm_by_platform(sParaTemp,sParaNode)
		Dim sParaXml
		sParaXml = SendGetInfo(sParaTemp, sKey, sSign_Type, sInput_Charset, "https://mapi.alipay.com/gateway.do?", sParaNode)
		Send_goods_confirm_by_platform = sParaXml
	End Function
	
	Function SendGetInfo(sParaTemp, key, sign_type, input_charset, gateway, sParaNode)
		Dim sUrl, objHttp, objXml, nCount, sParaXml()
		nCount = ubound(sParaNode)
		Dim sRequestData
		sRequestData = BuildRequestParaToString(sParaTemp, key, sign_type, input_charset)
		sUrl = gateway & sRequestData

		Set objHttp=Server.CreateObject("Microsoft.XMLHTTP")
		objHttp.open "GET", sUrl, False, "", ""
		objHttp.send()
		Set objXml=Server.CreateObject("Microsoft.XMLDOM")
		objXml.Async=true
		objXml.ValidateOnParse=False
		objXml.Load(objHttp.ResponseXML)
		Set objHttp = Nothing
		Dim objXmlData,i
		set objXmlData = objXml.getElementsByTagName("alipay").item(0)
		If Isnull(objXmlData.selectSingleNode("is_success")) Then
			Redim Preserve sParaXml(1)
			sParaXml(0) = "错误：非法XML格式数据"
		Else
			If objXmlData.selectSingleNode("is_success").text = "T" Then
				Redim sParaXml(-1)
				Dim j
				For i = 0 To nCount
					j=UBound(sParaXml)+1
					Redim Preserve sParaXml(j)
					sParaXml(j) = objXmlData.selectSingleNode(sParaNode(i)).text
				Next
			Else
				Redim Preserve sParaXml(1)
				sParaXml(0) = "错误："&objXmlData.selectSingleNode("error").text
			End If
		End If
		
		SendGetInfo = sParaXml
	End Function
	
	Private Function BuildRequestParaToString(sParaTemp, key, sign_type, input_charset)
		Dim sRequestData,sPara
		sPara = BuildRequestPara(sParaTemp, key, sign_type, input_charset)
		sRequestData = CreateLinkStringUrlEncode(sPara)
		
		BuildRequestParaToString = sRequestData
	End Function
	
	function CreateLinkStringUrlEncode(sPara)
		Dim nCount
		nCount = ubound(sPara)
		dim prestr,i,pos,nLen,itemName,itemValue
		for i = 0 to nCount
			pos = Instr(sPara(i),"=")
			nLen = Len(sPara(i))
			itemName = left(sPara(i),pos-1)
			itemValue = right(sPara(i),nLen-pos)
			
			if itemName <> "service" and itemName <> "_input_charset" then
				prestr = prestr & itemName &"=" & server.URLEncode(itemValue) & "&"
			else
				prestr = prestr & sPara(i) & "&"
			end if
		next
		CreateLinkStringUrlEncode = prestr
	end function

	Function BuildFormHtml(sParaTemp, key, sign_type, input_charset, gateway, sMethod)
		Dim sHtml, nCount,sPara,i
		Dim iPos,nLen,sItemName,sItemValue
		sPara = BuildRequestPara(sParaTemp, key, sign_type, input_charset)
		sHtml = "<form id='alipaysubmit' name='alipaysubmit' action='"& gateway &"_input_charset="&input_charset&"' method='"&sMethod&"'>"
		nCount = ubound(sPara)
		For i = 0 To nCount
			iPos = Instr(sPara(i),"=")
			nLen = Len(sPara(i))
			sItemName = left(sPara(i),iPos-1)
			sItemValue = right(sPara(i),nLen-iPos)
			sHtml = sHtml &"<input type='hidden' name='"& sItemName &"' value='"& sItemValue &"'/><br />"
		next
		sHtml = sHtml & "</form>"
		sHtml = sHtml & "<script>document.forms['alipaysubmit'].submit();</script>"
		BuildFormHtml = sHtml
	End Function
End Class
%>