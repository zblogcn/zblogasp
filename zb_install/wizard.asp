<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    wizard.asp
'// 开始时间:    2006.08.12
'// 最后修改:    
'// 备    注:    第一次使用时的向导页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<%
Const ZC_WD_MSG001="欢迎使用Z-Blog,在完成以下的设置后就可以开始你的BLOG之旅!"
Const ZC_WD_MSG002="BLOG的网络地址"
Const ZC_WD_MSG003="数据库的名称和地址"
Const ZC_WD_MSG004="管理员的名称"
Const ZC_WD_MSG005="密码"
Const ZC_WD_MSG006="密码确认"
Const ZC_WD_MSG007="BLOG唯一标识符"
Const ZC_WD_MSG008="系统自动随机生成"
Const ZC_WD_MSG009="Z-Blog安装向导"
Const ZC_WD_MSG010="设置完成!"
Const ZC_WD_MSG011="回到首页"
Const ZC_WD_MSG012="提交"
Const ZC_WD_MSG013="网址设置不正确"
Const ZC_WD_MSG014="用户名设置不正确"
Const ZC_WD_MSG015="密码为6位或更长"
Const ZC_WD_MSG016="请确认密码"
Const ZC_WD_MSG017="verify校验值不正确!"
Const ZC_WD_MSG018="data/zblog.mdb数据库不存在!"
Const ZC_WD_MSG019="或是"
Const ZC_WD_MSG020="登陆后台"

'--------------------------------------------------------------------
Const adOpenForwardOnly=0
Const adOpenKeyset=1
Const adOpenDynamic=2
Const adOpenStatic=3

Const adLockReadOnly=1
Const adLockPessimistic=2
Const adLockOptimistic=3
Const adLockBatchOptimistic=4

Const ForReading=1
Const ForWriting=2
Const ForAppending=8

Const adTypeBinary=1
Const adTypeText=2

Const adModeRead=1
Const adModeReadWrite=3

Const adSaveCreateNotExist=1
Const adSaveCreateOverWrite=2
'--------------------------------------------------------------------




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    
'*********************************************************
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)
 
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)
    
    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function F(x, y, z)
    F = (x And y) Or ((Not x) And z)
End Function

Private Function G(x, y, z)
    G = (x And z) Or (y And (Not z))
End Function

Private Function H(x, y, z)
    H = (x Xor y Xor z)
End Function

Private Function I(x, y, z)
    I = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount
    
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Public Function MD5(sMessage)
    Dim x
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim d
    
    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    x = ConvertToWordArray(sMessage)
    
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d
    
        FF a, b, c, d, x(k + 0), S11, &HD76AA478
        FF d, a, b, c, x(k + 1), S12, &HE8C7B756
        FF c, d, a, b, x(k + 2), S13, &H242070DB
        FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
        FF d, a, b, c, x(k + 5), S12, &H4787C62A
        FF c, d, a, b, x(k + 6), S13, &HA8304613
        FF b, c, d, a, x(k + 7), S14, &HFD469501
        FF a, b, c, d, x(k + 8), S11, &H698098D8
        FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
        FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        FF b, c, d, a, x(k + 11), S14, &H895CD7BE
        FF a, b, c, d, x(k + 12), S11, &H6B901122
        FF d, a, b, c, x(k + 13), S12, &HFD987193
        FF c, d, a, b, x(k + 14), S13, &HA679438E
        FF b, c, d, a, x(k + 15), S14, &H49B40821
    
        GG a, b, c, d, x(k + 1), S21, &HF61E2562
        GG d, a, b, c, x(k + 6), S22, &HC040B340
        GG c, d, a, b, x(k + 11), S23, &H265E5A51
        GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        GG a, b, c, d, x(k + 5), S21, &HD62F105D
        GG d, a, b, c, x(k + 10), S22, &H2441453
        GG c, d, a, b, x(k + 15), S23, &HD8A1E681
        GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
        GG d, a, b, c, x(k + 14), S22, &HC33707D6
        GG c, d, a, b, x(k + 3), S23, &HF4D50D87
        GG b, c, d, a, x(k + 8), S24, &H455A14ED
        GG a, b, c, d, x(k + 13), S21, &HA9E3E905
        GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        GG c, d, a, b, x(k + 7), S23, &H676F02D9
        GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            
        HH a, b, c, d, x(k + 5), S31, &HFFFA3942
        HH d, a, b, c, x(k + 8), S32, &H8771F681
        HH c, d, a, b, x(k + 11), S33, &H6D9D6122
        HH b, c, d, a, x(k + 14), S34, &HFDE5380C
        HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
        HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
        HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
        HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
        HH a, b, c, d, x(k + 13), S31, &H289B7EC6
        HH d, a, b, c, x(k + 0), S32, &HEAA127FA
        HH c, d, a, b, x(k + 3), S33, &HD4EF3085
        HH b, c, d, a, x(k + 6), S34, &H4881D05
        HH a, b, c, d, x(k + 9), S31, &HD9D4D039
        HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
        HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
        HH b, c, d, a, x(k + 2), S34, &HC4AC5665
    
        II a, b, c, d, x(k + 0), S41, &HF4292244
        II d, a, b, c, x(k + 7), S42, &H432AFF97
        II c, d, a, b, x(k + 14), S43, &HAB9423A7
        II b, c, d, a, x(k + 5), S44, &HFC93A039
        II a, b, c, d, x(k + 12), S41, &H655B59C3
        II d, a, b, c, x(k + 3), S42, &H8F0CCC92
        II c, d, a, b, x(k + 10), S43, &HFFEFF47D
        II b, c, d, a, x(k + 1), S44, &H85845DD1
        II a, b, c, d, x(k + 8), S41, &H6FA87E4F
        II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        II c, d, a, b, x(k + 6), S43, &HA3014314
        II b, c, d, a, x(k + 13), S44, &H4E0811A1
        II a, b, c, d, x(k + 4), S41, &HF7537E82
        II d, a, b, c, x(k + 11), S42, &HBD3AF235
        II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        II b, c, d, a, x(k + 9), S44, &HEB86D391
    
        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next
    
    MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
'*********************************************************




'*********************************************************
' 目的：    Save Text to File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function SaveToFile(strFullName,strContent,strCharset,bolRemoveBOM)

	On Error Resume Next

	Dim objStream

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.WriteText = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With
	Set objStream = Nothing

	If bolRemoveBOM Then

	End If

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Load Text form File
' 输入：    
' 输入：    
' 返回：    
'*********************************************************
Function LoadFromFile(strFullName,strCharset)

	On Error Resume Next

	Dim objStream

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeText
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.LoadFromFile strFullName
	LoadFromFile=.ReadText
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Value For Setting
'*********************************************************
Function SaveValueForSetting(ByRef strContent,bolConst,strTypeVar,strItem,ByVal strValue)

	Dim i,j,s,t
	Dim strConst
	Dim objRegExp

	If bolConst=True Then strConst="Const"

	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True

	If strTypeVar="String" Then

		strValue=Replace(strValue,"""","""""")
		strValue=""""& strValue &""""

		objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))(.+?)(\r\n|\n|$)"
		If objRegExp.Test(strContent)=True Then
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$8")
			SaveValueForSetting=True
			Exit Function
		End If

	End If

	If strTypeVar="Boolean" Then

		strValue=Trim(strValue)
		If LCase(strValue)="true" Then
			strValue=True
		Else
			strValue=False
		End If

		If objRegExp.Test(strContent)=True Then
			objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([a-z]+)( *)(\r\n|\n|$)"
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$9")
			SaveValueForSetting=True
			Exit Function
		End If


	End If

	If strTypeVar="Numeric" Then

		strValue=Trim(strValue)
		If IsNumeric(strValue)=False Then
			strValue=0
		End If

		If objRegExp.Test(strContent)=True Then
			objRegExp.Pattern="(^|\r\n|\n)(( *)" & strConst & "( *)" & strItem & "( *)=( *))([0-9.]+)( *)(\r\n|\n|$)"
			strContent=objRegExp.Replace(strContent,"$1$2"& strValue &"$9")
			SaveValueForSetting=True
			Exit Function
		End If

	End If

	SaveValueForSetting=False

End Function
'*********************************************************




'*********************************************************
' 目的：    检查引用
' 输入：    SQL值（引用）
' 返回：    
'*********************************************************
Function FilterSQL(strSQL)

	FilterSQL=CStr(Replace(strSQL,chr(39),chr(39)&chr(39)))

End Function
'*********************************************************









'/////////////////////////////////////////////////////////////////////////////////////////
Dim objConn
Dim BlogPath
BlogPath=Server.MapPath("wizard.asp")
BlogPath=Left(BlogPath,Len(BlogPath)-Len("wizard.asp"))


Dim BlogHost
Dim DataBasePath
Dim AdminUserName
Dim AdminPassWord
Dim BlogClsid


Dim DataBasePathOld
DataBasePathOld="data/zblog.mdb"

Dim ve,strVerify
ve=Request.QueryString("verify")
strVerify=MD5(DataBasePathOld & Replace(LCase(Request.ServerVariables("PATH_TRANSLATED")),"wizard.asp",""))
If Not ve=strVerify Then
	Response.Write ZC_WD_MSG017
	Response.End
End If


'如果没有data/zblog.mdb则停止输出
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If Not fs.FileExists(BlogPath & DataBasePathOld) Then
	Response.Write ZC_WD_MSG018
	Response.End
End If
Set fs=Nothing


Dim ok
ok=Request.QueryString("ok")
If TypeName(ok)<>"Empty" Then

	BlogHost=Request.Form("edtBlogHost")
	DataBasePath=Request.Form("edtDataBasePath")
	AdminUserName=Request.Form("edtAdminUserName")
	AdminPassWord=Request.Form("edtAdminPassWord")
	BlogClsid=Request.Form("edtBlogClsid")

	AdminPassWord=MD5(AdminPassWord)

	'转移数据库,改数据库名称
	Dim fso, file
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(BlogPath & DataBasePathOld) Then
		Set file = fso.GetFile(BlogPath & DataBasePathOld)
		DataBasePathOld=Mid(DataBasePath,6)
		file.Name=DataBasePathOld
	End If
	Set fso=Nothing


	'建立数据库连接,更改用户名和密码
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BlogPath & DataBasePath
	objConn.Execute("UPDATE [blog_Member] SET [mem_Name]='"&FilterSQL(AdminUserName)&"',[mem_PassWord]='"&FilterSQL(AdminPassWord)&"' WHERE [mem_Name]='zblogger'")
	objConn.Close
	Set objConn=Nothing


	'保存BlogHost,DataBasePath,BlogClsid
	Dim strContent
	strContent=LoadFromFile(BlogPath & "/c_custom.asp","utf-8")
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_HOST",BlogHost)
	Call SaveValueForSetting(strContent,True,"String","ZC_DATABASE_PATH",DataBasePath)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_MASTER",AdminUserName)
	Call SaveToFile(BlogPath & "/c_custom.asp",strContent,"utf-8",False)

	strContent=LoadFromFile(BlogPath & "/c_option.asp","utf-8")
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_CLSID",BlogClsid)
	Call SaveToFile(BlogPath & "/c_option.asp",strContent,"utf-8",False)


	'改写wizard.asp文件
	strContent=LoadFromFile(BlogPath & "/wizard.asp","utf-8")
	strContent=""
	Call SaveToFile(BlogPath & "/wizard.asp",strContent,"utf-8",False)

	'转到首页
	Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"

	Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_WD_MSG001 & "</div>"
	Response.Write "<div id=""divMain2""><form  name=""edit"" id=""edit"">"

	Response.Write "<p>" & ZC_WD_MSG010 &"</p>"
	Response.Write "<p><a href='"& BlogHost &"'>" & ZC_WD_MSG011 & "</a>,"& ZC_WD_MSG019  &"<a href='cmd.asp?act=login'>"& ZC_WD_MSG020 &"</a>.</p>"

	Response.Write "</form></div></div>"
	Response.Write "</body></html>"
	Response.End

End If



BlogHost="http://"  & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("PATH_INFO"),"wizard.asp","")
DataBasePath="data/#%20"& Left(MD5(getGUID()),20) &".mdb"
BlogClsid=getGUID()

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link rel="stylesheet" rev="stylesheet" href="CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="script/common.js" type="text/javascript"></script>
	<title><%=ZC_WD_MSG009%></title>
</head>
<body>
<div id="divMain">
<div class="Header"><%=ZC_WD_MSG001%></div>
<div id="divMain2">
<form id="edit" name="edit" method="post" action="wizard.asp?verify=<%=strVerify%>&ok">
<p><b></b></p>

<p>1.<%=ZC_WD_MSG002%>:</p>
<p><input id="edtBlogHost" name="edtBlogHost" style="width:400px;" type="text" value="<%=BlogHost%>" /></p>
<p><br/>2.<%=ZC_WD_MSG003%>(<%=ZC_WD_MSG008%>):</p>
<p><input readonly id="edtDataBasePath" name="edtDataBasePath" style="width:400px;" type="text" value="<%=DataBasePath%>" /></p>
<p><br/>3.<%=ZC_WD_MSG004%>:</p>
<p><input id="edtAdminUserName" name="edtAdminUserName" style="width:250px;" type="text" value="" /></p>
<p><%=ZC_WD_MSG005%>:</p>
<p><input id="edtAdminPassWord" name="edtAdminPassWord" style="width:250px;" type="password" value="" /></p>
<p><%=ZC_WD_MSG006%>:</p>
<p><input id="edtAdminPassWord2" name="edtAdminPassWord2" style="width:250px;" type="password" value="" /></p>
<div style="display:none;">
<p><br/>4.<%=ZC_WD_MSG007%>(<%=ZC_WD_MSG008%>):</p>
<p><input readonly id="edtBlogClsid" name="edtBlogClsid" style="width:400px;" type="text" value="<%=BlogClsid%>" /></p></div>
<p><br/><input type="submit" class="button" value="<%=ZC_WD_MSG012%>" id="btnPost" onclick='' /></p>

</form>
</div>
</div>
</body>
<script language="JavaScript" type="text/javascript">
	document.getElementById("edit").onsubmit=function(){

		if((!(document.getElementById("edtBlogHost").value).match('^[a-zA-Z]+:\/\/[a-zA-z0-9\-\./:]+?\/$'))){
				alert("<%=ZC_WD_MSG013%>");
				return false;
		}
		if((document.getElementById("edtAdminUserName").value=="")||(!(document.getElementById("edtAdminUserName").value).match('^[.A-Za-z0-9\u4e00-\u9fa5]+$'))){
				alert("<%=ZC_WD_MSG014%>");
				return false;
		}
		if((document.getElementById("edtAdminPassWord").value).length<=5){
				alert("<%=ZC_WD_MSG015%>");
				return false;
		}
		if((document.getElementById("edtAdminPassWord").value!==document.getElementById("edtAdminPassWord2").value)){
				alert("<%=ZC_WD_MSG016%>");
				return false;
		}
	}
</script>
</html>
<script language="javascript" runat="server">
	function getGUID(){
		var guid = "";
		for (var i = 1; i <= 32; i++){
			var n = Math.floor(Math.random() * 16.0).toString(16);
			guid += n;
			if ((i == 8) || (i == 12) || (i == 16) || (i == 20))
			guid += "-";
		}
		guid += "";
		return guid.toUpperCase();
	}
</script>