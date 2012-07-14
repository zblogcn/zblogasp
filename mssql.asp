<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:   
'// 开始时间:   
'// 最后修改:    
'// 备    注:   
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<%Response.Buffer=True %>
<%

Function CheckUpdateDB(a,b)
	Err.Clear
	On Error Resume Next
	Dim Rs
	Set Rs=objConn.execute("SELECT "&a&" FROM "&b)
	Set Rs=Nothing
	If Err.Number=0 Then
	CheckUpdateDB=True
	Else
	Err.Clear
	CheckUpdateDB=False
	End If	
End Function



' Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
' as set out in the memo RFC1321.
'
' See the VB6 project that accompanies this sample for full code comments on how
' it works.
'
' ASP VBScript code for generating an MD5 'digest' or 'signature' of a string. The
' MD5 algorithm is one of the industry standard methods for generating digital
' signatures. It is generically known as a digest, digital signature, one-way
' encryption, hash or checksum algorithm. A common use for MD5 is for password
' encryption as it is one-way in nature, that does not mean that your passwords
' are not free from a dictionary attack.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk

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

Private Function FFFF(x, y, z)
    FFFF = (x And y) Or ((Not x) And z)
End Function

Private Function GGGG(x, y, z)
    GGGG = (x And z) Or (y And (Not z))
End Function

Private Function HHHH(x, y, z)
    HHHH = (x Xor y Xor z)
End Function

Private Function IIII(x, y, z)
    IIII = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(FFFF(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(GGGG(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(HHHH(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(IIII(b, c, d), x), ac))
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
' 目的：   
'*********************************************************
Function RndGuid()

	Dim i,s

	Const c="0123456789ABCDEF"

	Randomize

	For i=1 To 32 

		s=s & Mid(c,Int(Rnd*16)+1,1)
		If i=1 And s="0" Then s="F"

		If i=8 Then s=s & "-"
		If i=12 Then s=s & "-"
		If i=16 Then s=s & "-"
		If i=20 Then s=s & "-"

	Next

	RndGuid=s

End Function 
'*********************************************************

Dim guid
guid=RndGuid

Dim ps
ps=MD5("aa055c6d7875a18fa49058c2c48f2140"&guid)

Function UpdateDateBase()
	
	If Not CheckUpdateDB("[log_IsTop]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_IsTop] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_Article] SET [log_IsTop]=0")
	End If

	If Not CheckUpdateDB("[log_Tag]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Tag] VARCHAR(255) default """"")
	End If

	If Not CheckUpdateDB("[tag_ID]","[blog_Tag]") Then
		objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255) default """",tag_Intro text default """",tag_ParentID int DEFAULT 0,tag_URL VARCHAR(255) default """",tag_Order int DEFAULT 0,tag_Count int DEFAULT 0)")
	End If

	If Not CheckUpdateDB("[coun_ID]","[blog_Counter]") Then
		objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(20) default """",coun_Agent text default """",coun_Refer VARCHAR(255) default """",coun_PostTime TIME DEFAULT Now())")
	End If

	If Not CheckUpdateDB("[key_ID]","[blog_Keyword]") Then
		objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255) default """",key_Intro text default """",key_URL VARCHAR(255) default """")")
	End If

	If Not CheckUpdateDB("[ul_Quote]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Quote] VARCHAR(255) default """"")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DownNum] int DEFAULT 0")
	End If

	If Not CheckUpdateDB("[ul_FileIntro]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_FileIntro] VARCHAR(255) default """"")
	End If

	If Not CheckUpdateDB("[ul_DirByTime]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_DirByTime] YESNO DEFAULT FALSE")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_DirByTime]=[ul_Quote]")
		objConn.execute("UPDATE [blog_UpLoad] SET [ul_Quote]=''")
	End If

	If Not CheckUpdateDB("[log_Meta]","[blog_Article]") Then
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Ratting] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Template] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_IsAnonymous] YESNO DEFAULT FALSE")
		objConn.execute("ALTER TABLE [blog_Article] ADD COLUMN [log_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[cate_Meta]","[blog_Category]") Then
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Temp] VARCHAR(255) default """"")
		objConn.execute("UPDATE [blog_Category] SET [cate_Temp]=[cate_Url]")
		objConn.execute("UPDATE [blog_Category] SET [cate_URL]=[cate_Intro]")
		objConn.execute("UPDATE [blog_Category] SET [cate_Intro]=[cate_Temp]")
		objConn.execute("ALTER TABLE [blog_Category] DROP COLUMN [cate_Temp]")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_ParentID] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Template] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Category] ADD COLUMN [cate_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[comm_Meta]","[blog_Comment]") Then
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Reply] text default """"")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyIP] VARCHAR(15) default """"")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_LastReplyTime] datetime default now()")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Yea] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Nay] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Ratting] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_ParentID] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_ParentCount] int DEFAULT 0")
		objConn.execute("ALTER TABLE [blog_Comment] ADD COLUMN [comm_Meta] text default """"")

		objConn.execute("UPDATE [blog_Comment] SET [comm_ParentID]=0")
		objConn.execute("UPDATE [blog_Comment] SET [comm_ParentCount]=0")
	End If
	
	If Not CheckUpdateDB("[mem_Meta]","[blog_Member]") Then
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Guid] VARCHAR(36) default """"")
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[ul_Meta]","[blog_UpLoad]") Then
		objConn.execute("ALTER TABLE [blog_UpLoad] ADD COLUMN [ul_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[tb_Meta]","[blog_TrackBack]") Then
		objConn.execute("ALTER TABLE [blog_TrackBack] ADD COLUMN [tb_Meta] text default """"")
	End If

	If Not CheckUpdateDB("[tag_Meta]","[blog_Tag]") Then
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Template] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_FullUrl] VARCHAR(255) default """"")
		objConn.execute("ALTER TABLE [blog_Tag] ADD COLUMN [tag_Meta] text default """"")
	End If

End Function




Dim objConn	
Dim objCat


'这段是升级数据库的



'Set objConn = Server.CreateObject("ADODB.Connection")
'objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("data/#%20d7d52c946f1f84b79884.mdb")

'Call UpdateDateBase()













'这段是创建一个全新的空的ACCESS数据库及默认数据


Set objCat=Server.CreateObject("ADOX.Catalog")   
objCat.Create  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("zblog.mdb")


Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("zblog.mdb")

objConn.execute("CREATE TABLE [blog_Tag] (tag_ID AutoIncrement primary key,tag_Name VARCHAR(255) default """",tag_Intro text default """",tag_ParentID int default 0,tag_URL VARCHAR(255) default """",tag_Order int default 0,tag_Count int default 0,tag_Template VARCHAR(255) default """",tag_FullUrl VARCHAR(255) default """",tag_Meta text default """")")

objConn.execute("CREATE TABLE [blog_Article] (log_ID AutoIncrement primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url VARCHAR(255) default """",log_Title VARCHAR(255) default """",log_Intro text default """",log_Content text default """",log_IP VARCHAR(15) default """",log_PostTime datetime default now(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag VARCHAR(255) default """",log_IsTop YESNO DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Ratting int default 0,log_Template VARCHAR(255) default """",log_FullUrl VARCHAR(255) default """",log_IsAnonymous YESNO DEFAULT FALSE,log_Meta text default """")")

objConn.execute("CREATE TABLE [blog_Category] (cate_ID AutoIncrement primary key,cate_Name VARCHAR(50) default """",cate_Order int default 0,cate_Intro VARCHAR(255) default """",cate_Count int default 0,cate_URL VARCHAR(255) default """",cate_ParentID int default 0,cate_Template VARCHAR(255) default """",cata_FullUrl VARCHAR(255) default """",cate_Meta text default """")")

objConn.execute("CREATE TABLE [blog_Comment] (comm_ID AutoIncrement primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author VARCHAR(20) default """",comm_Content text default """",comm_Email VARCHAR(50) default """",comm_HomePage VARCHAR(255) default """",comm_PostTime datetime default now(),comm_IP VARCHAR(15) default """",comm_Agent text default """",comm_Reply text default """",comm_LastReplyIP VARCHAR(15) default """",comm_LastReplyTime datetime default now(),comm_Yea int default 0,comm_Nay int default 0,comm_Ratting int default 0,comm_ParentID int default 0,comm_ParentCount int default 0,comm_Meta text default """")")

objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID AutoIncrement primary key,log_ID int default 0,tb_URL VARCHAR(255) default """",tb_Title VARCHAR(100) default """",tb_Blog VARCHAR(50) default """",tb_Excerpt text default """",tb_PostTime datetime default now(),tb_IP VARCHAR(15) default """",tb_Agent text default """",tb_Meta text default """")")

objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID AutoIncrement primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName VARCHAR(50) default """",ul_PostTime datetime default now(),ul_Quote VARCHAR(255) default """",ul_DownNum int default 0,ul_FileIntro VARCHAR(255) default """",ul_DirByTime YESNO DEFAULT 0,ul_Meta text default """")")

objConn.execute("CREATE TABLE [blog_Counter] (coun_ID AutoIncrement primary key,coun_IP VARCHAR(15) default """",coun_Agent text default """",coun_Refer VARCHAR(255) default """",coun_PostTime datetime default now() )")

objConn.execute("CREATE TABLE [blog_Keyword] (key_ID AutoIncrement primary key,key_Name VARCHAR(255) default """",key_Intro text default """",key_URL VARCHAR(255) default """")")

objConn.execute("CREATE TABLE [blog_Member] (mem_ID AutoIncrement primary key,mem_Level int default 0,mem_Name VARCHAR(20) default """",mem_Password VARCHAR(32) default """",mem_Sex int default 0,mem_Email VARCHAR(50) default """",mem_MSN VARCHAR(50) default """",mem_QQ VARCHAR(50) default """",mem_HomePage VARCHAR(255) default """",mem_LastVisit datetime default now(),mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro text default """",mem_IP VARCHAR(15) default """",mem_Count int default 0,mem_Guid VARCHAR(36) default """",mem_Meta text default """")")

objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro],[mem_Guid]) VALUES (1,'zblogger','"&ps&"','null@null.com','','','"&guid&"')")





'这段是在指定的MSSQL数据库里创建新表及默认数据





Set objConn = Server.CreateObject("ADODB.Connection")

objConn.Open "Provider=SqlOLEDB;Data Source=(local)\SQLEXPRESS;Initial Catalog=zb;Persist Security Info=True;User ID=sa;Password=;"

objConn.execute("CREATE TABLE [blog_Tag] (tag_ID int identity(1,1) not null primary key,tag_Name nvarchar(255) default '',tag_Intro ntext default '',tag_ParentID int default 0,tag_URL nvarchar(255) default '',tag_Order int default 0,tag_Count int default 0,tag_Template nvarchar(255) default '',tag_FullUrl nvarchar(255) default '',tag_Meta ntext default '')")

objConn.execute("CREATE TABLE [blog_Article] (log_ID int identity(1,1) not null primary key,log_CateID int default 0,log_AuthorID int default 0,log_Level int default 0,log_Url nvarchar(255) default '',log_Title nvarchar(255) default '',log_Intro ntext default '',log_Content ntext default '',log_IP nvarchar(15) default '',log_PostTime datetime default getdate(),log_CommNums int default 0,log_ViewNums int default 0,log_TrackBackNums int default 0,log_Tag nvarchar(255) default '',log_IsTop bit DEFAULT 0,log_Yea int default 0,log_Nay int default 0,log_Ratting int default 0,log_Template nvarchar(255) default '',log_FullUrl nvarchar(255) default '',log_IsAnonymous bit DEFAULT 0,log_Meta ntext default '')")

objConn.execute("CREATE TABLE [blog_Category] (cate_ID int identity(1,1) not null primary key,cate_Name nvarchar(50) default '',cate_Order int default 0,cate_Intro nvarchar(255) default '',cate_Count int default 0,cate_URL nvarchar(255) default '',cate_ParentID int default 0,cate_Template nvarchar(255) default '',cate_FullUrl nvarchar(255) default '',cate_Meta ntext default '')")

objConn.execute("CREATE TABLE [blog_Comment] (comm_ID int identity(1,1) not null primary key,log_ID int default 0,comm_AuthorID int default 0,comm_Author nvarchar(20) default '',comm_Content ntext default '',comm_Email nvarchar(50) default '',comm_HomePage nvarchar(255) default '',comm_PostTime datetime default getdate(),comm_IP nvarchar(15) default '',comm_Agent ntext default '',comm_Reply ntext default '',comm_LastReplyIP nvarchar(15) default '',comm_LastReplyTime datetime default getdate(),comm_Yea int default 0,comm_Nay int default 0,comm_Ratting int default 0,comm_ParentID int default 0,comm_ParentCount int default 0,comm_Meta ntext default '')")

objConn.execute("CREATE TABLE [blog_TrackBack] (tb_ID int identity(1,1) not null primary key,log_ID int default 0,tb_URL nvarchar(255) default '',tb_Title nvarchar(100) default '',tb_Blog nvarchar(50) default '',tb_Excerpt ntext default '',tb_PostTime datetime default getdate(),tb_IP nvarchar(15) default '',tb_Agent ntext default '',tb_Meta ntext default '')")

objConn.execute("CREATE TABLE [blog_UpLoad] (ul_ID int identity(1,1) not null primary key,ul_AuthorID int default 0,ul_FileSize int default 0,ul_FileName nvarchar(50) default '',ul_PostTime datetime default getdate(),ul_Quote nvarchar(255) default '',ul_DownNum int default 0,ul_FileIntro nvarchar(255) default '',ul_DirByTime bit DEFAULT 0,ul_Meta ntext default '')")

objConn.execute("CREATE TABLE [blog_Counter] (coun_ID int identity(1,1) not null primary key,coun_IP nvarchar(15) default '',coun_Agent ntext default '',coun_Refer nvarchar(255) default '',coun_PostTime datetime default getdate() )")

objConn.execute("CREATE TABLE [blog_Keyword] (key_ID int identity(1,1) not null primary key,key_Name nvarchar(255) default '',key_Intro ntext default '',key_URL nvarchar(255) default '')")

objConn.execute("CREATE TABLE [blog_Member] (mem_ID int identity(1,1) not null primary key,mem_Level int default 0,mem_Name nvarchar(20) default '',mem_Password nvarchar(32) default '',mem_Sex int default 0,mem_Email nvarchar(50) default '',mem_MSN nvarchar(50) default '',mem_QQ nvarchar(50) default '',mem_HomePage nvarchar(255) default '',mem_LastVisit datetime default getdate(),mem_Status int default 0,mem_PostLogs int default 0,mem_PostComms int default 0,mem_Intro ntext default '',mem_IP nvarchar(15) default '',mem_Count int default 0,mem_Guid  nvarchar(36) default '',mem_Meta ntext default '')")

objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro],[mem_Guid]) VALUES (1,'zblogger','"&ps&"','null@null.com','','','"&guid&"')")


objConn.Close




%>