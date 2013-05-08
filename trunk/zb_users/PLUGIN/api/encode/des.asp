<%
Class Cls_DES
    Private IPRule, CPRule, EPRule, PRule, SBox(7), PCRule(1), MvRule
    Private K(16), L(16), R(16)
    Private FillCode, DesStatus
    
    Private Sub Class_Initialize()
        DesStatus = -1
        FillCode = "0001101"
        IPRule = "58,50,42,34,26,18,10,2," &_
            "60,52,44,36,28,20,12,4," &_
            "62,54,46,38,30,22,14,6," &_
            "64,56,48,40,32,24,16,8," &_
            "57,49,41,33,25,17, 9,1," &_
            "59,51,43,35,27,19,11,3," &_
            "61,53,45,37,29,21,13,5," &_
            "63,55,47,39,31,23,15,7,"
        CPRule = "40, 8,48,16,56,24,64,32," &_
            "39, 7,47,15,55,23,63,31," &_
            "38, 6,46,14,54,22,62,30," &_
            "37, 5,45,13,53,21,61,29," &_
            "36, 4,44,12,52,20,60,28," &_
            "35, 3,43,11,51,19,59,27," &_
            "34, 2,42,10,50,18,58,26," &_
            "33, 1,41, 9,49,17,57,25,"
        EPRule = "32, 1, 2, 3, 4, 5," &_
            " 4, 5, 6, 7, 8, 9," &_
            " 8, 9,10,11,12,13," &_
            "12,13,14,15,16,17," &_
            "16,17,18,19,20,21," &_
            "20,21,22,23,24,25," &_
            "24,25,26,27,28,29," &_
            "28,29,30,31,32, 1,"
        PRule = "16, 7,20,21,29,12,28,17," &_
            " 1,15,23,26, 5,18,31,10," &_
            " 2, 8,24,14,32,27, 3, 9," &_
            "19,13,30, 6,22,11, 4,25,"
        SBox(0) = "14, 4,13, 1, 2,15,11, 8, 3,10, 6,12, 5, 9, 0, 7," &_
            " 0,15, 7, 4,14, 2,13, 1,10, 6,12,11, 9, 5, 3, 8," &_
            " 4, 1,14, 8,13, 6, 2,11,15,12, 9, 7, 3,10, 5, 0," &_
            "15,12, 8, 2, 4, 9, 1, 7, 5,11, 3,14,10, 0, 6,13,"
        SBox(1) = "15, 1, 8,14, 6,11, 3, 4, 9, 7, 2,13,12, 0, 5,10," &_
            " 3,13, 4, 7,15, 2, 8,14,12, 0, 1,10, 6, 9,11, 5," &_
            " 0,14, 7,11,10, 4,13, 1, 5, 8,12, 6, 9, 3, 2,15," &_
            "13, 8,10, 1, 3,15, 4, 2,11, 6, 7,12, 0, 5,14, 9,"
        SBox(2) = "10, 0, 9,14, 6, 3,15, 5, 1,13,12, 7,11, 4, 2, 8," &_
            "13, 7, 0, 9, 3, 4, 6,10, 2, 8, 5,14,12,11,15, 1," &_
            "13, 6, 4, 9, 8,15, 3, 0,11, 1, 2,12, 5,10,14, 7," &_
            " 1,10,13, 0, 6, 9, 8, 7, 4,15,14, 3,11, 5, 2,12,"
        SBox(3) = " 7,13,14, 3, 0, 6, 9,10, 1, 2, 8, 5,11,12, 4,15," &_
            "13, 8,11, 5, 6,15, 0, 3, 4, 7, 2,12, 1,10,14, 9," &_
            "10, 6, 9, 0,12,11, 7,13,15, 1, 3,14, 5, 2, 8, 4," &_
            " 3,15, 0, 6,10, 1,13, 8, 9, 4, 5,11,12, 7, 2,14,"
        SBox(4) = " 2,12, 4, 1, 7,10,11, 6, 8, 5, 3,15,13, 0,14, 9," &_
            "14,11, 2,12, 4, 7,13, 1, 5, 0,15,10, 3, 9, 8, 6," &_
            " 4, 2, 1,11,10,13, 7, 8,15, 9,12, 5, 6, 3, 0,14," &_
            "11, 8,12, 7, 1,14, 2,13, 6,15, 0, 9,10, 4, 5, 3,"
        SBox(5) = "12, 1,10,15, 9, 2, 6, 8, 0,13, 3, 4,14, 7, 5,11," &_
            "10,15, 4, 2, 7,12, 9, 5, 6, 1,13,14, 0,11, 3, 8," &_
            " 9,14,15, 5, 2, 8,12, 3, 7, 0, 4,10, 1,13,11, 6," &_
            " 4, 3, 2,12, 9, 5,15,10,11,14, 1, 7, 6, 0, 8,13,"
        SBox(6) = " 4,11, 2,14,15, 0, 8,13, 3,12, 9, 7, 5,10, 6, 1," &_
            "13, 0,11, 7, 4, 9, 1,10,14, 3, 5,12, 2,15, 8, 6," &_
            " 1, 4,11,13,12, 3, 7,14,10,15, 6, 8, 0, 5, 9, 2," &_
            " 6,11,13, 8, 1, 4,10, 7, 9, 5, 0,15,14, 2, 3,12,"
        SBox(7) = "13, 2, 8, 4, 6,15,11, 1,10, 9, 3,14, 5, 0,12, 7," &_
            " 1,15,13, 8,10, 3, 7, 4,12, 5, 6,11, 0,14, 9, 2," &_
            " 7,11, 4, 1, 9,12,14, 2, 0, 6,10,13,15, 3, 5, 8," &_
            " 2, 1,14, 7, 4,10, 8,13,15,12, 9, 0, 3, 5, 6,11,"
        PCRule(0) = "57,49,41,33,25,17, 9," &_
            " 1,58,50,42,34,26,18," &_
            "10, 2,59,51,43,35,27," &_
            "19,11, 3,60,52,44,36," &_
            "63,55,47,39,31,23,15," &_
            " 7,62,54,46,38,30,22," &_
            "14, 6,61,53,45,37,29," &_
            "21,13, 5,28,20,12, 4,"
        PCRule(1) = "14,17,11,24, 1, 5, 3,28," &_
            "15, 6,21,10,23,19,12, 4," &_
            "26, 8,16, 7,27,20,13, 2," &_
            "41,52,31,37,47,55,30,40," &_
            "51,45,33,48,44,49,39,56," &_
            "34,53,46,42,50,36,29,32,"
        MvRule = "1,1,2,2,2,2,2,2,1,2,2,2,2,2,2,1"
    End Sub
    
    Private Function Permute(ByVal Rule, ByVal Text)
        Dim P_Rule, Num, PText
        PText = ""
        P_Rule = Split(Rule, ",")
        For Each Num In P_Rule
            Num = Trim(Num) & ""
            If Num <> "" Then
                Num = CLng(Num)
                PText = PText & Mid(Text, Num, 1)
            End If
        Next
        Erase P_Rule
        Permute = PText
    End Function
    
    Private Function CreateKey()
        Dim IPKey, C(16), D(16), i, Mv_Rule, MvLen
        IPKey = Permute(PCRule(0), K(0))
        C(0) = Left(IPKey, 28)
        D(0) = Right(IPKey, 28)
        Mv_Rule = Split(MvRule, ",")
        For i = 1 To 16
            MvLen = CLng(Trim(Mv_Rule(i - 1)))
            C(i) = Right(C(i -1), Len(C(i -1)) - MvLen) & Left(C(i -1), MvLen)
            D(i) = Right(D(i -1), Len(D(i -1)) - MvLen) & Left(D(i -1), MvLen)
            K(i) = Permute(PCRule(1), C(i) & D(i))
        Next
    End Function

    Private Function IP(ByVal Text)
        Dim IPText
        IPText = Permute(IPRule, Text)
        L(0) = Left(IPText, 32)
        R(0) = Right(IPText, 32)
        IP = IPText
    End Function
    
    Private Function IterativeLR()
        Dim i
        For i = 1 To 16
            L(i) = R(i - 1)
            R(i) = B_XOR(L(i - 1), F(R(i - 1), K(i)))
        Next
    End Function
    
    Private Function F(ByVal RText, ByVal Keys)
        Dim EPText, XORText, Result, SKey(7), i, x, y
        Result = ""
        EPText = Permute(EPRule, RText)
        XORText = B_XOR(EPText, Keys)
        For i = 1 To Len(XORText) \ 6
            SKey(i - 1) = Mid(XORText, (i - 1) * 6 + 1, 6)
            x = BinaryToDecimal(Left(SKey(i - 1), 1) & Right(SKey(i - 1), 1))
            y = BinaryToDecimal(Mid(SKey(i - 1), 2, 4))
            SKey(i - 1) = DecimalToBinary(Trim(Split(SBox(i -1), ",")(x * 16 + y)))
            If Len(SKey(i - 1)) < 4 Then
                Select Case (4 - Len(SKey(i - 1)))
                    Case 1
                        SKey(i - 1) = "0" & SKey(i - 1)
                    Case 2
                        SKey(i - 1) = "00" & SKey(i - 1)
                    Case 3
                        SKey(i - 1) = "000" & SKey(i - 1)
                End Select
            End If
            Result = Result & SKey(i - 1)
        Next
        Result = Permute(PRule, Result)
        F = Result
    End Function
    
    Private Function B_XOR(ByVal Expression1, ByVal Expression2)
        Dim E, K, i, XORText
        XORText = ""
        E = Trim(Expression1) & ""
        K = Trim(Expression2) & ""
        For i = 1 To Len(K)
            XORText = XORText & CStr(CInt(Mid(E, i, 1)) Xor CInt(Mid(K, i, 1)))
        Next
        B_XOR = XORText
    End Function
    
    Private Function BinaryToDecimal(ByVal binNum)
        Dim Binary, Decimal, i, Length
        Decimal = 0
        Binary = Trim(binNum) & ""
        If Binary <> "" Then
            While Left(Binary, 1) = "0"
                Binary = Right(Binary, Len(Binary) - 1)
            Wend
            Length = Len(Binary)
            For i = 1 To Length
                Decimal = Decimal + CInt(Mid(Binary, i, 1)) * 2^(Length - i)
            Next
        End If
        BinaryToDecimal = Decimal
    End Function
    
    Private Function DecimalToBinary(ByVal decNum)
        Dim Decimal, Binary, division
        Binary = ""
        Decimal = Trim(decNum) & ""
        If Decimal <> "" Then
            Decimal = CLng(Decimal)
            While Decimal > 1
                Binary = Binary & CStr(Decimal Mod 2)
                Decimal = Decimal \ 2
            Wend
            Binary = StrReverse(Binary & Decimal)
        End If
        DecimalToBinary = Binary
    End Function
    
    Private Function StrToBinary(ByVal Str)
        Dim Data, Binary, Text, TextLen, i
        Text = ""
        Data = Str
        For i = 1 To Len(Data)
            Binary = CStr(DecimalToBinary(Asc(Mid(Data, i, 1))))
            If Len(Binary) < 7 Then
                Select Case (7 - Len(Binary))
                    Case 1
                        Binary = "0" & Binary
                    Case 2
                        Binary = "00" & Binary
                    Case 3
                        Binary = "000" & Binary
                    Case 4
                        Binary = "0000" & Binary
                    Case 5
                        Binary = "00000" & Binary
                    Case 6
                        Binary = "000000" & Binary
                End Select
            End If
            Text = Text & Binary
        Next
        TextLen = Len(Text)
        If TextLen >= 63 Then
            If (TextLen Mod 63) <> 0 Then
                For i = 1 To ((TextLen - TextLen Mod 63) \ 7)
                    Text = Text & FillCode
                Next
            End If
        Else
            For i = 1 To ((63 - TextLen) \ 7)
                Text = Text & FillCode
            Next
        End If

        Binary = Text
        Text = ""
        For i = 0 To (Len(Binary) \ 63 - 1)
            Text = Text & Mid(Binary, i * 63 + 1, 63) & "0"
        Next
        StrToBinary = Text
    End Function
    
    Private Function BinaryToStr(ByVal binNum)
        Dim Text, binText, Length, Group, i, j
        Text = ""
        binText = Trim(binNum) & ""
        If binText <> "" Then
            Length = Len(binText) \ 64 - 1
            ReDim Group(Length)
            For i = 0 To Length
                Group(i) = Left(Mid(binText, i * 64 + 1, 64), 63)
            Next
            While Right(Group(Length), 7) = FillCode
                Group(Length) = Left(Group(Length), Len(Group(Length)) - 7)
            Wend
            For i = 0 To Length
                For j = 1 To Len(Group(i)) \ 7
                    Text = Text & Chr(BinaryToDecimal(Mid(Group(i), (j - 1) * 7 + 1, 7)))
                Next
            Next
            Erase Group
        End If
        BinaryToStr = Text
    End Function
    
    Private Function BinaryToHex(ByVal binNum)
        Dim binText, Text, Length, FillLen, Temp, i
        Text = ""
        binText = Trim(binNum) & ""
        If binText <> "" Then
            Length = Len(binText)
            If Length >= 4 Then
                FillLen = Length Mod 4
            Else
                FillLen = 4 - Length
            End If
            Select Case FillLen
                Case 1
                    binText = "0" & binText
                Case 2
                    binText = "00" & binText
                Case 3
                    binText = "000" & binText
            End Select
            For i = 0 To (Len(binText) \ 4 - 1)
                Temp = Mid(binText, i * 4 + 1, 4)
                Select Case Temp
                    Case "0000"
                        Text = Text & "0"
                    Case "0001"
                        Text = Text & "1"
                    Case "0010"
                        Text = Text & "2"
                    Case "0011"
                        Text = Text & "3"
                    Case "0100"
                        Text = Text & "4"
                    Case "0101"
                        Text = Text & "5"
                    Case "0110"
                        Text = Text & "6"
                    Case "0111"
                        Text = Text & "7"
                    Case "1000"
                        Text = Text & "8"
                    Case "1001"
                        Text = Text & "9"
                    Case "1010"
                        Text = Text & "A"
                    Case "1011"
                        Text = Text & "B"
                    Case "1100"
                        Text = Text & "C"
                    Case "1101"
                        Text = Text & "D"
                    Case "1110"
                        Text = Text & "E"
                    Case "1111"
                        Text = Text & "F"
                End Select
            Next
        End If
        BinaryToHex = Text
    End Function
    
    Private Function HexToBinary(ByVal hexNum)
        Dim hexText, Text, Temp, i
        Text = ""
        hexText = Trim(hexNum) & ""
        For i = 1 To Len(hexText)
            Temp = UCase(Mid(hexText, i, 1))
            Select Case Temp
                Case "0"
                    Text = Text & "0000"
                Case "1"
                    Text = Text & "0001"
                Case "2"
                    Text = Text & "0010"
                Case "3"
                    Text = Text & "0011"
                Case "4"
                    Text = Text & "0100"
                Case "5"
                    Text = Text & "0101"
                Case "6"
                    Text = Text & "0110"
                Case "7"
                    Text = Text & "0111"
                Case "8"
                    Text = Text & "1000"
                Case "9"
                    Text = Text & "1001"
                Case "A"
                    Text = Text & "1010"
                Case "B"
                    Text = Text & "1011"
                Case "C"
                    Text = Text & "1100"
                Case "D"
                    Text = Text & "1101"
                Case "E"
                    Text = Text & "1110"
                Case "F"
                    Text = Text & "1111"
            End Select
        Next
        HexToBinary = Text
    End Function
    
    Private Function KeyReverse()
        Dim Temp, i
        For i = 1 To 8
            Temp = K(i)
            K(i) = K(16 - i + 1)
            K(16 - i + 1) = Temp
        Next
    End Function
    
    Public Function DES(ByVal Data, ByVal Keys, ByVal Work)
        Dim Text, i, Group, GroupLen
        Text = Data
        K(0) = HexToBinary(Keys)
        If Work = 0 Then
            Text = StrToBinary(Text)
        Else
            Text = HexToBinary(Text)
        End If
        GroupLen = Len(Text) \ 64 - 1
        ReDim Group(GroupLen)
        For i = 0 To GroupLen
            Group(i) = Mid(Text, i * 64 + 1, 64)
        Next
        Text = ""
        CreateKey()
        For i = 0 To GroupLen
            IP(Group(i))
            If Work <> 0 And DesStatus <> 1 Then
                KeyReverse()
                DesStatus = 1
            ElseIf Work = 0 And DesStatus = 1 Then
                KeyReverse()
                DesStatus = 0
            End If
            IterativeLR()
            Text = Text & Permute(CPRule, R(16) & L(16))
        Next
        Erase Group
        If Work = 0 Then
            Text = BinaryToHex(Text)
        Else
            Text = BinaryToStr(Text)
        End If

        DES = Text
    End Function
End Class
function Des(ByVal Data, ByVal Keys, ByVal Work)
    Set DesCrypt = New Cls_DES
    Des = DesCrypt.DES(Data, Keys, Work)
    Set DesCrypt = Nothing
end function
%>
<%
'Response.Write("cftea(key:F7F741E99D0137) -&gt; " & Des("cftea", "F7F741E99D0137", 0))
%>