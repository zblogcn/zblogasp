Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TextBox2.Text = My.Computer.FileSystem.GetFiles(TextBox1.Text, FileIO.SearchOption.SearchAllSubDirectories).Count




        For Each c As String In My.Computer.FileSystem.GetDirectories(My.Application.Info.DirectoryPath, FileIO.SearchOption.SearchTopLevelOnly)

            If My.Computer.FileSystem.GetDirectoryInfo(c).Name.Length <> 6 Then Exit For

            Dim b2 As String = ""
            Dim d As String = ""
            Dim d_r As String = ""
            Dim d_n As String = ""
            Dim f As String = ""

            b2 += ("<files build='" + My.Computer.FileSystem.GetDirectoryInfo(c).Name + "'>" + vbCrLf)

            Dim l As New System.Collections.Generic.Dictionary(Of String, String)
            Dim m As New System.Collections.Generic.Dictionary(Of String, String)

            For Each a As String In My.Computer.FileSystem.GetFiles(c, FileIO.SearchOption.SearchAllSubDirectories)

                Dim b As String = ""

                b += vbTab + "<file"

                b += " name='" + a.Replace(c + "\", "") + "'"

                Dim bytMD5 As Byte() = New System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(My.Computer.FileSystem.ReadAllBytes(a))
                Dim tems As String = My.Computer.FileSystem.ReadAllText(a, System.Text.ASCIIEncoding.UTF8)
                tems = Replace(tems, vbCrLf, vbCr)
                tems = Replace(tems, vbLf, vbCr)

                Dim bytMD5_r As Byte() = New System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(System.Text.ASCIIEncoding.UTF8.GetBytes(tems))
                d = BitConverter.ToString(bytMD5).Replace("-", String.Empty).ToUpper
                d_r = BitConverter.ToString(bytMD5_r).Replace("-", String.Empty).ToUpper
                tems = Replace(tems, vbCr, vbLf)
                Dim bytMD5_n As Byte() = New System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(System.Text.ASCIIEncoding.UTF8.GetBytes(tems))
                d_n = BitConverter.ToString(bytMD5_n).Replace("-", String.Empty).ToUpper
                b += " md5='" + d + "'"
                b += " md5_r='" + d_r + "'"
                b += " md5_n='" + d_n + "'"

                Dim g As New Crc32
                If (My.Computer.FileSystem.ReadAllBytes(a).Length = 0) Then
                    f = ""
                Else
                    f = Convert.ToString(g.CalculateBlock(My.Computer.FileSystem.ReadAllBytes(a)), 16).ToUpper
                End If

                b += " crc32='" + f + "'"
                b += "/>" + vbCrLf

                l.Add(a.Replace(c + "\", ""), b)
                'b2 += b
            Next

            ' Iterate through a dictionary
            For Each x As String In l.Keys
                If x = "zb_system\function\c_system_plugin.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\c_system_debug.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\c_system_common.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\c_system_event.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\c_system_base.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\c_system_admin.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\defend\option.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\lib\base.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\lib\dbsql.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\lib\metas.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\lib\template.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\lib\network.php" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\function\lib\zblogphp.php" Then m.Add(x, l(x))
            Next


            For Each x As String In l.Keys
                If x = "zb_system\FUNCTION\c_function.asp" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\FUNCTION\c_system_plugin.asp" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\FUNCTION\c_system_lib.asp" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\FUNCTION\c_system_base.asp" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\FUNCTION\c_system_manage.asp" Then m.Add(x, l(x))
            Next
            For Each x As String In l.Keys
                If x = "zb_system\FUNCTION\c_system_event.asp" Then m.Add(x, l(x))
            Next




            For Each x As String In l.Keys
                If m.ContainsKey(x) Then Continue For
                m.Add(x, l(x))
            Next

            For Each x As String In m.Values
                b2 += x
            Next



            b2 += "</files>"
            My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.GetDirectoryInfo(c).Name + ".xml", b2, False)
            Me.TextBox2.AppendText("生成文件:" + My.Computer.FileSystem.GetDirectoryInfo(c).Name + ".xml" + vbCrLf)
        Next

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click


        Dim a As Integer = Me.ComboBox1.SelectedIndex + 1
        Dim b As String = Me.ComboBox1.Items(a - 1)
        Dim f(a) As Collections.Generic.Dictionary(Of String, String)
        Dim f2(a) As Collections.Generic.Dictionary(Of String, String)

        For i As Integer = 1 To a
            f2(i) = New Collections.Generic.Dictionary(Of String, String)
        Next

        Dim x As New Xml.XmlDocument
        For i As Integer = 1 To a
            f(i) = New Collections.Generic.Dictionary(Of String, String)
            x.Load(My.Application.Info.DirectoryPath + "\" + Me.ComboBox1.Items(i - 1) + ".xml")
            For Each y As Xml.XmlNode In x.DocumentElement.SelectNodes("file")
                'If y.Attributes("name").InnerText.Contains("zb_users\") = False Then

                f(i).Add(y.Attributes("name").InnerText, y.Attributes("crc32").InnerText)
                'Else
                '    If y.Attributes("name").InnerText.Contains("zb_users\LANGUAGE\SimpChinese.asp") = True Then
                '        f(i).Add(y.Attributes("name").InnerText, y.Attributes("crc32").InnerText)
                '    End If

                '    If y.Attributes("name").InnerText.Contains("zb_users\EMOTION\") = True Then
                '        f(i).Add(y.Attributes("name").InnerText, y.Attributes("crc32").InnerText)
                '    End If

                'End If

            Next
            'MsgBox(f(i).Count)
        Next

        For i As Integer = a To 2 Step -1
            If i = a Then
                For Each s As String In f(i).Keys

                    If f(i - 1).ContainsKey(s) = False Then
                        Me.TextBox2.AppendText(s + vbCrLf)
                        f2(i).Add(s, f(i)(s))
                    Else
                        If String.Compare(f(i - 1)(s), f(i)(s)) <> 0 Then
                            Me.TextBox2.AppendText(s + vbCrLf)
                            f2(i).Add(s, f(i)(s))
                        End If

                    End If

                Next
            Else
                'For Each s As String In f2(i + 1).Keys
                '    If f(i - 1).ContainsKey(s) = False Then
                '        Me.TextBox2.AppendText(s + vbCrLf)
                '        f2(i).Add(s, f2(i + 1)(s))
                '    Else
                '        If String.Compare(f(i - 1)(s), f2(i + 1)(s)) <> 0 Then
                '            Me.TextBox2.AppendText(s + vbCrLf)
                '            f2(i).Add(s, f2(i + 1)(s))
                '        End If
                '    End If
                'Next
                For Each s As String In f(i).Keys
                    If f(i - 1).ContainsKey(s) = False Then
                        Me.TextBox2.AppendText(s + vbCrLf)
                        f2(i).Add(s, f(i)(s))
                    Else
                        If String.Compare(f(i - 1)(s), f(i)(s)) <> 0 Then
                            Me.TextBox2.AppendText(s + vbCrLf)
                            f2(i).Add(s, f(i)(s))
                        End If
                    End If
                Next
                For Each s As String In f2(i + 1).Keys
                    If f2(i).ContainsKey(s) = False Then
                        f2(i).Add(s, f2(i + 1)(s))
                    End If
                Next

            End If
            'MsgBox(f2(i).Count)
        Next

        For i As Integer = 0 To a - 2

            Dim s As String = Nothing
            Dim d As String = Nothing

            s = My.Application.Info.DirectoryPath + "\" + Me.ComboBox1.Items(i) + "-" + Me.ComboBox1.Items(a - 1) + ".xml"
            Dim t As String = Nothing



            t = "<files codepage='65001' xmlns:dt='urn:schemas-microsoft-com:datatypes'>" + vbCrLf
            For Each j As String In f2(i + 2).Keys
                If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath + "\" + b + "\" + j) = True Then

                    d = System.Convert.ToBase64String(My.Computer.FileSystem.ReadAllBytes(My.Application.Info.DirectoryPath + "\" + b + "\" + j))

                    Dim g As New Crc32
                    Dim m As String = Nothing
                    If My.Computer.FileSystem.ReadAllBytes(My.Application.Info.DirectoryPath + "\" + b + "\" + j).Length = 0 Then
                        m = Nothing
                    Else
                        m = Convert.ToString(g.CalculateBlock(My.Computer.FileSystem.ReadAllBytes(My.Application.Info.DirectoryPath + "\" + b + "\" + j)), 16).ToUpper()

                    End If


                    t += vbTab + "<file name='" + j + "' crc32='" + m + "'  dt:dt='bin.base64'>" + d + "</file>" + vbCrLf
                End If
            Next
            t += "</files>"

            My.Computer.FileSystem.WriteAllText(s, t, False)


        Next


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim b As String = ""
        Dim i As Integer = 0
        b += "<builds>" + vbCrLf
        For Each c As String In My.Computer.FileSystem.GetDirectories(My.Application.Info.DirectoryPath, FileIO.SearchOption.SearchTopLevelOnly)
            If My.Computer.FileSystem.GetDirectoryInfo(c).Name.Length = 6 Then
                b += vbTab + "<build>" + My.Computer.FileSystem.GetDirectoryInfo(c).Name + "</build>" + vbCrLf
                i += 1
            End If

        Next
        b += "</builds>"
        My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath + "\builds.xml", b, False)
        Me.TextBox1.Text = "发现了" + i.ToString+"个版本"



        Dim x As New Xml.XmlDocument
        'MsgBox(My.Application.Info.DirectoryPath)
        x.Load(My.Application.Info.DirectoryPath + "\builds.xml")

        For Each y As Xml.XmlNode In x.DocumentElement.SelectNodes("build")
            Me.ComboBox1.Items.Add(y.InnerText)
            Me.ComboBox1.Text = y.InnerText
        Next



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        For Each c As String In My.Computer.FileSystem.GetDirectories(My.Application.Info.DirectoryPath, FileIO.SearchOption.SearchTopLevelOnly)

            If My.Computer.FileSystem.GetDirectoryInfo(c).Name = "Release" Then



                Dim b As String = ""
                Dim d As String = ""
                Dim f As String = ""
                Dim h As String = ""
                Dim l As New System.Collections.Generic.List(Of String)
                b += ("<files build='" + Me.ComboBox1.Text + "' codepage='65001' xmlns:dt='urn:schemas-microsoft-com:datatypes'>" + vbCrLf)
                l.Add(b)
                b = ""
                For Each a As String In My.Computer.FileSystem.GetFiles(c, FileIO.SearchOption.SearchAllSubDirectories)
                    b += vbTab + "<file"

                    b += " name='" + a.Replace(c + "\", "") + "'"

                    Dim bytMD5 As Byte() = New System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(My.Computer.FileSystem.ReadAllBytes(a))

                    d = BitConverter.ToString(bytMD5).Replace("-", String.Empty).ToUpper

                    'b += " md5='" + d + "'"


                    Dim g As New Crc32
                    If My.Computer.FileSystem.ReadAllBytes(a).Length = 0 Then
                        f = Nothing
                    Else

                        f = Convert.ToString(g.CalculateBlock(My.Computer.FileSystem.ReadAllBytes(a)), 16).ToUpper
                    End If


                    b += " crc32='" + f + "'"

                    b += "  dt:dt='bin.base64'>"


                    h = System.Convert.ToBase64String(My.Computer.FileSystem.ReadAllBytes(a))


                    b += h + "</file>" + vbCrLf
                    l.Add(b)
                    b = ""


                Next
                b += "</files>"
                l.Add(b)
                b = Join(l.ToArray)
                My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.GetDirectoryInfo(c).Name + ".xml", b, False)
                Me.TextBox2.AppendText("生成文件:" + My.Computer.FileSystem.GetDirectoryInfo(c).Name + ".xml" + vbCrLf)
            End If
        Next

    End Sub
End Class



Public Class Crc32
    Private Const TABLESIZE As Integer = 256
    Private Const DEFAULTPOLYNOMIAL As Integer = &HEDB88320
    Private Const DEFAULTINITIALVALUE As Integer = &HFFFFFFFF
    Private lookup(TABLESIZE - 1) As Integer
    Private crcPolynomial As Integer = 0
    Public Sub New()
        Me.New(DEFAULTPOLYNOMIAL)
    End Sub
    Public Sub New(ByVal crcPolynomial As Integer)
        Me.crcPolynomial = crcPolynomial
        InitLookupTable()
    End Sub
    Public Property Polynomial() As Integer
        Get
            Return crcPolynomial
        End Get
        Set(ByVal Value As Integer)
            Me.crcPolynomial = value
            InitLookupTable()
        End Set
    End Property
    Public Overloads Function CalculateBlock(ByVal bytes() As Byte) _
                                             As Integer
        Return CalculateBlock(bytes, 0, bytes.Length)
    End Function
    Public Overloads Function CalculateBlock(ByVal bytes() As Byte, _
                                             ByVal index As Integer, _
                                             ByVal length As Integer _
                                            ) As Integer
        Return CalculateBlock(bytes, index, length, DEFAULTINITIALVALUE)
    End Function
    Public Overloads Function CalculateBlock( _
                              ByVal bytes() As Byte, _
                              ByVal index As Integer, _
                              ByVal length As Integer, _
                              ByVal initialValue As Integer) _
                              As Integer
        If bytes Is Nothing Then
            Throw New ArgumentNullException("CalculateBlock(): bytes")
        ElseIf index < 0 Or length <= 0 _
               Or index + length > bytes.Length Then
            Throw New ArgumentOutOfRangeException()
        End If
        Return Not InternalCalculateBlock(bytes, index, _
                                          length, initialValue)
    End Function
    Private Function InternalCalculateBlock( _
                     ByVal bytes() As Byte, _
                     ByVal index As Integer, _
                     ByVal length As Integer, _
                     ByVal initialValue As Integer) _
                     As Integer
        Dim crc As Integer = initialValue
        Dim shiftedCrc As Integer
        Dim position As Integer
        For position = index To length - 1
            shiftedCrc = crc And &HFFFFFF00
            shiftedCrc = shiftedCrc \ &H100
            shiftedCrc = shiftedCrc And &HFFFFFF
            crc = shiftedCrc Xor lookup(bytes(position) Xor _
                                                        (crc And &HFF))
        Next
        Return crc
    End Function
    Public Overloads Function CalculateFile(ByVal path As String) _
                                            As Integer
        Return CalculateFile(path, DEFAULTINITIALVALUE)
    End Function
    Public Overloads Function CalculateFile( _
                              ByVal path As String, _
                              ByVal initialValue As Integer) _
                              As Integer
        If path Is Nothing Then
            Throw New ArgumentNullException("path")
        ElseIf path.Length = 0 Then
            Throw New ArgumentException("Invalid path")
        End If
        Return Not InternalCalculateFile(path, initialValue)
    End Function
    Private Function InternalCalculateFile( _
                     ByVal path As String, _
                     ByVal initialValue As Integer) _
                     As Integer
        Const blockSize As Integer = 4096
        Dim count As Integer
        Dim inStream As IO.FileStream = Nothing
        Dim bytes(blockSize - 1) As Byte
        Dim crc As Integer = initialValue
        Try
            inStream = IO.File.Open(path, IO.FileMode.Open, IO.FileAccess.Read)
            While inStream.Position < inStream.Length
                count = inStream.Read(bytes, 0, blockSize)
                crc = InternalCalculateBlock(bytes, 0, count, crc)
            End While
        Finally
            If Not inStream Is Nothing Then
                inStream.Close()
            End If
        End Try
        Return crc
    End Function
    Private Sub InitLookupTable()
        Dim byteCount, bitCount As Integer
        Dim crc, shiftedCrc As Integer
        For byteCount = 0 To TABLESIZE - 1
            crc = byteCount
            For bitCount = 0 To 7
                shiftedCrc = crc And &HFFFFFFFE
                shiftedCrc = shiftedCrc \ &H2
                shiftedCrc = shiftedCrc And &H7FFFFFFF
                If (crc And &H1) Then
                    crc = shiftedCrc Xor crcPolynomial
                Else
                    crc = shiftedCrc
                End If
            Next
            lookup(byteCount) = crc
        Next
    End Sub
End Class