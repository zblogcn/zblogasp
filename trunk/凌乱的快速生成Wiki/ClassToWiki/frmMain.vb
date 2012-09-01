Imports System.Text.RegularExpressions
Public Class frmMain



    Dim a As Object
    Dim b As Object
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        a = CreateObject("vbscript.regexp")


    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim d, m, n, aaaaaa
        d = CreateObject("vbscript.regexp")
        a.ignorecase = True
        a.global = True
        a.pattern = "class (.+)[\D\d]+?end class"
        m = TextBox1.Text
        n = a.execute(m)
        Dim s

        Dim j
        Dim k As String

        Dim Match
        Dim sb
        For Each s In n
            aaaaaa &= "====" & s.submatches(0).replace(vbCrLf, "").replace(vbLf, "").replace(vbCr, "") & "===="
            aaaaaa &= vbCrLf & "^名称 ^类型 ^参数 ^作用（返回值）^"
            ' aaaaaa &= vbCrLf

            a.ignorecase = True
            a.global = True

            j = s.value
            a.pattern = "public (?!property|function)(.+)"
            b = a.execute(j)
            k = ""

            For Each Match In b
                sb = Match.submatches(0).replace(vbCrLf, "").replace(vbLf, "").replace(vbCr, "")
                k = k & "|" & sb & "|变量| |" & zuoyong(sb) & "|" & vbCrLf
            Next

            a.pattern = "Property Get (.+)"
            b = a.execute(j)

            For Each Match In b
                sb = Match.submatches(0).replace(vbCrLf, "").replace(vbLf, "").replace(vbCr, "")
                k = k & "|" & sb & "|成员| |" & zuoyong(sb) & "|" & vbCrLf
            Next

            a.pattern = "Property Let (.+?)\((.+?)\)"
            b = a.execute(j)
            For Each Match In b
                sb = Match.submatches(0).replace(vbCrLf, "").replace(vbLf, "").replace(vbCr, "")
                k = k & "|" & sb & "|方法|" & Match.submatches(1) & "|" & zuoyong(sb) & "|" & vbCrLf
            Next
            TextBox1.Text = k

            a.pattern = "function (.+?)(\(.+?\)|" & vbCrLf & ")"
            b = a.execute(j)
            For Each Match In b
                sb = Match.submatches(0).replace(vbCrLf, "").replace(vbLf, "").replace(vbCr, "")
                k = k & "|" & sb & "|函数|" & IIf(Match.submatches(1) = vbCrLf, " ", Match.submatches(1)) & "|" & zuoyong(sb) & "|" & vbCrLf
            Next


            a.pattern = "sub (.+?)(\(.+?\)|" & vbCrLf & ")"
            b = a.execute(j)
            For Each Match In b
                sb = Match.submatches(0).replace(vbCrLf, "").replace(vbLf, "").replace(vbCr, "")
                k = k & "|" & sb & "|过程|" & IIf(Match.submatches(1) = vbCrLf, " ", Match.submatches(1)) & "|" & zuoyong(sb) & "|" & vbCrLf
            Next

            aaaaaa &= vbCrLf & k & vbCrLf
        Next

        TextBox1.Text = aaaaaa
    End Sub

    Function zuoyong(str As String)
        Select Case str.ToLower
            Case "id" : Return "ID"
            Case "intro" : Return "摘要"
            Case "content" : Return "内容"
            Case "cateid" : Return "分类ID"
            Case "class_initialize()" : Return "类初始化"
            Case "ip" : Return "IP"
            Case "agent" : Return "User-Agent"
            Case "meta" : Return "Meta类"
            Case "metastring" : Return "Meta字符串"
            Case "order" : Return "排序"
            Case "parentid" : Return "父ID"
            Case "url" : Return "地址"
            Case "fullurl" : Return "完整地址"
            Case "alias" : Return "别名"
            Case "html" : Return "HTML代码"
            Case "del" : Return "删除"
            Case "loadinfobyid" : Return "根据ID读取数据"
            Case "loadinfobyarray" : Return "根据数组读取数据"
            Case "authorid", "userid" : Return "作者ID"
            Case "name" : Return "名字"
            Case "title" : Return "标题"
            Case "save()" : Return "保存"
            Case "refer", "referer" : Return "来源"
            Case "count" : Return "总数"
            Case "post()" : Return "提交"
            Case "level" : Return "文章等级"
            Case "export" : Return "输出"
            Case "build()" : Return "生成"
            Case "templatename" : Return "模板名"
            Case "posttime" : Return "提交时间"
            Case "email" : Return "电子邮件"

            Case Else : Return " "
        End Select
    End Function
End Class
