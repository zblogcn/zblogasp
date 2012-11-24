<!--#include file="function.asp" -->
<%
Const qfminiwidth = 129
          '相关文章缩略图宽度
Const qfminiheight = 84
         '相关文章缩略图高度

Call RegisterPlugin("QF_Export_Mutuality","ActivePlugin_QF_Export_Mutuality")
Function ActivePlugin_QF_Export_Mutuality()
Call Add_Action_Plugin("Action_Plugin_TArticle_Export_Mutuality_Begin","Export_Mutuality=QF_Export_Mutuality(Disable_Export_Mutuality,Template_Article_Mutuality,ID,Tag):Exit Function")
End Function

Function QF_Export_Mutuality(Disable_Export_Mutuality,Template_Article_Mutuality,ID,Tag)

		If Disable_Export_Mutuality=True Then Exit Function

		If ZC_MUTUALITY_COUNT=0 Then 
			QF_Export_Mutuality=True
			Exit Function
		End If

		If Tag<>"" Then

			Dim strCC_Count,strCC_ID,strCC_Name,strCC_Url,strCC_PostTime,strCC_Title,strCC_Img
			Dim strCC
			Dim i,j,s
			Dim objRS
			Dim strSQL

			Set objRS=Server.CreateObject("ADODB.Recordset")

			strSQL="SELECT TOP "& ZC_MUTUALITY_COUNT &" [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>2) AND [log_ID]<>"& ID
			strSQL = strSQL & " AND ("

			Dim aryTAGs
			s=Replace(Tag,"}","")
			aryTAGs=Split(s,"{")

			For j = LBound(aryTAGs) To UBound(aryTAGs)
				If aryTAGs(j)<>"" Then
					strSQL = strSQL & "([log_Tag] Like '%{"&FilterSQL(aryTAGs(j))&"}%')"
					If j=UBound(aryTAGs) Then Exit For
					If aryTAGs(j)<>"" Then strSQL = strSQL & " OR "
				End If
			Next

			strSQL = strSQL & ")"
			strSQL = strSQL + " ORDER BY [log_PostTime] DESC "

			Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			objRS.Source=strSQL
			objRS.Open()
			If (Not objRS.bof) And (Not objRS.eof) Then

				Dim objArticle
				For i=1 To ZC_MUTUALITY_COUNT '相关文章数目，可自行设定

					Set objArticle=New TArticle

					If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17)))  Then

						strCC_Count=strCC_Count+1
						strCC_ID=objArticle.ID
						strCC_Url=objArticle.Url
						strCC_PostTime=objArticle.PostTime
						strCC_Title=objArticle.Title
						strCC_Img=QF_CreatMini(QF_GetImgSrc(objArticle.Content),qfminiwidth,qfminiheight)

						strCC=GetTemplate("TEMPLATE_B_ARTICLE_MUTUALITY")

						strCC=Replace(strCC,"<#article/mutuality/id#>",strCC_ID)
						strCC=Replace(strCC,"<#article/mutuality/url#>",strCC_Url)
						strCC=Replace(strCC,"<#article/mutuality/posttime#>",strCC_PostTime)
						strCC=Replace(strCC,"<#article/mutuality/name#>",strCC_Title)
							strCC=Replace(strCC,"<#article/mutuality/img#>",strCC_Img)

						Template_Article_Mutuality=Template_Article_Mutuality & strCC

					End If

					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next

			End if

			objRS.Close()
			Set objRS=Nothing

		End If

		QF_Export_Mutuality=True
End Function






%>