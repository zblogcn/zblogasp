<%
'///////////////////////////////////////////////////////////////////////////////
'//				Z-Blog
'// 作	 者:   	瑜廷
'// 技术支持:    33195#qq.com
'// 程序名称:    	YT.Build
'// 开始时间:    	2010.12.21
'// 最后修改:    2012.08.24
'// 备	 注:    	only for zblog1.8
'///////////////////////////////////////////////////////////////////////////////

Dim YT_Build_objArticle
Dim YT_Build_objComment

Call RegisterPlugin("YTBuild","ActivePlugin_YT_Build")

Sub ActivePlugin_YT_Build()
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"控制面板",GetCurrentHost&"zb_users/plugin/YTBuild/YT.Panel.asp","nav_quoted","aYTBuildMng",""))
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","YT_Build_Filter_Plugin_PostArticle_Succeed")
End Sub

'卸载插件
Function UnInstallPlugin_YTBuild()
	Dim oTConfig
	Set oTConfig = New TConfig
		oTConfig.Load("YTBuild")
		If oTConfig.Exists("BUILD_HOME") Then oTConfig.Delete
	Set oTConfig = Nothing
End Function
'安装插件
Function InstallPlugin_YTBuild()
	Dim oTConfig
	Set oTConfig = New TConfig
	With oTConfig
		.Load("YTBuild")
		If Not .Exists("BUILD_HOME") Then
			.Write "BUILD_HOME",True
			.Write "BUILD_CATE",True
			.Write "BUILD_TAG",False
			.Write "BUILD_USER",False
			.Write "BUILD_DATE",False
			.Save
		End If
	End With
	Set oTConfig = Nothing
End Function

Function YT_Build_Filter_Plugin_PostArticle_Succeed(ByRef objArticle)
	Dim oTConfig,l,a,b,t,i,s
	Set oTConfig = New TConfig
		oTConfig.Load("YTBuild")
		If oTConfig.Exists("BUILD_HOME") Then
			Dim jbl
			Set jbl = new YTBuildLib
				If oTConfig.Read("BUILD_HOME") Then Call jbl.Default()
				If oTConfig.Read("BUILD_CATE") Then
					s=jbl.Catalog("Cate",objArticle.CateID)
					If Not isEmpty(s) Then
						Set l=cmd.exec(s)
							For Each a In l
								For b=1 To a.intPageCount
									Call jbl.ThreadCatalog(a.Key,a.ID,b)
								Next
							Next
						Set l=Nothing
					End If
				End If
				If oTConfig.Read("BUILD_TAG") Then
					If objArticle.Tag<>"" Then
						s=objArticle.Tag
						s=Replace(s,"}","")
						t=Split(s,"{")
						For i=LBound(t) To UBound(t)
							If t(i)<>"" Then
								s=jbl.Catalog("Tags",t(i))
								If Not isEmpty(s) Then
									Set l=cmd.exec(s)
										For Each a In l
											For b=1 To a.intPageCount
												Call jbl.ThreadCatalog(a.Key,a.ID,b)
											Next
										Next
									Set l=Nothing
								End If
							End If
						Next
					End If
				End If
				If oTConfig.Read("BUILD_USER") Then
					s=jbl.Catalog("Auth",objArticle.AuthorID)
					If Not isEmpty(s) Then
						Set l=cmd.exec(s)
							For Each a In l
								For b=1 To a.intPageCount
									Call jbl.ThreadCatalog(a.Key,a.ID,b)
								Next
							Next
						Set l=Nothing
					End If
				End If
				If oTConfig.Read("BUILD_DATE") Then
					s=jbl.Catalog("Date",objArticle.PostTime)
					If Not isEmpty(s) Then
						Set l=cmd.exec(s)
							For Each a In l
								For b=1 To a.intPageCount
									Call jbl.ThreadCatalog(a.Key,a.ID,b)
								Next
							Next
						Set l=Nothing
					End If
				End If
			Set jbl = Nothing
		End If
	Set oTConfig = Nothing
End Function
%>
<script language="javascript" type="text/javascript" runat="server">
	var cmd={
		exec:function(s){return eval('('+s+')');}
	};
</script>
<!--#include file="YT.Static.asp" -->
<!--#include file="YT.Lib.asp" -->