<%
'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
Dim GuetsBook_ID

'注册插件
Call RegisterPlugin("ZBQQweibo","ActivePlugin_ZBQQweibo")
'挂口部分
Function ActivePlugin_ZBQQweibo()

	Dim Config
	Set Config=New TConfig
	Config.Load "ZBQQweibo"
	GuetsBook_ID=CInt(Config.Read("g"))

	Call Add_Action_Plugin("Action_Plugin_TArticle_Url","If ID=GuetsBook_ID Then Url=ZC_BLOG_HOST & ""ZBQQweibo.asp"":Exit Property")

	Call Add_Action_Plugin("Action_Plugin_TArticle_Save_Begin","If ID=GuetsBook_ID Then Exit Function")
	
End Function


%>