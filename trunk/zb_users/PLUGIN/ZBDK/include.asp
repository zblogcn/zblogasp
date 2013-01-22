<%
Call ZBDK.submenu.add("default","首页",BlogHost & "zb_users/plugin/ZBDK/main.asp","m-left")
%>
<!-- #include file="inc_func.asp"-->
<%

Dim sAction_Plugin_ZBDK_Else
Dim Action_Plugin_ZBDK_Else()
Redim Action_Plugin_ZBDK_Else(0)
Dim Response_Plugin_ZBDK_Default




'注册插件
Call RegisterPlugin("ZBDK","ActivePlugin_ZBDK")
'挂口部分

Function ActivePlugin_ZBDK()

	
	Call Add_Action_Plugin("Action_Plugin_ZBDK_Else","Call Add_Response_Plugin(""Response_Plugin_Admin_Top"",MakeTopMenu(1,""开发工具"",BlogHost&""zb_users/plugin/ZBDK/main.asp"",""zbdk"",""""))")
	

	
	Call ZBDK_SomethingElse
	
	
	
	
End Function

Function ZBDK_SomethingElse()

	On Error Resume Next
	For Each sAction_Plugin_ZBDK_Else In Action_Plugin_ZBDK_Else
		If Not IsEmpty(sAction_Plugin_ZBDK_Else)  Then Call Execute(sAction_Plugin_ZBDK_Else)
	Next
	
End Function





Function ZBDK_AddStatus(id,msg,url,text)
	Call ZBDK.submenu.add(id,msg,url,"m-left")
	Stop
	Call ZBDK.mainpage.add(id,url,text)
End Function






Function ZBDK_GetASPCode(name_,isExecute)
	Dim strTemp
	strTemp=LoadFromFile(BlogPath &"zb_users\plugin\zbdk\" & name_,"utf-8")
	strTemp=Replace(strTemp,"<"&"%","")
	strTemp=Replace(strTemp,"%"&">","")
	ZBDK_GetASPCode=strTemp
	
	If InStr(strTemp,"<!-- #"&"include file=""") Then 
		
		Dim strInclude,aryInc,oRegExp
		Set oRegExp=New RegExp
		oRegExp.Pattern="<!-- ?#include file=""(.+?)"" ?-->"
		oRegExp.Global=True
		oRegExp.IgnoreCase=True
		Set aryInc=oRegExp.Execute(strTemp)
		For Each strInclude In aryInc
			Call ZBDK_GetASPCode(strInclude.SubMatches(0),True)
		Next
		ZBDK_GetASPCode=oRegExp.Replace(strTemp,"")
		
	End If
	If isExecute Then Execute ZBDK_GetASPCode
	
End Function

Function ZBDK_ScanPluginInclude()
	Dim s,i
	s=Split(ZBDK_GetASPCode("tools_func.asp",False),vbCrlf)
	For i=1 To Ubound(s)
		Call ZBDK_GetASPCode(s(i) & "\include.asp",True)
	Next
End Function




%>
<script language="javascript" runat="server">
var ZBDK={
	submenu:{
		list:{},
		add:function(id,msg,url,css){
			var o={
				"id":id,
				"msg":msg,
				"url":url,
				"css":css	
			};
			ZBDK.submenu.list[id]=o;
			return o;
		}
		,
		remove:function(id){delete ZBDK.submenu.list[id];}
		,
		"export":function(id){
			var st="";
			for(var lst in ZBDK.submenu.list){
				st+=MakeSubMenu(ZBDK.submenu.list[lst].msg,ZBDK.submenu.list[lst].url,ZBDK.submenu.list[lst].css+((ZBDK.submenu.list[lst].id==id)?" m-now":""),false);
			}
			return st;
		}
	},
	mainpage:{
		list:{},
		add:function(id,url,text){
			var o={
				"id":id,
				"url":url,
				"text":text	
			};
			ZBDK.mainpage.list[id]=o;
			return o;
		}
		,
		remove:function(id){
			delete ZBDK.mainpage.list[id];
		}
		,
		"export":function(id){
			var st="",c=0;
			for(var lst in ZBDK.mainpage.list){
				c++;
				st+="<tr height='40'><td>"+c+"</td><td><a href='"+ZBDK.mainpage.list[lst].url+"'>"+ZBDK.mainpage.list[lst].id+"</a></td><td>"+ZBDK.mainpage.list[lst].text+"</td></tr>";
			}
			return st;
		}
	}
}

</script>