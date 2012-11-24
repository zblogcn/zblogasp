<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->

<%
Dim Frm,id,id2,objDom
Dim i,str
ShowError_Custom="Response.Write ""{'success':false,'error':""&id&""}"":Response.End"
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ReplaceWord")=False Then Call ShowError(48)
BlogTitle="敏感词替换器"
replaceword.init()
Select Case Request.QueryString("act")
	Case "delete"
		Set objDom=replaceword.words(id)
		replaceword.xmldom.documentElement.removeChild objDom
	Case "import"
		On Error Resume Next
		Select Case Request.Form("type")
			Case 2
				replaceword.xmldom.loadXML "<?xml version=""1.0"" encoding=""utf-8""?><words></words>"
		End Select
		str=Request.Form("txaContent")
		str=Replace(str,vbLf,vbCr)
		str=Replace(str,vbCr,vbCrlf)
		str=Split(str,vbCrlf)
		For i=0 To Ubound(str)
			Frm=Split(str(i),"====")
			If Ubound(Frm)=3 Then
				Set objDom=replaceword.create(i)
				objDom.attributes.getNamedItem("regexp").value=IIf(Frm(0)="1","True","False")
				objDom.selectSingleNode("str").text=Frm(1)
				objDom.selectSingleNode("replace").text=Frm(2)
				objDom.selectSingleNode("description").text=Frm(3)
			End If
		Next
	Case Else

		For Each Frm In Request.Form
			id=Split(Frm,"_")(1)
			id2=Left(Frm,3)
			If id2="new" Then
				id2=id
				id=Split(Frm,"_")(2)		
				Set objDom=replaceword.create(id)
			Else
				Set objDom=replaceword.words(id)
			End If
			Select Case id2
				Case "exp"
					objDom.attributes.getNamedItem("regexp").value=Request.Form(Frm).Item
				Case "str"
					objDom.selectSingleNode("str").text=Request.Form(Frm).Item
				Case "rep"
					objDom.selectSingleNode("replace").text=Request.Form(Frm).Item
				Case "des"
					objDom.selectSingleNode("description").text=Request.Form(Frm).Item
				Case "del"
					If Request.Form(Frm)="True" Then
						replaceword.del_.push CLng(id)
					End If
			End Select
		Next
End Select
replaceword.del()
If replaceword.xmldom.documentElement.selectNodes("antidownload").length<=0 Then
	Stop
	
	Set objDom=Server.CreateObject("Microsoft.XMLDom")
	objDom.LoadXML "  <antidownload>"& vbCrlf & "    <![CDATA["& vbCrlf & "      <" & "% 'On Error Resume Next %" & ">"& vbCrlf & "      <" & "% Response.Charset=""UTF-8"" %" & ">"& vbCrlf & "      <!-"&"- #include file=""..\..\c_option.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\..\..\zb_system\function\c_function.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\..\..\zb_system\function\c_system_lib.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\..\..\zb_system\function\c_system_base.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\..\..\zb_system\function\c_system_event.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\..\..\zb_system\function\c_system_manage.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\..\..\zb_system\function\c_system_plugin.asp"" -->"& vbCrlf & "      <!-"&"- #include file=""..\p_config.asp"" -->"& vbCrlf & "	  <" & "%System_Initialize:If BlogUser.Level>1 Then Call ShowError(6)%" & ">"& vbCrlf & "	  ]]>"& vbCrlf & "  </antidownload>"
	replaceword.xmldom.documentElement.appendChild(objDom.documentElement)
End If
replaceword.xmldom.Save(Server.MapPath("config.asp"))
Response.Write "{'success':true}"
'Response.Redirect "main.asp"
%>
<script language="javascript" runat="server">
replaceword["new_id"]=[];
replaceword["dom"]=[];
replaceword["del_"]=[];
replaceword["del"]=function(){
	var id=0;
	  replaceword.del_.sort();
	for(var i=0; i<=replaceword.del_.length-1; i++){
		id=replaceword.del_[i];
		//id=id-i;
		replaceword.xmldom.documentElement.removeChild(replaceword.words(id))
	}
}
replaceword["create"]=function(id){
	for(var i=0; i<=replaceword.new_id.length-1; i++){
		if(replaceword.new_id[i]==id){return replaceword.dom[i]}
	}
	replaceword.new_id.push(id);
	var objDom=replaceword.xmldom.createElement("word");
	objDom.setAttribute("user",BlogUser.Level);
	objDom.setAttribute("regexp","False");
	objDom.appendChild(replaceword.xmldom.createElement("str"));
	objDom.appendChild(replaceword.xmldom.createElement("replace"));
	objDom.appendChild(replaceword.xmldom.createElement("description"));
	replaceword.xmldom.documentElement.appendChild(objDom);
	replaceword.dom.push(objDom);
	return objDom;
}
</script>