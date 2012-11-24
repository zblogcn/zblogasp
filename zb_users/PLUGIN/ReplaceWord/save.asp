<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<!-- #include file="function.asp" -->

<%
Dim XmlDom
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
		Set XmlDom=replaceword.words(id)
		replaceword.xmldom.documentElement.removeChild xmlDom
	Case Else
		Stop
		Dim Frm,id,id2,objDom
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
replaceword.xmldom.Save(Server.MapPath("config.xml"))
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