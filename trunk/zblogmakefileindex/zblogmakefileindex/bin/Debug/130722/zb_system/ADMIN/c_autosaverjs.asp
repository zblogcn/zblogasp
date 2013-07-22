<%@CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    zsx
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_autosaverjs.asp
'// 开始时间:    2006-1-19
'// 最后修改:    2006-7-27
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->


<%
'
Call System_Initialize()

Public ZC_AUTOSAVE_FILENAME,ZC_AUTOSAVE_FILEMODIFIED
ReGetFile
Sub ReGetFile()
	ZC_AUTOSAVE_FILENAME="autosave_"&MD5("autosave"&"_"&MD5(ZC_BLOG_CLSID & BlogUser.ID))&".txt"  
	ZC_AUTOSAVE_FILEMODIFIED=GetFileModified(BlogPath&"zb_users/cache/"&ZC_AUTOSAVE_FILENAME)
End Sub 

If BlogUser.Level>3 Then
	Response.Write ZVA_ErrorMsg(6)
	Response.End 
End If

Select Case Request.QueryString("act")
	Case "restore"
		Response.ContentType="text/plain"
		Restore
	Case "save"
		Response.ContentType="text/plain"
		SaveContent
	Case "del"
		Response.ContentType="text/plain"
		DeleteBackup
	Case Else
		Response.ContentType="application/javascript"
		ExportAutoSaveJS
End Select

'*********************************************************
' 目的：    Convert Bytes To Str
'*********************************************************
Function BytesToBstr(body,Cset)
		On Error Resume Next
		Dim objstream
		Set objstream = Server.CreateObject("adodb.stream")
		objstream.Type = 1
		objstream.Mode =3
		objstream.Open
		objstream.Write body
		objstream.Position = 0
		objstream.Type = 2
		objstream.Charset = Cset
		BytesToBstr = objstream.ReadText 
		objstream.Close
		Set objstream = Nothing
End Function

'*********************************************************
' 目的：    Save Draft And DisPlay
'*********************************************************
Sub SaveContent()
	Dim strJSON
	strJSON="{""title"":"""&jsEncode(Request.Form("title"))&""",""alias"":"""&jsEncode(Request.Form("alias"))&""","
	strJSON=strJSON & """tag"":"""&jsEncode(Request.Form("tag"))&""",""cate"":"""&jsEncode(Request.Form("cate"))&""",""content"":"""&jsEncode(Request.Form("content"))&""",""success"":true}"
	Call SaveToFile(BlogPath & "ZB_USERS/CACHE/"&ZC_AUTOSAVE_FILENAME,strJSON,"utf-8",False)
	
	ReGetFile
	Response.Write "{'result':'<span style="""">&nbsp;"&formatdatetime(now,4)&":"&Right("0"&second(now),2)&"<a href=""javascript:try{autosave.view()}catch(e){};"" target=""_blank"" style=""text-decoration: none;"">"&ZC_MSG258&"</a>&nbsp;</span>','file':{'name':'"&ZC_AUTOSAVE_FILENAME&"','modified':'"&ZC_AUTOSAVE_FILEMODIFIED&"'}}"
	Response.End
End Sub

Sub Restore()
	If ZC_AUTOSAVE_FILEMODIFIED=Now Then
		Response.Write "{'title':'"&ZC_MSG180&"','alias':'','tag':'','cate':0,'content':'"&ZC_MSG133&"','success':false}"
	Else
		Response.Write LoadFromFile(BlogPath & "ZB_USERS/CACHE/"&ZC_AUTOSAVE_FILENAME,"utf-8")
	End If
End Sub

Sub DeleteBackup()
	Call DelToFile(BlogPath & "ZB_USERS/CACHE/"&ZC_AUTOSAVE_FILENAME)
End Sub

Function jsEncode(str)
	Dim charmap(127), haystack()
	charmap(8)  = "\b"
	charmap(9)  = "\t"
	charmap(10) = "\n"
	charmap(12) = "\f"
	charmap(13) = "\r"
	charmap(34) = "\"""
	charmap(47) = "\/"
	charmap(92) = "\\"
	Dim strlen : strlen = Len(str) - 1
	ReDim haystack(strlen)
	Dim i, charcode
	For i = 0 To strlen
		haystack(i) = Mid(str, i + 1, 1)
		charcode = AscW(haystack(i)) And 65535
		If charcode < 127 Then
			If Not IsEmpty(charmap(charcode)) Then
				haystack(i) = charmap(charcode)
			ElseIf charcode < 32 Then
				haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
			End If
		Else
			haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
		End If
	Next

	jsEncode = Join(haystack, "")
End Function

'*********************************************************
' 目的：   输出自动保存脚本
'*********************************************************
Function ExportAutoSaveJS()
	Response.Clear
	'//////////////
%>
var autosave = {
    file: {
        name: "<%=ZC_AUTOSAVE_FILENAME%>",
        modified: "<%=ZC_AUTOSAVE_FILEMODIFIED%>"

    },
    time: {
        max: 60,
        remain: 60

    },
    elements: {
        msg: $("#msg"),
        time: $("#timemsg")


    },
    save: function() {
        if (editor_api.editor.content.get() == "") {
            autosave.elements.msg.html("<%=ZC_MSG256%>");
            return false
        }
        $.post(bloghost+"zb_system/admin/c_autosaverjs.asp?act=save", {
            title: $("#edtTitle").val(),
            alias: $("#edtAlias").val(),
            tag: $("#edtTag").val(),
            cate: $("#cmbCate").val(),
            content: editor_api.editor.content.get()

        },
        function(data) {
            var m = eval("(" + data + ")");
            console.log(m)
            autosave.elements.msg.html(m.result);
            autosave.file.name = m.file.name;
            autosave.file.modified = m.file.modified;

        })

    },
    restore: function() {
        $.get(bloghost+"zb_system/admin/c_autosaverjs.asp", {
            act: "restore",
            rnd: Math.random()
        },
        function(data) {
            var m = eval("(" + data + ")");
            if (m.success) {
                $("#edtTitle").val(m.title);
                $("#edtAlias").val(m.alias);
                $("#edtTag").val(m.tag);
                $("#cmbCate").val(m.cate);
                editor_api.editor.content.put(m.content);

            }

        });

    },
    view: function() {
        var r = Math.floor(Math.random() * 100);
        var o = "<div id='autosave_get" + r + "'><p><%=ZC_MSG117%></p></div>";
        $("#divMain2").append(o);
        var k = $("#autosave_get" + r).dialog({
            title: "<%=ZC_MSG017%>",
            modal: true

        });
        $.get(bloghost+"zb_system/admin/c_autosaverjs.asp", {
            act: "restore",
        	rnd: Math.random()
       	},
        function(data) {
            var m = eval("(" + data + ")"),
            s = "";
            s += "<p><span style='font-weight:bold'><%=ZC_MSG060%>：</span>" + m.title + "</p>";
            s += "<p><span style='font-weight:bold'><%=ZC_MSG147%>：</span>" + m.alias + "</p>";
            s += "<p><span style='font-weight:bold'><%=ZC_MSG138%>：</span>" + m.tag + "</p>";
            s += "<p><span style='font-weight:bold'><%=ZC_MSG012%>ID：</span>" + m.cate + "</p>";
            s += "<p><a href='javascript:;' onclick='autosave.runcode(" + r + ")'><span style='font-weight:bold'><%=ZC_MSG090%>：</span></a><div id='autosave_content" + r + "'>" + m.content + "</div></p>";
            k.html(s)

        });

    },
    timer: function() {
        autosave.time.remain--;
        autosave.elements.time.html(autosave.time.remain + "<%=ZC_MSG251%>");
        if (autosave.time.remain >= 0) {
            window.setTimeout("autosave.timer()", 1000);

        } else {
            if (autosave.time.remain <= -1000) {
                autosave.time.remain = autosave.time.max;
                autosave.timer();

            } else {
                autosave.elements.time.html("<%=ZC_MSG250%>");
                autosave.save();
                autosave.time.remain = autosave.time.max;
                autosave.timer();

            }

        }

    },
    runcode: function(obj) {
        var winname = window.open('', "_blank", '');
        winname.document.open('text/html', 'replace');
        winname.opener = null;
        winname.document.write($('#autosave_content' + obj).html());
        winname.document.close();

    },
    del: function() {
        $.get(bloghost+"zb_system/admin/c_autosaverjs.asp?act=del");
        autosave.file.name = "";
        autosave.file.modified = "";
        autosave.elements.msg.html("<%=ZC_MSG228%>")

    }

}

$(document).ready(function() {
    document.getElementById("msg2").innerHTML = "&nbsp;&nbsp;<a href='javascript:;' onclick='autosave.view();return false' style='cursor:hand;'>[<%=ZC_MSG015%>]</a>&nbsp;&nbsp;<a href='javascript:;' onclick='if(confirm(\"<%=ZC_MSG254%>\")) autosave.restore();return false;' style='cursor:hand;'>[<%=ZC_MSG252%>]</a>&nbsp;&nbsp;<a href='javascript:;' onclick='autosave.save();return false' style='cursor:hand;'>[<%=ZC_MSG004%>]</a>&nbsp;&nbsp;<a href='javascript:;' onclick='autosave.del();return false' style='cursor:hand;'>[<%=ZC_MSG063%>]</a>";
    <%If ZC_AUTOSAVE_FILEMODIFIED <>Now Then Response.Write "document.getElementById('msg').innerHTML='" & Replace(ZC_MSG102,"%s",ZC_AUTOSAVE_FILEMODIFIED) & "';"%>
    autosave.timer();
    $("#edit").submit(function() {
        autosave.del()
    })

});

<%
End Function

Call System_Terminate()
%>
