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
<%

Call OpenConnect()
Call GetConfigs()
BlogConfig.Load("Blog")
BlogUser.Verify()


Dim ac_config,id
Set ac_config=New TConfig
ac_config.Load "AdminColor"
id=ac_config.read("Color4U_"&BlogUser.ID)

If Request.QueryString("color")<>"" Then

	ac_config.Write "Color4U_"&BlogUser.ID,Request.QueryString("color")
	If BlogUser.Level<5 Then
		ac_config.Save
	End If
	Response.Redirect bloghost & "zb_system/cmd.asp?act=admin"

End If

If id="" Then id=0
If id="random" Then
	Randomize
	id=Int((Ubound(BlodColor) - LBound(BlodColor) + 1) * Rnd + LBound(BlodColor))
End If

Response.Expires=0
Response.ContentType = "text/css"

Dim c
''c=LoadFromFile(BlogPath & "zb_system\CSS\admin2.css","utf-8")

'c=Replace(c,"url(../","url(../../../zb_system/")

c=c & "#header .top {background-color:#3a6ea5;}" & vbCrlf
c=c & "input.button,input[type='submit'],input[type='button'] {background-color:#3a6ea5;}" & vbCrlf
c=c & "div.theme-now .betterTip img{box-shadow: 0 0 10px #3a6ea5;}" & vbCrlf

c=c & "#divMain a,#divMain2 a{color:#1d4c7d;}" & vbCrlf

c=c & ".menu ul li a:hover {background-color: #b0cdee;}" & vbCrlf
c=c & "#main .main_left #leftmenu a:hover { background-color: #b0cdee;}" & vbCrlf
c=c & "div.theme-now{background-color:#b0cdee;}" & vbCrlf
c=c & "div.theme-other .betterTip img:hover{border-color:#b0cdee;}" & vbCrlf
c=c & ".SubMenu a:hover {background-color:#b0cdee;}" & vbCrlf
c=c & ".siderbar-header:hover {background-color:#b0cdee;}" & vbCrlf

c=c & "#main .main_left #leftmenu .on a,#main .main_left #leftmenu #on a:hover {background-color:#3399cc;}" & vbCrlf
c=c & "input.button,input[type=""submit""],input[type=""button""] { border-color:#3399cc;}" & vbCrlf
c=c & "input.button:hover {background-color: #3399cc;}" & vbCrlf
c=c & "div.theme-other .betterTip img:hover{box-shadow: 0 0 10px #3399cc;}" & vbCrlf
c=c & ".SubMenu{border-bottom-color:#3399cc;}" & vbCrlf
c=c & ".SubMenu span.m-now{background-color:#3399cc;}" & vbCrlf
c=c & "div #BT_title {background-color: #3399cc;border-color:#3399cc;}" & vbCrlf

c=c & "a:hover { color:#d60000;}" & vbCrlf
c=c & "#divMain a:hover,#divMain2  a:hover{color:#d60000;}" & vbCrlf

'appcenter
c=c & ".tabs { border-bottom-color:#3a6ea5!important;}" & vbCrlf
c=c & ".tabs li a.selected {background-color:#3a6ea5!important;}" & vbCrlf
c=c & "div.heart-vote {background-color:#3a6ea5!important;}" & vbCrlf
c=c & "div.heart-vote ul {border-color:#3a6ea5!important;}" & vbCrlf
c=c & ".install {background-color:#3a6ea5!important;}" & vbCrlf
c=c & ".install:hover{background-color: #3399cc!important;}" & vbCrlf
c=c & "input.button{background-color:#3a6ea5!important;border-color:#3399cc!important;}" & vbCrlf
c=c & "input.button:hover{background-color:#3399cc!important;}" & vbCrlf
c=c & ".themes_body ul li img:hover,.plugin_body ul li img:hover,.main_plugin ul li img:hover,.main_theme ul li img:hover{box-shadow: 0 0 10px #3399cc!important;}" & vbCrlf
c=c & ".left_nav h2,.text h2 {color: #3a6ea5!important;}" & vbCrlf
c=c & ".pagebar span{ background:#3399cc!important; border-color:#3399cc!important;color:#fff;}" & vbCrlf
c=c & ".pagebar span.now-page,.pagebar span:hover{ background:#eee!important;border-color:#eee!important; color:#3399cc!important;}" & vbCrlf



Dim c1,c2,c3,c4,c5
c1="#1d4c7d"
c2="#3a6ea5"
c3="#b0cdee"
c4="#3399cc"
c5="#d60000"

c=Replace(c,c1,BlodColor(id))
c=Replace(c,c2,NormalColor(id))
c=Replace(c,c3,LightColor(id))
c=Replace(c,c4,HighColor(id))
c=Replace(c,c5,AntiColor(id))
'c=Replace(c,"../IMAGE","../../../zb_system/IMAGE")

c=c & vbCrlf & vbCrlf & "/*AdminColor*/" & vbCrlf & "#admin_color{float:right;line-height: 2.5em;font-size: 0.5em;letter-spacing: -0.1em;}"
Response.Write(c) 

Call CloseConnect()
%>