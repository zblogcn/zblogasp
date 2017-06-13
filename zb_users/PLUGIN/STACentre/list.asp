<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Dim RootPath
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)
BlogTitle="静态管理中心"
RootPath=Replace(BlogPath,Replace(CookiesPath(),"/","\"),"\")

Dim tempString
If Request("mak")="1" Then
	Call SaveToFile(RootPath & "httpd.ini",MakeIIS6Rewrite2(),"iso-8859-1",False)
	Call SetBlogHint_Custom("创建httpd.ini成功!")
End If
If Request("mak")="2" Then
	Call SaveToFile(RootPath & ".htaccess",MakeIIS6Rewrite3(),"utf-8",False)
	Call SetBlogHint_Custom("创建.htaccess成功!")
End If
If Request("mak")="3" Then
	Call SaveToFile(BlogPath & "web.config",MakeIIS7UrlRewrite(),"utf-8",False)
	Call SetBlogHint_Custom("创建web.config成功!")
End If
If Request("add")="1" Then
	tempString=LoadFromFile(RootPath & "httpd.ini","iso-8859-1")
	If InStr(tempString,"[ISAPI_Rewrite]") Then tempString=Split(tempString,"[ISAPI_Rewrite]")(1)
	Call SaveToFile(RootPath & "httpd.ini",MakeIIS6Rewrite2() & vbCrLf & tempString,"iso-8859-1",False)
	Call SetBlogHint_Custom("追加httpd.ini成功!")
End If
If Request("del")="1" Then
	Call DelToFile(BlogPath & "httpd.ini")
	Call SetBlogHint_Custom("删除httpd.ini成功!")
End If
If Request("del")="2" Then
	Call DelToFile(BlogPath & ".htaccess")
	Call SetBlogHint_Custom("删除.htaccess成功!")
End If
If Request("del")="3" Then
	Call DelToFile(BlogPath & "web.config")
	Call SetBlogHint_Custom("删除web.config成功!")
End IF

Function FileExists(Path)
	FileExists=0
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(Path) Then
		FileExists=1
	End If
	Set fso=Nothing
End Function
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
pre {
	border: 1px solid #ededed;
	margin: 0px;
}
</style>
<script type="text/javascript">
function showmsg(int){
	var ary=[<%=FileExists(RootPath&"httpd.ini")%>,<%=FileExists(RootPath&".htaccess")%>,<%=FileExists(BlogPath&"web.config")%>];
	if(ary[2]===1&&int==3){
		if(!confirm("检测到您的网站web.config文件存在，是否要继续操作？")) return false
	}
	else if(ary[0]===1&&int==1){
		if(!confirm("检测到您的根目录下httpd.ini文件存在，是否要继续操作？")) return false
	}
	else if(ary[1]===1&&int==2){
		if(!confirm("检测到您的根目录下.htaccess文件存在，是否要继续操作？")) return false
	}
	return true
}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <a href="main.asp"><span class="m-left">配置页面</span></a><a href="list.asp"><span class="m-left m-now">ReWrite规则</span></a> <a href="help.asp"><span class="m-right">帮助</span></a></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <%If ZC_POST_STATIC_MODE="REWRITE" Or ZC_STATIC_MODE="REWRITE" Then%>
            <div class="content-box"><!-- Start Content Box -->
              
              <div class="content-box-header">
                <ul class="content-box-tabs">
                  <li><a href="#tab1" <%=IIf(Not CheckRegExp(Request.ServerVariables("SERVER_SOFTWARE"),"Microsoft-IIS/[56]"),"","class=""default-tab""")%> ><span>IIS6+ISAPI Rewrite 2.X</span></a></li>
                  <li><a href="#tab2"><span>IIS6+ISAPI Rewrite 3.X</span></a></li>
                  <li><a href="#tab3" <%=IIf(CheckRegExp(Request.ServerVariables("SERVER_SOFTWARE"),"Microsoft-IIS/[56]"),"","class=""default-tab""")%> ><span>IIS7、7.5、8+Url Rewrite</span></a></li>
                </ul>
                <div class="clear"></div>
              </div>
              <!-- End .content-box-header -->
              
              <div class="content-box-content">
                <div class="tab-content <%=IIf(Not CheckRegExp(Request.ServerVariables("SERVER_SOFTWARE"),"Microsoft-IIS/[56]"),"","default-tab")%> " style='border:none;padding:0px;margin:0;' id="tab1">
                  <textarea style="width:80%;height:300px" readonly>
<%=TransferHTML(MakeIIS6Rewrite2(),"[html-format]")%>
</textarea>
                  <hr/>
                  <p>
                    <input type="button" onClick="if(showmsg(1)){window.location.href='?mak=1'}" value="创建httpd.ini" />
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" onClick="if(showmsg(1)){window.location.href='?del=1'}" value="删除httpd.ini" />
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" onClick="if(showmsg(1)){window.location.href='?add=1'}" value="追加httpd.ini" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<span class="star">请在网站根目录创建httpd.ini文件并把相关内容复制进去,httpd.ini文件必须为ANSI编码,也可以点击按钮生成.</span></p>
                </div>
                <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
                  <textarea style="width:80%;height:300px" readonly>
<%=TransferHTML(MakeIIS6Rewrite3(),"[html-format]")%>
</textarea>
                  <hr/>
                  <p>
                    <input type="button" onClick="if(showmsg(2)){window.location.href='?mak=2'}" value="创建.htaccess" />
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" onClick="if(showmsg(2)){window.location.href='?del=2'}" value="删除.htaccess" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<span class="star">请在网站根目录创建.htaccess文件并把相关内容复制进去,也可以点击按钮生成..</span></p>
                </div>
                <div class="tab-content <%=IIf(CheckRegExp(Request.ServerVariables("SERVER_SOFTWARE"),"Microsoft-IIS/[56]"),"","default-tab")%> " style='border:none;padding:0px;margin:0;' id="tab3">
                  <textarea style="width:80%;height:300px" readonly>
<%=TransferHTML(MakeIIS7UrlRewrite(),"[html-format]")%>
</textarea>
                  <hr/>
                  <p>
                    <input type="button" onClick="if(showmsg(3)){window.location.href='?mak=3'}" value="创建web.config" />
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" onClick="if(showmsg(3)){window.location.href='?del=3'}" value="删除web.config" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<span class="star">请在网站<u>"当前目录"</u>创建web.config文件并把相关内容复制进去,也可以点击按钮生成..</span></p>
                </div>
              </div>
              <!-- End .content-box-content --> 
              
            </div>
            <!-- End .content-box -->
            <%Else%>
            <hr/>
            <p><b>文章及页面和分类页都没有启用动态模式+Rewrite支持,所以没有可用规则.</b></p>
            <%End If%>
            <p><br/>
            </p>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->




<%
Function FormatUrl(ByRef url,page)
	If InStr(ZC_DEFAULT_REGEX,"{%page%}")>0 And InStr(url,"{%page%}")=0 Then
		url=Replace(url,"/default.html",IIF(page=True,"/{%page%}/default.html","/"))
	End If
	url=Replace(url,"{%page%}/default.html",IIF(page=True,"{%page%}","")&"/")
	url=Replace(url,"/default.html",IIF(page=True,"_{%page%}","")&"/")
	url=Replace(url,".html",IIF(page=True,"_{%page%}","")&".html")
End Function

Function AddRule1(s,regex,page)
	Dim t,r
	If regex="ZC_DEFAULT_REGEX" Then
		If InStr(ZC_DEFAULT_REGEX,"{%page%}")=0 Then
			r=Replace(ZC_DEFAULT_REGEX,".html","_{%page%}\.html")
		Else
			r=Replace(ZC_DEFAULT_REGEX,"default.html","")
			' r=ZC_DEFAULT_REGEX
		End If
		t=t & "RewriteRule "& r &" {%host%}/catalog\.asp\?page=$1" & vbCrlf
	End If


	If regex="ZC_CATEGORY_REGEX" Then
		r=ZC_CATEGORY_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule "& r &" {%host%}/catalog\.asp\?cate=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If

	If regex="ZC_USER_REGEX" Then
		r=ZC_USER_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule "& r &" {%host%}/catalog\.asp\?auth=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If


	If regex="ZC_TAGS_REGEX" Then
		r=ZC_TAGS_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule "& r &" {%host%}/catalog\.asp\?tags=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If

	If regex="ZC_DATE_REGEX" Then
		r=ZC_DATE_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule "& r &" {%host%}/catalog\.asp\?date=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If


	If regex="ZC_ARTICLE_REGEX" Then
		r=ZC_ARTICLE_REGEX
		r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
		r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
		t=t & "RewriteRule "& r &" {%host%}/view\.asp\?id=$1" & vbCrlf
	End If


	If regex="ZC_PAGE_REGEX" Then
		r=ZC_PAGE_REGEX
		r=Replace(r,"{%name%}","{%alias%}")
		r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
		r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
		t=t & "RewriteRule "& r &" {%host%}/view\.asp\?id=$1" & vbCrlf
	End If


	t=Replace(t,"{%host%}/",CookiesPath())
	t=Replace(t,"{%name%}","(?!zb_)(.*)")
	t=Replace(t,"{%alias%}","(?!zb_)(.*)")
	t=Replace(t,"{%id%}","([0-9]+)")
	t=Replace(t,"{%date%}","([0-9\-]+)")
	t=Replace(t,"{%post%}",ZC_STATIC_DIRECTORY)
	t=Replace(t,"{%category%}","(?!zb_).*")
	t=Replace(t,"{%author%}","(?!zb_).*")
	t=Replace(t,"{%year%}","[0-9\-]+")
	t=Replace(t,"{%month%}","[0-9\-]+")
	t=Replace(t,"{%day%}","[0-9\-]+")
	t=Replace(t,"{%page%}","([0-9]+)")



	AddRule1=s & t

End Function

Function MakeIIS6Rewrite2()
		Dim s

	s=s & "[ISAPI_Rewrite]" & vbCrlf

	s=s & "" & vbCrlf

	If ZC_STATIC_MODE="REWRITE" Then

		s = AddRule1(s,"ZC_DEFAULT_REGEX",True)
		s = AddRule1(s,"ZC_CATEGORY_REGEX",True)
		s = AddRule1(s,"ZC_CATEGORY_REGEX",False)
		s = AddRule1(s,"ZC_USER_REGEX",True)
		s = AddRule1(s,"ZC_USER_REGEX",False)
		s = AddRule1(s,"ZC_TAGS_REGEX",True)
		s = AddRule1(s,"ZC_TAGS_REGEX",False)
		s = AddRule1(s,"ZC_DATE_REGEX",True)
		s = AddRule1(s,"ZC_DATE_REGEX",False)

	End If


	If ZC_POST_STATIC_MODE="REWRITE" Then

		s = AddRule1(s,"ZC_ARTICLE_REGEX",False)
		s = AddRule1(s,"ZC_PAGE_REGEX",False)

	End If


		MakeIIS6Rewrite2=s
End Function

Function AddRule2(s,regex,page)
	Dim t,r
	If regex="ZC_DEFAULT_REGEX" Then
		If InStr(ZC_DEFAULT_REGEX,"{%page%}")=0 Then
			r=Replace(ZC_DEFAULT_REGEX,".html","_{%page%}\.html")
		Else
			r=Replace(ZC_DEFAULT_REGEX,"default.html","")
			' r=ZC_DEFAULT_REGEX
		End If
		t=t & "RewriteRule ^"& r &"$ /catalog.asp\?page=$1" & vbCrlf
	End If


	If regex="ZC_CATEGORY_REGEX" Then
		r=ZC_CATEGORY_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule ^"& r &"$ /catalog.asp\?cate=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If

	If regex="ZC_USER_REGEX" Then
		r=ZC_USER_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule ^"& r &"$ /catalog.asp\?auth=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If


	If regex="ZC_TAGS_REGEX" Then
		r=ZC_TAGS_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule ^"& r &"$ /catalog.asp\?tags=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If

	If regex="ZC_DATE_REGEX" Then
		r=ZC_DATE_REGEX
		Call FormatUrl(r,page)
		t=t & "RewriteRule ^"& r &"$ /catalog.asp\?date=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If


	If regex="ZC_ARTICLE_REGEX" Then
		r=ZC_ARTICLE_REGEX
		r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
		r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
		t=t & "RewriteRule ^"& r &"$ /view.asp\?id=$1 [NU]" & vbCrlf
	End If


	If regex="ZC_PAGE_REGEX" Then
		r=ZC_PAGE_REGEX
		r=Replace(r,"{%name%}","{%alias%}")
		r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
		r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
		t=t & "RewriteRule ^"& r &"$ /view.asp\?id=$1 [NU]" & vbCrlf
	End If


	t=Replace(t,"{%host%}/","")
	t=Replace(t,"{%name%}","(?!zb_)(.*)")
	t=Replace(t,"{%alias%}","(?!zb_)(.*)")
	t=Replace(t,"{%id%}","([0-9]+)")
	t=Replace(t,"{%date%}","([0-9\-]+)")
	t=Replace(t,"{%post%}",ZC_STATIC_DIRECTORY)
	t=Replace(t,"{%category%}","(?!zb_).*")
	t=Replace(t,"{%author%}","(?!zb_).*")
	t=Replace(t,"{%year%}","[0-9\-]+")
	t=Replace(t,"{%month%}","[0-9\-]+")
	t=Replace(t,"{%day%}","[0-9\-]+")
	t=Replace(t,"{%page%}","([0-9]+)")



	AddRule2=s & t

End Function


Function MakeIIS6Rewrite3()
	Dim s

	s=s & "#ISAPI Rewrite 3" & vbCrlf

	s=s & "RewriteBase "& CookiesPath() & vbCrlf

	If ZC_STATIC_MODE="REWRITE" Then

		s= AddRule2(s,"ZC_DEFAULT_REGEX",True)
		s= AddRule2(s,"ZC_CATEGORY_REGEX",True)
		s= AddRule2(s,"ZC_CATEGORY_REGEX",False)
		s= AddRule2(s,"ZC_USER_REGEX",True)
		s= AddRule2(s,"ZC_USER_REGEX",False)
		s= AddRule2(s,"ZC_TAGS_REGEX",True)
		s= AddRule2(s,"ZC_TAGS_REGEX",False)
		s= AddRule2(s,"ZC_DATE_REGEX",True)
		s= AddRule2(s,"ZC_DATE_REGEX",False)

	End If


	If ZC_POST_STATIC_MODE="REWRITE" Then

		s= AddRule2(s,"ZC_ARTICLE_REGEX",False)
		s= AddRule2(s,"ZC_PAGE_REGEX",False)

	End If


	MakeIIS6Rewrite3=s
End Function

Function AddRule3(s,regex,page)
	Dim t,r
	If regex="ZC_DEFAULT_REGEX" Then
		If InStr(ZC_DEFAULT_REGEX,"{%page%}")=0 Then
			r=Replace(ZC_DEFAULT_REGEX,".html","_{%page%}\.html")
		Else
			r=Replace(ZC_DEFAULT_REGEX,"default.html","")
			' r=ZC_DEFAULT_REGEX
		End If
		t=t & "     <rule name=""Imported Rule Default"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""catalog.asp?page={R:1}"" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If


	If regex="ZC_CATEGORY_REGEX" Then
		r=ZC_CATEGORY_REGEX
		Call FormatUrl(r,page)
		t=t & "     <rule name=""Imported Rule Category"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""catalog.asp?cate={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If

	If regex="ZC_USER_REGEX" Then
		r=ZC_USER_REGEX
		Call FormatUrl(r,page)
		t=t & "     <rule name=""Imported Rule Author"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""catalog.asp?auth={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If


	If regex="ZC_TAGS_REGEX" Then
		r=ZC_TAGS_REGEX
		Call FormatUrl(r,page)
		t=t & "     <rule name=""Imported Rule Tags"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""catalog.asp?tags={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If

	If regex="ZC_DATE_REGEX" Then
		r=ZC_DATE_REGEX
		Call FormatUrl(r,page)
		t=t & "     <rule name=""Imported Rule Date"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""catalog.asp?date={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If


	If regex="ZC_ARTICLE_REGEX" Then
		r=ZC_ARTICLE_REGEX
		r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
		r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
		t=t & "     <rule name=""Imported Rule Article"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""view.asp?id={R:1}"" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If


	If regex="ZC_PAGE_REGEX" Then
		r=ZC_PAGE_REGEX
		r=Replace(r,"{%name%}","{%alias%}")
		r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
		r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
		t=t & "     <rule name=""Imported Rule Page"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
		t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
		t=t & "      <action type=""Rewrite"" url=""view.asp?id={R:1}"" />" & vbCrlf
		t=t & "     </rule>" & vbCrlf
	End If


	t=Replace(t,"{%host%}/","")
	t=Replace(t,"{%name%}","(?!zb_)(.*)")
	t=Replace(t,"{%alias%}","(?!zb_)(.*)")
	t=Replace(t,"{%id%}","([0-9]+)")
	t=Replace(t,"{%date%}","([0-9\-]+)")
	t=Replace(t,"{%post%}",ZC_STATIC_DIRECTORY)
	t=Replace(t,"{%category%}","(?!zb_).*")
	t=Replace(t,"{%author%}","(?!zb_).*")
	t=Replace(t,"{%year%}","[0-9\-]+")
	t=Replace(t,"{%month%}","[0-9\-]+")
	t=Replace(t,"{%day%}","[0-9\-]+")
	t=Replace(t,"{%page%}","([0-9]+)")



	AddRule3=s & t

End Function

Function MakeIIS7UrlRewrite()
	Dim s

	s=s & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrlf
	s=s & "<configuration>" & vbCrlf
	s=s & " <system.webServer>" & vbCrlf
	s=s & "  <rewrite>" & vbCrlf
	s=s & "   <rules>" & vbCrlf

	If ZC_STATIC_MODE="REWRITE" Then

		s= AddRule3(s,"ZC_DEFAULT_REGEX",True)
		s= AddRule3(s,"ZC_CATEGORY_REGEX",True)
		s= AddRule3(s,"ZC_CATEGORY_REGEX",False)
		s= AddRule3(s,"ZC_USER_REGEX",True)
		s= AddRule3(s,"ZC_USER_REGEX",False)
		s= AddRule3(s,"ZC_TAGS_REGEX",True)
		s= AddRule3(s,"ZC_TAGS_REGEX",False)
		s= AddRule3(s,"ZC_DATE_REGEX",True)
		s= AddRule3(s,"ZC_DATE_REGEX",False)

	End If


	If ZC_POST_STATIC_MODE="REWRITE" Then

		s= AddRule3(s,"ZC_ARTICLE_REGEX",False)
		s= AddRule3(s,"ZC_PAGE_REGEX",False)

	End If

	s=s & "   </rules>" & vbCrlf
	s=s & "  </rewrite>" & vbCrlf
	s=s & " </system.webServer>" & vbCrlf
	s=s & "</configuration>" & vbCrlf


	MakeIIS7UrlRewrite=s
End Function

%>