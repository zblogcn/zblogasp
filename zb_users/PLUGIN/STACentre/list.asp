<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)
BlogTitle="静态管理中心"

Function AddRule1(s,regex,page)
	Dim t,r
	If regex="ZC_DEFAULT_REGEX" Then
	r=Replace(ZC_DEFAULT_REGEX,".html","{%page%}\.html")
t=t & "RewriteRule "& r &" /catalog\.asp\?page=$1" & vbCrlf
	End If


	If regex="ZC_CATEGORY_REGEX" Then
	r=ZC_CATEGORY_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule "& r &" /catalog\.asp\?cate=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If

	If regex="ZC_USER_REGEX" Then
	r=ZC_USER_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule "& r &" /catalog\.asp\?auth=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If


	If regex="ZC_TAGS_REGEX" Then
	r=ZC_TAGS_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule "& r &" /catalog\.asp\?tags=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If

	If regex="ZC_DATE_REGEX" Then
	r=ZC_DATE_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule "& r &" /catalog\.asp\?date=$1"&IIF(page=True,"&page=$2","") & vbCrlf
	End If


	If regex="ZC_ARTICLE_REGEX" Then
	r=ZC_ARTICLE_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule "& r &" /view\.asp\?id=$1" & vbCrlf
	End If


	If regex="ZC_PAGE_REGEX" Then
	r=ZC_PAGE_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule "& r &" /view\.asp\?id=$1" & vbCrlf
	End If


	t=Replace(t,"{%host%}/",CookiesPath())
	t=Replace(t,"{%name%}","(.*)")
	t=Replace(t,"{%alias%}","(.*)")
	t=Replace(t,"{%id%}","([0-9]+)")
	t=Replace(t,"{%date%}","([0-9\-]+)")
	t=Replace(t,"{%post%}",ZC_STATIC_DIRECTORY)
	t=Replace(t,"{%category%}",".*")
	t=Replace(t,"{%author%}",".*")
	t=Replace(t,"{%year%}","[0-9\-]+")
	t=Replace(t,"{%month%}","[0-9\-]+")
	t=Replace(t,"{%day%}","[0-9\-]+")
	t=Replace(t,"{%page%}","_([0-9]+)")



	AddRule1=s & t

End Function

Function MakeIIS6Rewrite2()
	Dim s

s=s & "[ISAPI_Rewrite]" & vbCrlf

s=s & "" & vbCrlf

If ZC_STATIC_MODE="REWRITE" Then

s= AddRule1(s,"ZC_DEFAULT_REGEX",True)
s= AddRule1(s,"ZC_CATEGORY_REGEX",True)
s= AddRule1(s,"ZC_CATEGORY_REGEX",False)
s= AddRule1(s,"ZC_USER_REGEX",True)
s= AddRule1(s,"ZC_USER_REGEX",False)
s= AddRule1(s,"ZC_TAGS_REGEX",True)
s= AddRule1(s,"ZC_TAGS_REGEX",False)
s= AddRule1(s,"ZC_DATE_REGEX",True)
s= AddRule1(s,"ZC_DATE_REGEX",False)

End If


If ZC_POST_STATIC_MODE="REWRITE" Then

s= AddRule1(s,"ZC_ARTICLE_REGEX",False)
s= AddRule1(s,"ZC_PAGE_REGEX",False)

End If


	MakeIIS6Rewrite2=TransferHTML(s,"[html-format]")
End Function

Function AddRule2(s,regex,page)
	Dim t,r
	If regex="ZC_DEFAULT_REGEX" Then
	r=Replace(ZC_DEFAULT_REGEX,".html","{%page%}\.html")
t=t & "RewriteRule ^"& r &"$ /catalog.asp\?page=$1" & vbCrlf
	End If


	If regex="ZC_CATEGORY_REGEX" Then
	r=ZC_CATEGORY_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule ^"& r &"$ /catalog.asp\?cate=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If

	If regex="ZC_USER_REGEX" Then
	r=ZC_USER_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule ^"& r &"$ /catalog.asp\?auth=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If


	If regex="ZC_TAGS_REGEX" Then
	r=ZC_TAGS_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule ^"& r &"$ /catalog.asp\?tags=$1"&IIF(page=True,"&page=$2","")&" [NU]" & vbCrlf
	End If

	If regex="ZC_DATE_REGEX" Then
	r=ZC_DATE_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
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
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "RewriteRule ^"& r &"$ /view.asp\?id=$1 [NU]" & vbCrlf
	End If


	t=Replace(t,"{%host%}/","")
	t=Replace(t,"{%name%}","(.*)")
	t=Replace(t,"{%alias%}","(.*)")
	t=Replace(t,"{%id%}","([0-9]+)")
	t=Replace(t,"{%date%}","([0-9\-]+)")
	t=Replace(t,"{%post%}",ZC_STATIC_DIRECTORY)
	t=Replace(t,"{%category%}",".*")
	t=Replace(t,"{%author%}",".*")
	t=Replace(t,"{%year%}","[0-9\-]+")
	t=Replace(t,"{%month%}","[0-9\-]+")
	t=Replace(t,"{%day%}","[0-9\-]+")
	t=Replace(t,"{%page%}","_([0-9]+)")



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


	MakeIIS6Rewrite3=TransferHTML(s,"[html-format]")
End Function


Function AddRule3(s,regex,page)
	Dim t,r
	If regex="ZC_DEFAULT_REGEX" Then
	r=Replace(ZC_DEFAULT_REGEX,".html","{%page%}\.html")
t=t & "     <rule name=""Imported Rule Default"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
t=t & "      <action type=""Rewrite"" url=""catalog.asp?page={R:1}"" />" & vbCrlf
t=t & "     </rule>" & vbCrlf
	End If


	If regex="ZC_CATEGORY_REGEX" Then
	r=ZC_CATEGORY_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "     <rule name=""Imported Rule Category"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
t=t & "      <action type=""Rewrite"" url=""catalog.asp?cate={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
t=t & "     </rule>" & vbCrlf
	End If

	If regex="ZC_USER_REGEX" Then
	r=ZC_USER_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "     <rule name=""Imported Rule Author"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
t=t & "      <action type=""Rewrite"" url=""catalog.asp?auth={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
t=t & "     </rule>" & vbCrlf
	End If


	If regex="ZC_TAGS_REGEX" Then
	r=ZC_TAGS_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "     <rule name=""Imported Rule Tags"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
t=t & "      <action type=""Rewrite"" url=""catalog.asp?auth={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
t=t & "     </rule>" & vbCrlf
	End If

	If regex="ZC_DATE_REGEX" Then
	r=ZC_DATE_REGEX
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "     <rule name=""Imported Rule Date"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
t=t & "      <action type=""Rewrite"" url=""catalog.asp?auth={R:1}"&IIF(page=True,"&amp;page={R:2}","")&""" />" & vbCrlf
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
	r=Replace(r,"/default.html",IIF(page=True,"{%page%}","")&"/")
	r=Replace(r,".html",IIF(page=True,"{%page%}","")&".html")
t=t & "     <rule name=""Imported Rule Page"&IIF(page=True,"+Page","")&""" stopProcessing=""true"">" & vbCrlf
t=t & "      <match url=""^"& r &"$"" ignoreCase=""false"" />" & vbCrlf
t=t & "      <action type=""Rewrite"" url=""view.asp?id={R:1}"" />" & vbCrlf
t=t & "     </rule>" & vbCrlf
	End If


	t=Replace(t,"{%host%}/","")
	t=Replace(t,"{%name%}","(.*)")
	t=Replace(t,"{%alias%}","(.*)")
	t=Replace(t,"{%id%}","([0-9]+)")
	t=Replace(t,"{%date%}","([0-9\-]+)")
	t=Replace(t,"{%post%}",ZC_STATIC_DIRECTORY)
	t=Replace(t,"{%category%}",".*")
	t=Replace(t,"{%author%}",".*")
	t=Replace(t,"{%year%}","[0-9\-]+")
	t=Replace(t,"{%month%}","[0-9\-]+")
	t=Replace(t,"{%day%}","[0-9\-]+")
	t=Replace(t,"{%page%}","_([0-9]+)")



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


	MakeIIS7UrlRewrite=TransferHTML(s,"[html-format]")
End Function

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
pre{
	border:1px solid #ededed;
	margin:0px;
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left">配置页面</span></a><a href="list.asp"><span class="m-left m-now">ReWrite规则</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
<%If ZC_POST_STATIC_MODE="REWRITE" Or ZC_STATIC_MODE="REWRITE" Then%>
			<div class="content-box"><!-- Start Content Box -->
				
				<div class="content-box-header">
			
					<ul class="content-box-tabs">

	<li><a href="#tab1" class="default-tab"><span>IIS6+ISAPI Rewrite 2.X</span></a></li>
	<li><a href="#tab2"><span>IIS6+ISAPI Rewrite 3.X</span></a></li>
	<li><a href="#tab3"><span>IIS7,7.5+Url Rewrite</span></a></li>
					</ul>
					
					<div class="clear"></div>
					
				</div> <!-- End .content-box-header -->

				<div class="content-box-content">
<div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab1">
<pre>
<%=MakeIIS6Rewrite2()%>
</pre>
<hr/>
<p><span class="star">请在网站根目录创建httpd.ini文件并把相关内容复制进去.</span></p>
</div>


<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
<pre>
<%=MakeIIS6Rewrite3()%>
</pre>
<hr/>
<p><span class="star">请在网站根目录创建.htaccess文件并把相关内容复制进去.</span></p>
</div>

<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
<pre>
<%=MakeIIS7UrlRewrite()%>
</pre>
<hr/>
<p><span class="star">请在网站<u>"当前目录"</u>创建web.config文件并把相关内容复制进去.</span></p>
</div>

				</div> <!-- End .content-box-content -->
				
			</div> <!-- End .content-box -->
<%Else%>
<hr/>
<p><b>文章及页面和分类页都没有启用动态模式+Rewrite支持,所以没有可用规则.</b></p>
<%End If%>
<p><br/></p>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
