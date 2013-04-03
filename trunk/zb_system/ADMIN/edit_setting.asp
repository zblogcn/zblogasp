<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 彩虹网志个人版
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    edit_setting.asp
'// 开始时间:    2005.03.16
'// 最后修改:    
'// 备    注:    编辑设置页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Edit_Setting_Begin in Action_Plugin_Edit_Setting_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Setting_Begin) Then Call Execute(sAction_Plugin_Edit_Setting_Begin)
Next

'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("SettingMng") Then Call ShowError(6)

Dim EditArticle

BlogTitle=ZC_MSG247

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
        <div id="divMain">
          <% Call GetBlogHint() %>
          <div class="divHeader"><%=BlogTitle%></div>
          <%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_SettingMng_SubMenu & "</div>"
%>
          <form method="post" action="../cmd.asp?act=SettingSav">
            <div id="divMain2">
              <div class="content-box"><!-- Start Content Box -->
                
                <div class="content-box-header">
                  <ul class="content-box-tabs">
                    <li><a href="#tab1" class="default-tab"><span><%=ZC_MSG105%></span></a></li>
                    <li><a href="#tab2"><span><%=ZC_MSG173%></span></a></li>
                    <li><a href="#tab3"><span><%=ZC_MSG186%></span></a></li>
                    <li><a href="#tab4"><span><%=ZC_MSG215%></span></a></li>
                    <li <%=IIF(ZC_POST_STATIC_MODE<>"STATIC","style='display:none;'","")%>><a href="#tab5"><span><%=ZC_MSG255%></span></a></li>
                  </ul>
                  <div class="clear"></div>
                </div>
                <!-- End .content-box-header -->
                
                <div class="content-box-content">
                  <%

	Function SplitNameAndNote(s)

		Dim i,j

		i=InStr(s,"(")
		j=InStr(s,")")

		If i>0 And j>0 Then 
			SplitNameAndNote="<p  align='left'><b>· " & Left(s,i-1) & "</b>"
			SplitNameAndNote=SplitNameAndNote & "<br/><span class='note'>&nbsp;&nbsp;" & Mid(s,i+1,Len(s)-i+1-2) & "</span></p>"
		Else
			SplitNameAndNote="<p  align='left'><b>· " & s & "</b></p>"
		End If
		
	End Function


	Dim i,j
	Dim tmpSng

	tmpSng=LoadFromFile(BlogPath & "zb_users/c_custom.asp","utf-8")

	Dim strZC_BLOG_HOST
	Dim strZC_BLOG_TITLE
	Dim strZC_BLOG_SUBTITLE
	Dim strZC_BLOG_NAME
	Dim strZC_BLOG_SUB_NAME
	Dim strZC_BLOG_COPYRIGHT
	Dim strZC_BLOG_MASTER
	Dim strZC_PERMANENT_DOMAIN_ENABLE
	

	strZC_BLOG_HOST=TransferHTML(ZC_BLOG_HOST,"[html-format]")
	strZC_BLOG_TITLE=TransferHTML(ZC_BLOG_TITLE,"[html-format]")
	strZC_BLOG_SUBTITLE=TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]")
	strZC_BLOG_NAME=TransferHTML(ZC_BLOG_NAME,"[html-format]")
	strZC_BLOG_SUB_NAME=TransferHTML(ZC_BLOG_SUB_NAME,"[html-format]")
	strZC_BLOG_COPYRIGHT=TransferHTML(ZC_BLOG_COPYRIGHT,"[html-format]")
	strZC_BLOG_MASTER=TransferHTML(ZC_BLOG_MASTER,"[html-format]")
	strZC_PERMANENT_DOMAIN_ENABLE=TransferHTML(ZC_PERMANENT_DOMAIN_ENABLE,"[html-format]")


	Response.Write "<div class=""tab-content default-tab"" style='border:none;padding:0px;margin:0;' id=""tab1"">"
	Response.Write "<input id=""edtZC_PERMANENT_DOMAIN_ENABLE"" name=""edtZC_PERMANENT_DOMAIN_ENABLE"" type=""hidden"" value=""" & strZC_PERMANENT_DOMAIN_ENABLE & """ />"
	Response.Write "<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>"
	Response.Write "<tr><td width='30%'>" & SplitNameAndNote(ZC_MSG126) & "</td><td><p><input id=""edtZC_BLOG_HOST"" name=""edtZC_BLOG_HOST"" style=""width:600px;"" type=""text"" "&IIF(CBool(ZC_PERMANENT_DOMAIN_ENABLE),"","readonly=""readonly""")&" value=""" & strZC_BLOG_HOST & """ /><br/><label><input type='radio' name='ZC_PERMANENT_DOMAIN_ENABLE' "&IIF(CBool(ZC_PERMANENT_DOMAIN_ENABLE)=True,"","checked=""checked""")&" value='False' onchange=""$('#edtZC_PERMANENT_DOMAIN_ENABLE').val('False');$('#edtZC_BLOG_HOST').prop('readonly', true);"" />"&ZC_MSG297&"</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<label><input type='radio' name='ZC_PERMANENT_DOMAIN_ENABLE' "&IIF(CBool(ZC_PERMANENT_DOMAIN_ENABLE)=False,"","checked=""checked""")&" value='True'   onchange=""$('#edtZC_PERMANENT_DOMAIN_ENABLE').val('True');$('#edtZC_BLOG_HOST').prop('readonly', false);"" />"&ZC_MSG296&"</label></p></td></tr>"
	'Response.Write "<tr><td width='30%'>" & SplitNameAndNote(ZC_MSG091) & "</td><td><p><input id=""edtZC_BLOG_NAME"" name=""edtZC_BLOG_NAME"" style=""width:600px;"" type=""text"" value=""" & strZC_BLOG_NAME & """ /></p></td></tr>"
	'Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG092) & "</td><td><p><input id=""edtZC_BLOG_SUB_NAME"" name=""edtZC_BLOG_SUB_NAME"" style=""width:600px;""  type=""text"" value=""" & strZC_BLOG_SUB_NAME & """ /></p></td></tr>"
	Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG093) & "</td><td><p><input id=""edtZC_BLOG_TITLE"" name=""edtZC_BLOG_TITLE"" style=""width:600px;""  type=""text"" value=""" & strZC_BLOG_TITLE &""" /></p></td></tr>"
	Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG094) & "</td><td><p><input id=""edtZC_BLOG_SUBTITLE"" name=""edtZC_BLOG_SUBTITLE"" style=""width:600px;""  type=""text"" value=""" & strZC_BLOG_SUBTITLE & """ /></p></td></tr>"
	Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG096) & "</td><td><p><textarea cols=""3"" rows=""6"" id=""edtZC_BLOG_COPYRIGHT"" name=""edtZC_BLOG_COPYRIGHT"" style=""width:600px;"">" & strZC_BLOG_COPYRIGHT & "</textarea></p></td></tr>"
	Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG300) & "</td><td><p><select id=""edtZC_BLOG_LANGUAGEPACK"" name=""edtZC_BLOG_LANGUAGEPACK"">"
	Dim obj,strl
	For Each obj In CreateObject("scripting.filesystemobject").GetFolder(BlogPath&"\zb_users\language").Files
		strl=LoadFromFile(obj.Path,"utf-8")
		strl=Split(strl,"</language>")
		If Ubound(strl)>0 Then 	
			Response.Write "<option value="""&Split(obj.Name,".")(0)&""" "&IIf(ZC_BLOG_LANGUAGEPACK&".asp"=obj.Name,"selected","")&">"
			Response.Write TransferHTML(Split(strl(0),"<language>")(1) & "(" & obj.Name & ")","[html-format]") &"</option>"
		End If
	Next
	Response.Write "</select></p></td></tr>"


	Response.Write "</table>"
	Response.Write "</div>"



	Response.Write "<div class=""tab-content"" style='border:none;padding:0px;margin:0;' id=""tab2"">"
	Response.Write "<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>"
	tmpSng=LoadFromFile(BlogPath & "zb_users/c_option.asp","utf-8")



		'ZC_BLOG_CLSID=TransferHTML(ZC_BLOG_CLSID,"[html-format]")
		'Response.Write "<tr><td width='30%'>" & SplitNameAndNote(ZC_MSG174) & "</td><td><p><input id=""edtZC_BLOG_CLSID"" name=""edtZC_BLOG_CLSID"" style=""width:600px;"" type=""text"" value=""" & ZC_BLOG_CLSID & """ /></p></td></tr>"


		ZC_TIME_ZONE=TransferHTML(ZC_TIME_ZONE,"[html-format]")
		Response.Write "<tr><td width='30%'>" & SplitNameAndNote(ZC_MSG175) & "</td><td><p><input id=""edtZC_TIME_ZONE"" name=""edtZC_TIME_ZONE"" style=""width:600px;"" type=""text"" value=""" & ZC_TIME_ZONE & """ /></p></td></tr>"




		ZC_HOST_TIME_ZONE=TransferHTML(ZC_HOST_TIME_ZONE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG206) & "</td><td><p><input id=""edtZC_HOST_TIME_ZONE"" name=""edtZC_HOST_TIME_ZONE"" style=""width:600px;"" type=""text"" value=""" & ZC_HOST_TIME_ZONE & """ /></p></td></tr>"


		ZC_BLOG_LANGUAGE=TransferHTML(ZC_BLOG_LANGUAGE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG176) & "</td><td><p><input id=""edtZC_BLOG_LANGUAGE"" name=""edtZC_BLOG_LANGUAGE"" style=""width:600px;"" type=""text"" value=""" & ZC_BLOG_LANGUAGE & """ /></p></td></tr>"



		ZC_UPLOAD_FILETYPE=TransferHTML(ZC_UPLOAD_FILETYPE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG183) & "</td><td><p><input id=""edtZC_UPLOAD_FILETYPE"" name=""edtZC_UPLOAD_FILETYPE"" style=""width:600px;"" type=""text"" value=""" & ZC_UPLOAD_FILETYPE & """ /></p></td></tr>"


		ZC_UPLOAD_FILESIZE=TransferHTML(ZC_UPLOAD_FILESIZE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG184) & "</td><td><p><input id=""edtZC_UPLOAD_FILESIZE"" name=""edtZC_UPLOAD_FILESIZE"" style=""width:600px;"" type=""text"" value=""" & ZC_UPLOAD_FILESIZE & """ /></p></td></tr>"


		ZC_RSS_EXPORT_WHOLE=TransferHTML(ZC_RSS_EXPORT_WHOLE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG226) & "</td><td><p><input id=""edtZC_RSS_EXPORT_WHOLE"" name=""edtZC_RSS_EXPORT_WHOLE"" style="""" type=""text"" value=""" & ZC_RSS_EXPORT_WHOLE & """ class=""checkbox""/></p></td></tr>"



	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tab-content"" style='border:none;padding:0px;margin:0;' id=""tab3"">"
	Response.Write "<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>"


		ZC_DISPLAY_COUNT=TransferHTML(ZC_DISPLAY_COUNT,"[html-format]")
		Response.Write "<tr><td width='30%'>" & SplitNameAndNote(ZC_MSG190) & "</td><td><p><input id=""edtZC_DISPLAY_COUNT"" name=""edtZC_DISPLAY_COUNT"" style=""width:600px;"" type=""text"" value=""" & ZC_DISPLAY_COUNT & """ /></p></td></tr>"


		ZC_PAGEBAR_COUNT=TransferHTML(ZC_PAGEBAR_COUNT,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG132) & "</td><td><p><input id=""edtZC_PAGEBAR_COUNT"" name=""edtZC_PAGEBAR_COUNT"" style=""width:600px;"" type=""text"" value=""" & ZC_PAGEBAR_COUNT & """ /></p></td></tr>"


		ZC_SEARCH_COUNT=TransferHTML(ZC_SEARCH_COUNT,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG274) & "</td><td><p><input id=""edtZC_SEARCH_COUNT"" name=""edtZC_SEARCH_COUNT"" style=""width:600px;"" type=""text"" value=""" & ZC_SEARCH_COUNT & """ /></p></td></tr>"


		ZC_USE_NAVIGATE_ARTICLE=TransferHTML(ZC_USE_NAVIGATE_ARTICLE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG209) & "</td><td><p><input id=""edtZC_USE_NAVIGATE_ARTICLE"" name=""edtZC_USE_NAVIGATE_ARTICLE"" style="""" type=""text"" value=""" & ZC_USE_NAVIGATE_ARTICLE & """ class=""checkbox""/></p></td></tr>"


		ZC_MUTUALITY_COUNT=TransferHTML(ZC_MUTUALITY_COUNT,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG230) & "</td><td><p><input id=""edtZC_MUTUALITY_COUNT"" name=""edtZC_MUTUALITY_COUNT"" style=""width:600px;"" type=""text"" value=""" & ZC_MUTUALITY_COUNT & """ /></p></td></tr>"

		ZC_SYNTAXHIGHLIGHTER_ENABLE=TransferHTML(ZC_SYNTAXHIGHLIGHTER_ENABLE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG301) & "</td><td><p><input id=""edtZC_USE_NAVIGATE_ARTICLE"" name=""edtZC_SYNTAXHIGHLIGHTER_ENABLE"" style="""" type=""text"" value=""" & ZC_SYNTAXHIGHLIGHTER_ENABLE & """ class=""checkbox""/></p></td></tr>"


	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tab-content"" style='border:none;padding:0px;margin:0;' id=""tab4"">"
	Response.Write "<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>"


		ZC_COMMENT_TURNOFF=TransferHTML(ZC_COMMENT_TURNOFF,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG262) & "</td><td><p><input id=""edtZC_COMMENT_TURNOFF"" name=""edtZC_COMMENT_TURNOFF"" style="""" type=""text"" value=""" & ZC_COMMENT_TURNOFF & """ class=""checkbox""/></p></td></tr>"


		ZC_COMMENTS_DISPLAY_COUNT=TransferHTML(ZC_COMMENTS_DISPLAY_COUNT,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG171) & "</td><td><p><input id=""edtZC_COMMENTS_DISPLAY_COUNT"" name=""edtZC_COMMENTS_DISPLAY_COUNT"" style=""width:600px;"" type=""text"" value=""" & ZC_COMMENTS_DISPLAY_COUNT & """/></p></td></tr>"


		ZC_COMMENT_REVERSE_ORDER_EXPORT=TransferHTML(ZC_COMMENT_REVERSE_ORDER_EXPORT,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG271) & "</td><td><p><input id=""edtZC_COMMENT_REVERSE_ORDER_EXPORT"" name=""edtZC_COMMENT_REVERSE_ORDER_EXPORT"" style="""" type=""text"" value=""" & ZC_COMMENT_REVERSE_ORDER_EXPORT & """ class=""checkbox""/></p></td></tr>"


		ZC_COMMENT_VERIFY_ENABLE=TransferHTML(ZC_COMMENT_VERIFY_ENABLE,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG185) & "</td><td><p><input id=""edtZC_COMMENT_VERIFY_ENABLE"" name=""edtZC_COMMENT_VERIFY_ENABLE"" style="""" type=""text"" value=""" & ZC_COMMENT_VERIFY_ENABLE & """ class=""checkbox""/></p></td></tr>"


	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tab-content"" style='border:none;padding:0px;margin:0;' id=""tab5"">"
	Response.Write "<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>"



		ZC_STATIC_DIRECTORY=TransferHTML(ZC_STATIC_DIRECTORY,"[html-format]")
		Response.Write "<tr><td width='30%'>" & SplitNameAndNote(ZC_MSG178) & "</td><td><p><input id=""edtZC_STATIC_DIRECTORY"" name=""edtZC_STATIC_DIRECTORY"" style=""width:600px;"" type=""text"" value=""" & ZC_STATIC_DIRECTORY & """ /></p></td></tr>"


		ZC_REBUILD_FILE_COUNT=TransferHTML(ZC_REBUILD_FILE_COUNT,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG181) & "</td><td><p><input id=""edtZC_REBUILD_FILE_COUNT"" name=""edtZC_REBUILD_FILE_COUNT"" style=""width:600px;"" type=""text"" value=""" & ZC_REBUILD_FILE_COUNT & """ /></p></td></tr>"



		ZC_REBUILD_FILE_INTERVAL=TransferHTML(ZC_REBUILD_FILE_INTERVAL,"[html-format]")
		Response.Write "<tr><td>" & SplitNameAndNote(ZC_MSG182) & "</td><td><p><input id=""edtZC_REBUILD_FILE_INTERVAL"" name=""edtZC_REBUILD_FILE_INTERVAL"" style=""width:600px;"" type=""text"" value=""" & ZC_REBUILD_FILE_INTERVAL & """ /></p></td></tr>"



	Response.Write "</table>"
	Response.Write "</div>"

%>
                </div>
                <!-- End .content-box-content --> 
                
              </div>
              <!-- End .content-box -->
              <%



	Response.Write "<p><br/><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='' /></p>"

%>
            </div>
          </form>
        </div>
        <script type="text/javascript">

ActiveTopMenu('topmenu2');

</script> 
        <!--#include file="admin_footer.asp"-->
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>