<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("Totoro")=False Then Call ShowError(48)

BlogTitle="TotoroⅢ（基于TotoroⅡ的Z-Blog的评论管理审核系统增强版）"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint"><%=GetBlogHint%></div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><a href="setting.asp"><span class="m-left m-now">TotoroⅢ设置</span></a><a href="regexptest.asp"><span class="m-right">黑词测试</span></a><a href="onlinetest.asp"><span class="m-right">模拟测试</span></a></div>
          <div id="divMain2">
            <form id="edit" name="edit" method="post">
              <div class="content-box">
              <!-- Start Content Box -->
              
              <div class="content-box-header">
                <ul class="content-box-tabs">
                  <li><a href="#tab1" class="default-tab"><span>加分减分细则设置</span></a></li>
                  <li><a href="#tab2"><span>过滤列表设置</span></a></li>
                  <li><a href="#tab4"><span>过滤设置</span></a></li>
                  <li><a href="#tab5"><span>提示语设置</span></a></li>
                  <li><a href="#tab6"><span>其他设置</span></a></li>
                  <li><a href="#tab3"><span>关于TotoroⅢ</span></a></li>
                </ul>
                <div class="clear"></div>
              </div>
              <!-- End .content-box-header -->
              
              <div class="content-box-content" id="totorobox">
              <div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab1">
              <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
              <tr height="40">
                <td width='50'>序号</td>
                <td width='200' align="center">规则</td>
                <td width="80" align="center">分数</td>
                <td align="center">说明</td>
              </tr>
              <%

	Call Totoro_Initialize

	
	Dim strZC_TOTORO_HYPERLINK_VALUE
	strZC_TOTORO_HYPERLINK_VALUE=Totoro_Config.Read("TOTORO_HYPERLINK_VALUE")
	strZC_TOTORO_HYPERLINK_VALUE=TransferHTML(strZC_TOTORO_HYPERLINK_VALUE,"[html-format]")
	Response.Write "<tr><td>1</td><td>链接评分</td><td><input name=""strZC_TOTORO_HYPERLINK_VALUE""type=""text"" value=""" & strZC_TOTORO_HYPERLINK_VALUE & """/></td><td>(默认：10)，每多一个链接SV翻倍加分</td></tr>"
	
	
	Dim strTOTORO_INTERVAL_VALUE
	strTOTORO_INTERVAL_VALUE=Totoro_Config.Read("TOTORO_INTERVAL_VALUE")
	strTOTORO_INTERVAL_VALUE=TransferHTML(strTOTORO_INTERVAL_VALUE,"[html-format]")
	Response.Write "<tr><td>2</td><td>提交频率评分<td><input name=""strZC_TOTORO_INTERVAL_VALUE""type=""text"" value=""" & strTOTORO_INTERVAL_VALUE & """/></td><Td>(默认：25)，根据1小时内同一IP的评论数量加分。(1小时内10条评论加SV的1/5，以此类推，最多加SV的6/5)</td></tr>"
	
	Dim strTOTORO_BADWORD_VALUE
	strTOTORO_BADWORD_VALUE=Totoro_Config.Read("TOTORO_BADWORD_VALUE")
	strTOTORO_BADWORD_VALUE=TransferHTML(strTOTORO_BADWORD_VALUE,"[html-format]")
	Response.Write "<tr><td>3</td><td>评论里的每一个黑词都加</td><td><input name=""strZC_TOTORO_BADWORD_VALUE""type=""text"" value=""" & strTOTORO_BADWORD_VALUE & """/></td><td>(默认：50)</td></tr>"
	
	Dim strTOTORO_LEVEL_VALUE
	strTOTORO_LEVEL_VALUE=Totoro_Config.Read("TOTORO_LEVEL_VALUE")
	strTOTORO_LEVEL_VALUE=TransferHTML(strTOTORO_LEVEL_VALUE,"[html-format]")
	Response.Write "<tr><td>4</td><td>用户信任度评分</td><td><input name=""strZC_TOTORO_LEVEL_VALUE""type=""text"" value=""" & strTOTORO_LEVEL_VALUE & """/></td><td>(默认：100)，初级用户评论时SV减基数×1，中级用户SV减基数×2，高级用户减SV减基数×3，管理员SV减基数×4</td></tr>"

	Dim strTOTORO_NAME_VALUE
	strTOTORO_NAME_VALUE=Totoro_Config.Read("TOTORO_NAME_VALUE")
	strTOTORO_NAME_VALUE=TransferHTML(strTOTORO_NAME_VALUE,"[html-format]")
	Response.Write "<tr><td>5</td><td>访客熟悉度评分</td><td><input name=""strZC_TOTORO_NAME_VALUE""type=""text"" value=""" & strTOTORO_NAME_VALUE & """/></td><td>(默认：45)，同一访客在BLOG留言1-10条内的SV减10分,10-20条的SV减10分再减基数×1，20-50条的SV减10分再减基数×2，大于50条的SV减10分再减基数×3</td></tr>"
	
	Dim strTOTORO_NUMBER_VALUE
	strTOTORO_NUMBER_VALUE=Totoro_Config.Read("TOTORO_NUMBER_VALUE")
	strTOTORO_NUMBER_VALUE=TransferHTML(strTOTORO_NUMBER_VALUE,"[html-format]")
	Response.Write "<tr><td>6</td><td>数字长度评分</td><td><input name=""strTOTORO_NUMBER_VALUE""type=""text"" value=""" & strTOTORO_NUMBER_VALUE & """/></td><td>(默认：10)。若数字长度达到10位，自动加上基数。多几位，加几次基数。</td></tr>"

	Dim strZC_TOTORO_REPLACE_KEYWORD
		strZC_TOTORO_REPLACE_KEYWORD=Totoro_Config.Read("TOTORO_REPLACE_KEYWORD")
		strZC_TOTORO_REPLACE_KEYWORD=TransferHTML(strZC_TOTORO_REPLACE_KEYWORD,"[html-format]")
	Response.Write "<tr><td>7</td><td>自动把敏感词替换为</td><td><input type=""text"" value="""& strZC_TOTORO_REPLACE_KEYWORD &""" name=""strZC_TOTORO_REPLACE_KEYWORD""/></td><td></td></tr>"

	Dim strZC_TOTORO_CHINESESV
		strZC_TOTORO_CHINESESV=Totoro_Config.Read("TOTORO_CHINESESV")
		strZC_TOTORO_CHINESESV=TransferHTML(strZC_TOTORO_CHINESESV,"[html-format]")
	Response.Write "<tr><td>8</td><td>一旦评论内没有汉字自动加SV</td><td><input type=""text"" value="""& strZC_TOTORO_CHINESESV &""" name=""strZC_TOTORO_CHINESESV""/></td><td>(默认：150)</td></tr>"

	Response.Write "<tr>"
  	Response.Write "<td>9</td><td>设置系统审核阙值</td>"
	Dim strZC_TOTORO_SV_THRESHOLD
	strZC_TOTORO_SV_THRESHOLD=Totoro_Config.Read("TOTORO_SV_THRESHOLD")
	strZC_TOTORO_SV_THRESHOLD=TransferHTML(strZC_TOTORO_SV_THRESHOLD,"[html-format]")
	Response.Write "<td><input name=""strZC_TOTORO_SV_THRESHOLD"" type=""text"" value=""" & strZC_TOTORO_SV_THRESHOLD & """/></td><td>(默认：50)，阙值越小越严格，低于0则使游客的评论全进入审核</td></tr>"
	
  	Response.Write "<tr><td>10</td><td>设置自动删除阙值</td><td>"
	Dim strZC_TOTORO_SV_THRESHOLD2
	strZC_TOTORO_SV_THRESHOLD2=Totoro_Config.Read("TOTORO_SV_THRESHOLD2")
	strZC_TOTORO_SV_THRESHOLD2=TransferHTML(strZC_TOTORO_SV_THRESHOLD2,"[html-format]")
	Response.Write "<input name=""strZC_TOTORO_SV_THRESHOLD2"" type=""text"" value=""" & strZC_TOTORO_SV_THRESHOLD2 & """/></td><td>(默认：150)，阙值达到该值并且阙值达到系统审核阙值则不审核直接删除。为0则不删除)</td></tr>"

  	Response.Write "<tr><td>11</td><td>设置IP回溯值</td><td>"
	Dim strZC_TOTORO_KILLIP
	strZC_TOTORO_KILLIP=Totoro_Config.Read("TOTORO_KILLIP")
	strZC_TOTORO_KILLIP=TransferHTML(strZC_TOTORO_KILLIP,"[html-format]")
	Response.Write "<input name=""strZC_TOTORO_KILLIP"" type=""text"" value=""" & strZC_TOTORO_KILLIP & """/></td><td>(默认：3) 一旦某个IP一天内被拦截的评论超过设定的值，则将该IP一天内的评论全部进入审核。若该IP有一条评论直接被拦截，所有评论也将进入审核状态。</td></tr></table></div>"

%>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
                <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
                  <tr>
                    <td height="40">过滤IP(分隔符'|')</td>
                  </tr>
                  <tr>
                    <td><%
                	Dim strTOTORO_FILTERIP
	strTOTORO_FILTERIP=Totoro_Config.Read("TOTORO_FILTERIP")
		strTOTORO_FILTERIP=TransferHTML(strTOTORO_FILTERIP,"[html-format]")
		Response.Write "<textarea rows=""6"" name=""strTOTORO_FILTERIP"" style=""width:99%"" >"& strTOTORO_FILTERIP &"</textarea>"
%></td>
                  </tr>
                  <tr>
                    <td height="40">黑词列表(请使用正则,最后一个字符不能是|):</td>
                  </tr>
                  <tr>
                    <td><%

	Dim strZC_TOTORO_BADWORD_LIST
	strZC_TOTORO_BADWORD_LIST=Totoro_Config.Read("TOTORO_BADWORD_LIST")
	strZC_TOTORO_BADWORD_LIST=TransferHTML(strZC_TOTORO_BADWORD_LIST,"[html-format]")
	If Left(strZC_TOTORO_BADWORD_LIST,1)<>"%" Then strZC_TOTORO_BADWORD_LIST=vbsescape(strZC_TOTORO_BADWORD_LIST)
		Response.Write "<textarea style=""display:none"" name=""strZC_TOTORO_BADWORD_LIST"" id=""escape_badword"">"&strZC_TOTORO_BADWORD_LIST&"</textarea>"
		Response.Write "<textarea rows=""6""  id=""unescape_badword"" style=""width:99%"" >请稍候，正在为您解码..</textarea>"
%></td>
                  </tr>
                  <tr>
                    <td height="40">敏感词列表(请使用正则,最后一个字符不能是|):</td>
                  </tr>
                  <tr>
                    <td><%	

	Dim strZC_TOTORO_REPLACE_LIST
	strZC_TOTORO_REPLACE_LIST=Totoro_Config.Read("TOTORO_REPLACE_LIST")
	strZC_TOTORO_REPLACE_LIST=TransferHTML(strZC_TOTORO_REPLACE_LIST,"[html-format]")
	If Left(strZC_TOTORO_REPLACE_LIST,1)<>"%" Then strZC_TOTORO_REPLACE_LIST=vbsescape(strZC_TOTORO_REPLACE_LIST)
		Response.Write "<textarea style=""display:none"" name=""strZC_TOTORO_REPLACE_LIST""  id=""escape_replace"" style=""display:none"">"&strZC_TOTORO_REPLACE_LIST&"</textarea>"
		Response.Write "<textarea rows=""6"" id=""unescape_replace"" style=""width:99%"" >请稍候，正在为您解码..</textarea>"

%></td>
                  </tr>
                </table>
              </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab4">
                <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
                  <tr height="40">
                    <td width='140'>配置项</td>
                    <td width='50'>开关</td>
                    <td>详细说明</td>
                  </tr>
                  <tr height='32'>
                    <td>自动转换火星文</td>
                    <td><%	Response.Write "<input name=""bolTOTORO_ConHuoxingwen"" id=""bolTOTORO_ConHuoxingwen"" type=""text"" class=""checkbox"" value="""&CStr(Totoro_Config.Read("TOTORO_ConHuoxingwen"))&""""%>
                      ></td>
                    <td>将把希腊文俄文字母、罗马数字、列表符、全角字符、汉语拼音、菊花文、西欧字符转换为半角英文字母、半角数字、半角符号再进行反spam测试，不影响实际显示的评论</td>
                  </tr>
                  <tr height='32'>
                    <td>简繁转换</td>
                    <td><%	Response.Write "<input name=""bolTOTORO_TRANTOSIMP"" id=""bolTOTORO_TRANTOSIMP"" type=""text"" class=""checkbox"" value="""&CStr(Totoro_Config.Read("TOTORO_TRANTOSIMP"))&""""%>
                      ></td>
                    <td>将把繁体字转换为简化字再进行反spam测试，不影响实际显示的评论</td>
                  </tr>
                  <tr height='32'>
                    <td>后台审核</td>
                    <td><%	Response.Write "<input name=""bolTOTORO_DEL_DIRECTLY"" id=""bolTOTORO_DEL_DIRECTLY"" type=""text"" class=""checkbox"" value="""&CStr(Totoro_Config.Read("TOTORO_DEL_DIRECTLY"))&""""%>
                      ></td>
                    <td>点击[<img src="<%=BlogHost%>zb_system/image/admin/minus-shield.png" alt="加入审核"/>]提取域名后直接删除评论（若不删除则进入审核）</td>
                  </tr>
                  <tr height='32'>
                    <td>标点过滤</td>
                    <td><%	Response.Write "<input name=""bolTOTORO_PM"" id=""bolTOTORO_PM"" type=""text"" class=""checkbox"" value="""&CStr(Totoro_Config.Read("TOTORO_PM"))&""""%>
                      ></td>
                    <td>把大部分标点和HTML代码过滤再进行反spam测试，不影响实际显示的评论</td>
                  </tr>
                </table>
              </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab5">
                <table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
                  <tr>
                    <td height="40">评论被过滤时的提示</td>
                  </tr>
                  <tr>
                    <td><%
                	Dim strTOTORO_CHECKSTR
	strTOTORO_CHECKSTR=Totoro_Config.Read("TOTORO_CHECKSTR")
		strTOTORO_CHECKSTR=TransferHTML(strTOTORO_CHECKSTR,"[html-format]")
		Response.Write "<textarea rows=""6"" name=""strTOTORO_CHECKSTR"" style=""width:99%"" >"& strTOTORO_CHECKSTR &"</textarea>"
%></td>
                  </tr>
                  <tr>
                    <td height="40">评论被拦截时的提示</td>
                  </tr>
                  <tr>
                    <td><%

	Dim strTOTORO_THROWSTR
	strTOTORO_THROWSTR=Totoro_Config.Read("TOTORO_THROWSTR")
		strTOTORO_THROWSTR=TransferHTML(strTOTORO_THROWSTR,"[html-format]")
		Response.Write "<textarea rows=""6"" name=""strTOTORO_THROWSTR"" style=""width:99%"" >"& strTOTORO_THROWSTR &"</textarea>"
%></td>
                  </tr>
                  <tr>
                    <td height="40">IP被拦截时的提示</td>
                  </tr>
                  <tr>
                    <td><%	

	Dim strTOTORO_KILLIPSTR
	strTOTORO_KILLIPSTR=Totoro_Config.Read("TOTORO_KILLIPSTR")
		strTOTORO_KILLIPSTR=TransferHTML(strTOTORO_KILLIPSTR,"[html-format]")
		Response.Write "<textarea rows=""6"" name=""strTOTORO_KILLIPSTR"" style=""width:99%"" >"& strTOTORO_KILLIPSTR &"</textarea>"	
%></td>
                  </tr>
                </table>
              </div>
              <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
                <dl class="totoro">
                  <dd>Totoro是个采用评分机制的防止垃圾留言的插件，原作<a href="http://www.zblogcn.com/" target="_blank">zx.asd</a>。<br/>
                    TotoroⅡ是<a href="http://ZxMYS.COM" target="_blank">Zx.MYS</a>在Totoro的基础上修改而成的增强版，加入了诸多新特性，同时修正一些问题。<br/>
                    TotoroⅢ是由<a href="http://www.zsxsoft.com" target="_blank">zsx</a>将TotoroII升级到2.0版本后增添新特性的版本。</dd>
                  <dd>Spam Value(SV)初始值为0，经过相关运算后的SV分值越高Spam嫌疑越大，超过设定的阈值这条评论就进入审核状态或直接被删除。</dd>
                  <dd>配置完成之后，请一定要测试，切记切记！</dd>
                  <dd></dd>
                </dl>
              </div>
              <div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab6">
               <dl class="totoro">
              <dd>若您的配置因为主机的关键词过滤而失效，您可以点击右侧按钮初始化Totoro设置以修复。<input style="float:right" class="button" type="button" value="初始化Totoro设置" onClick="if(confirm('您确定要初始化Totoro设置吗？该操作不可逆！')){$('#edit').attr('action','savesetting.asp?act=delall');$('#edit').submit()}"/></dd>
              <dd></dd>
              </dl>
              </div>
              </div>
              
              <!-- End .content-box-content -->
              
              </div>
              <p>
                <input type="submit" class="button" value="提交" id="btnPost" onclick='document.getElementById("edit").action="savesetting.asp";' />
              </p>
            </form>
          </div>
        </div>
        <script type="text/javascript">
		ActiveLeftMenu("aCommentMng");
		$(document).ready(function(){
			$("#unescape_badword").text(unescape($("#escape_badword").val()));
			$("#unescape_replace").text(unescape($("#escape_replace").val()));
			$("form").submit(function(){
				$("#escape_badword").text(escape($("#unescape_badword").val()));
				$("#escape_replace").text(escape($("#unescape_replace").val()));
			});
		})</script> 
      </div>
    </div>
    <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
