<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 
'// 作    者:    zsx
'// 技术支持:   zsx@zsxsoft.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    edit_article.asp
'// 开始时间:    2012.9.6
'// 最后修改:
'// 备    注:    编辑页
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
For Each sAction_Plugin_Edit_Article_Begin in Action_Plugin_Edit_Article_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Article_Begin) Then Call Execute(sAction_Plugin_Edit_Article_Begin)
Next

'检查非法链接
Call CheckReference("")
'检查权限
If Not CheckRights("ArticleEdt") Then Call ShowError(6)

Call GetUser()

Dim IsPage

IsPage=Request.QueryString("type")="Page"

If IsPage Then If CheckRights("Root")=False And CheckRights("ArticleAll")=False Then Call ShowError(6)

Dim EditArticle
Set EditArticle=New TArticle

If Not IsEmpty(Request.QueryString("id")) Then
	If EditArticle.LoadInfobyID(Request.QueryString("id")) Then
		If EditArticle.AuthorID<>BlogUser.ID Then
			If CheckRights("Root")=False And CheckRights("ArticleAll")=False Then
				Call ShowError(6)
			End If
		End If
		If EditArticle.FType=ZC_POST_TYPE_PAGE Then IsPage=True
		If InStr(EditArticle.Intro,"<!--autointro-->")>0 Then EditArticle.Intro=""
	Else
		Call ShowError(9)
	End If
Else
	EditArticle.AuthorID=BlogUser.ID
	If IsPage=True THen EditArticle.FType=ZC_POST_TYPE_PAGE
End If


EditArticle.Content=UBBCode(EditArticle.Content,"[link][email][font][code][face][image][flash][typeset][media][autolink]")
EditArticle.Title=UBBCode(EditArticle.Title,"[link][email][font][code][face][image][flash][typeset][media][autolink]")

EditArticle.Content=TransferHTML(Replace(EditArticle.Content,"<!--more-->","<hr class=""more"" />"),"[html-japan]")

EditArticle.Title=TransferHTML(EditArticle.Title,"[html-format]")

BlogTitle=IIf(IsPage,ZC_MSG161,ZC_MSG047)

'为1,2,3号输出输口准备的Action接口
'plugin node
For Each sAction_Plugin_Edit_Form in Action_Plugin_Edit_Form
	If Not IsEmpty(sAction_Plugin_Edit_Form) Then Call Execute(sAction_Plugin_Edit_Form)
Next

%>
<!--#include file="admin_header.asp"-->

<script type="text/javascript" src="../script/jquery.tagto.js"></script>
<script type="text/javascript" src="../script/jquery-ui-timepicker-addon.js"></script>
<%=Response_Plugin_Edit_Article_Header%>
<!--#include file="admin_top.asp"-->
        <%If IsPage=False Then%>
        <%If EditArticle.ID=0 Then%>
        <script type="text/javascript">ActiveLeftMenu("aArticleEdt");</script>
        <%Else%>
        <script type="text/javascript">ActiveLeftMenu("aArticleMng");</script>
        <%End If%>
        <%Else%>
        <script type="text/javascript">ActiveLeftMenu("aPageMng");</script>
        <%End If%>
        <div id="divMain">
          <%	Call GetBlogHint()	%>
          <div class="divHeader2"><%=IIf(IsPage,ZC_MSG161,ZC_MSG047)%></div>
          <%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleEdt_SubMenu & "</div>"
%>
          <div id="divMain2">
            <form id="edit" name="edit" method="post" action="">
              <div id="divEditLeft">
                <div id="divEditTitle">
                  <input type="hidden" name="edtID" id="edtID" value="<%=EditArticle.ID%>" />
                  <input type="hidden" name="edtFType" id="edtFType" value="<%=EditArticle.FType%>" />
                  <!-- title( -->
                  <p><span class='editinputname'><%=ZC_MSG060%>:</span>
                    <input type="text" name="edtTitle" id="edtTitle" style="width:60%;max-width:520px" maxlength="100" onBlur="if(this.value=='') this.value='<%=ZC_MSG099%>'" onFocus="if(this.value=='<%=ZC_MSG099%>') this.value=''" value="<%=EditArticle.Title%>" />
                  </p>
                  <!-- )title --> 
                  
                  <!-- alias( -->
                  <p><span class='editinputname'><%=ZC_MSG147%>:</span>
                    <input type="text" style="width:60%;max-width:520px" name="edtAlias" id="edtAlias" maxlength="250" value="<%=TransferHTML(EditArticle.Alias,"[html-format]")%>" />
                    .<%=ZC_STATIC_TYPE%> </p>
                  <!-- )alias --> 
                  
                  <!-- tags( -->
                  <% If Request.QueryString("type")<>"Page" Then %>
                  <p><span class='editinputname' style='padding:0 0 0 0;'><%=ZC_MSG138%>:</span>
                    <input type="text" style="width:60%;max-width:520px" name="edtTag" id="edtTag" value="<%=TransferHTML(EditArticle.TagToName,"[html-format]")%>" />
                    (<%=ZC_MSG208%>) <a href="#" id="showtags"><%=ZC_MSG139%></a></p>
                  <!-- Tags -->
                  <div id="ulTag" style="display:none;">
                    <div id="ajaxtags"><%=ZC_MSG165%></div>
                  </div>
                  <% End If %>
                  <!-- )tags --> 
                  
                </div>
                
                <!-- 1号输出接口 -->
                <% If Response_Plugin_Edit_Form<>"" Then %>
                <div id="divEditForm1"><%=Response_Plugin_Edit_Form%></div>
                <% End If %>
                <div id="divContent" style="clear:both;">
                  <p style="text-align:left;"><span class='editinputname'><%=ZC_MSG055%>:</span>&nbsp;&nbsp;<span id="timemsg"></span><span id="msg2"></span><span id="msg"></span><span class='editinputname'></span><script type="text/javascript" src="c_autosaverjs.asp?act=edit"></script></p>
                  <textarea id="editor_txt" name="txaContent" style="display:none"><%=TransferHTML(EditArticle.Content,"[textarea]")%></textarea>
                  <div id="contentready"><img id="statloading" src="../image/admin/loading.gif"/><%=ZC_MSG276%></div>
                  <p><span><%=ZC_MSG216%><a href="" onClick="try{AutoIntro();return false;}catch(e){}">[<%=ZC_MSG200%>]</a></span></p>
                </div>
                <div id="divIntro" style="display:<%If EditArticle.Intro="" Then Response.Write "none" Else Response.Write "block"%>;">
                  <p><span class='editinputname'><%=ZC_MSG016%>:</span></p>
                  <textarea id="editor_txt2" name="txaIntro" style="display:none"><%=TransferHTML(EditArticle.Intro,"[textarea]")%></textarea>
                  <div id="introready"><img id="statloading" src="../image/admin/loading.gif"/><%=ZC_MSG276%></div>
                </div>
                
                <!-- 2号输出接口 -->
                <% If Response_Plugin_Edit_Form2<>"" Then %>
                <div id="divEditForm2"><%=Response_Plugin_Edit_Form2%></div>
                <% End If %>
              </div>
              <!-- divEditLeft -->
              
              <div id="divEditRight">
                <div id="divEditPost">
                  <div id="divBox">
                    <div id="divFloat">
                      <p>
                        <input class="button" style="width:150px;height:30px;" type="submit" value="<%=ZC_MSG087%>" id="btnPost" onclick='return checkArticleInfo();' />
                      </p>
                      
                      <!-- cate -->
                      <p>
                        <%
If Request.QueryString("type")<>"Page" Then
%>
                        <span class='editinputname'><%=ZC_MSG012%>:</span>
                        <select style="width:150px;" class="edit" size="1" id="cmbCate" onChange="edtCateID.value=this.options[this.selectedIndex].value;selectlogtemplate(this.options[this.selectedIndex].value);">
                          <option value="0"><%=Categorys(0).Name%></option>
                          <%
	Dim aryCateInOrder : aryCateInOrder=GetCategoryOrder()
	Dim m,n
	For m=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
		If Categorys(aryCateInOrder(m)).ParentID=0 Then
			Response.Write "<option value="""&Categorys(aryCateInOrder(m)).ID&""" "
			If EditArticle.CateID=Categorys(aryCateInOrder(m)).ID Then Response.Write "selected=""selected"""
			Response.Write ">"&TransferHTML( Categorys(aryCateInOrder(m)).Name,"[html-format]")&"</option>"

			For n=0 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(n)).ParentID=Categorys(aryCateInOrder(m)).ID Then
					Response.Write "<option value="""&Categorys(aryCateInOrder(n)).ID&""" "
					If EditArticle.CateID=Categorys(aryCateInOrder(n)).ID Then Response.Write "selected=""selected"""
					Response.Write ">&nbsp;└ "&TransferHTML( Categorys(aryCateInOrder(n)).Name,"[html-format]")&"</option>"
				End If
			Next
		End If
	Next
%>
                        </select>
                        <input type="hidden" name="edtCateID" id="edtCateID" value="<%=EditArticle.CateID%>" />
                        <%
Else
%>
                        <input type="hidden" name="edtCateID" id="edtCateID" value="0" />
                        <%
End If
%>
                      </p>
                      <!-- cate --> 
                      
                      <!-- template( -->
                      <p> <span class='editinputname'><%=ZC_MSG188%>:</span>
                        <select style="width:150px;" class="edit" size="1" id="cmbTemplate" onChange="edtTemplate.value=this.options[this.selectedIndex].value">
                          <%
	Dim aryFileList

	aryFileList=LoadIncludeFilesOnlyType("zb_users\theme" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY)

	If IsArray(aryFileList) Then
		Dim j,t,b
		b=False
		j=UBound(aryFileList)
		For i=1 to j
			t=UCase(Left(aryFileList(i),InStr(aryFileList(i),".")-1))
			If Left(t,2)<>"B_" AND t<>"FOOTER" And t<>"HEADER" And t<>"DEFAULT" And t<>"CATALOG" Then
				If EditArticle.GetDefaultTemplateName=t And b=False Then
					Response.Write "<option value="""&t&""" selected=""selected"">"&t&IIF(EditArticle.TemplateName="","("&ZC_MSG187&")","")&"</option>"
					b=True
				Else
					Response.Write "<option value="""&t&""">"&t&"</option>"
				End If
			End If
		Next
	End If

%>
                        </select>
                        <input type="hidden" name="edtTemplate" id="edtTemplate" value="<%=EditArticle.TemplateName%>" />
                      </p>
                      <!-- )template --> 
                      
                      <!-- level -->
                      <p> <span class='editinputname'><%=ZC_MSG061%>:</span>
                        <select class="edit" style="width:150px;" size="1" id="cmbArticleLevel" onChange="edtLevel.value=this.options[this.selectedIndex].value">
                          <%
	Dim ArticleLevel
	Dim i:i=0
	For Each ArticleLevel in ZVA_Article_Level_Name
		If i>0 Then
			Response.Write "<option value="""& i &""" "
			If EditArticle.Level=i Then Response.Write "selected=""selected"""
			Response.Write ">"& ZVA_Article_Level_Name(i) &"</option>"
		End If
		i=i+1
	Next
%>
                        </select>
                        <input type="hidden" name="edtLevel" id="edtLevel" value="<%=EditArticle.Level%>" />
                      </p>
                      <!-- )level --> 
                      
                      <!-- user( -->
                      <p> <span class='editinputname'><%=ZC_MSG003%>:</span>
                        <select style="width:150px;" size="1" id="cmbUser" onChange="edtAuthorID.value=this.options[this.selectedIndex].value">
                          <%
	GetUser()
	Dim User
	For Each User in Users
		If IsObject(User) Then
			If User.Level<4 Then
			If CheckRights("Root")=True Or CheckRights("ArticleAll")=True Then
				Response.Write "<option value="""&User.ID&""" "
				If User.ID=EditArticle.AuthorID Then
					Response.Write "selected=""selected"""
				End If
				Response.Write ">"&TransferHTML(User.Name,"[html-format]")&"</option>"
			Else
				If User.ID=EditArticle.AuthorID Then
					Response.Write "<option value="""&User.ID&""" "
					Response.Write "selected=""selected"""
					Response.Write ">"&TransferHTML(User.Name,"[html-format]")&"</option>"
				End If
			End If
			End If
		End If
	Next

%>
                        </select>
                        <input type="hidden" name="edtAuthorID" id="edtAuthorID" value="<%=EditArticle.AuthorID%>" />
                      </p>
                      <!-- )user --> 
                      
                      <!-- newdatetime( -->
                      <p> <span class='editinputname'><%=ZC_MSG062%>:</span><span>
                        <input type="text" name="edtDateTime" id="edtDateTime"  value="<%=EditArticle.PostTime%>" style="width:141px;"/>
                        </span> </p>
                      <!-- )newdatetime --> 
                      
                      <!-- Istop( -->
                      <p>
                        <% If Request.QueryString("type")<>"Page" And BlogUser.Level=1 Then %>
                        <label><span class='editinputname'><%=ZC_MSG051%>:
                          <%If EditArticle.Istop Then%>
                          <input type="checkbox" name="edtIstop" id="edtIstop" value="True" checked=""/>
                          <%Else%>
                          <input type="checkbox" name="edtIstop" id="edtIstop" value="True"/>
                          <%End If%>
                          </span></label>
                        <%Else%>
                        <input type="hidden" name="edtIstop" id="edtIstop" value=""/>
                        <% End If %>
                      </p>
                      <!-- )Istop --> 
                      
                      <!-- Navbar( -->
                      <% If Request.QueryString("type")="Page" Then %>
                      <p>
                        <label><span class='editinputname'><%=ZC_MSG046%>:
                          <%If EditArticle.ID=0 Then%>
                          <input type="checkbox" name="edtAddNavbar" id="edtAddNavbar" value="True" checked=""/>
                          <%Else%>
                          <input type="checkbox" name="edtAddNavbar" id="edtAddNavbar" value="True"/>
                          <%End If%>
                          </span></label>
                      </p>
                      <% End If %>
                      <!-- )Navbar --> 
                      
                      <!-- 3号输出接口 -->
                      <% If Response_Plugin_Edit_Form3<>"" Then %>
                      <div id="divEditForm3"><%=Response_Plugin_Edit_Form3%></div>
                      <% End If %>
                    </div>
                  </div>
                </div>
              </div>
              <!-- divEditRight -->
              
            </form>
          </div>
        </div>
        <script type="text/javascript">

var tag_loaded=false; //是否已经ajax读取过TAGS
var sContent="",sIntro="";//原内容与摘要
var isSubmit=0;//是否提交保存
var str10="<%=ZC_MSG115%>";
var str11="<%=ZC_MSG116%>";
var editor_api={
	editor:	{
		content:{
			obj:{},
			get:function(){return ""},
			put:function(){return ""},
			focus:function(){return ""}
		},
		intro:{
			obj:{},
			get:function(){return ""},
			put:function(){return ""},
			focus:function(){return ""}
		}
	}
}
var EditorIntroOption = {
	toolbars:[['Source', 'bold', 'italic','link','insertimage','Undo', 'Redo']],
	autoHeightEnabled:false,
	initialFrameHeight:200
};


$(document).click(function (event){$('#ulTag').slideUp("fast");});  

//文章内容或摘要变动提示保存
window.onbeforeunload = function(){
	if (!isSubmit && (sContent!=editor_api.editor.content.get() || sIntro!=editor_api.editor.intro.get())) return "<%=ZC_MSG302%>";
}

editor_init();
<%=Response_Plugin_Edit_Article_EditorInit%>







//日期时间控件
$.datepicker.regional['zh-cn'] = {
	closeText: '完成',
	prevText: '上个月',
	nextText: '下个月',
	currentText: '现在',
	monthNames: ['一月','二月','三月','四月','五月','六月','七月','八月','九月','十月','十一月','十二月'],
	monthNamesShort: ['一月','二月','三月','四月','五月','六月','七月','八月','九月','十月','十一月','十二月'],
	dayNames: ['星期日','星期一','星期二','星期三','星期四','星期五','星期六'],
	dayNamesShort: ['周日','周一','周二','周三','周四','周五','周六'],
	dayNamesMin: ['日','一','二','三','四','五','六'],
	weekHeader: '周',
	dateFormat: 'yy-m-d',
	firstDay: 1,
	isRTL: false,
	showMonthAfterYear: true,
	yearSuffix: ' 年  '
};
$.datepicker.setDefaults($.datepicker.regional['zh-cn']);
$.timepicker.regional['zh-cn'] = {
	timeOnlyTitle: '时间',
	timeText: '时间',
	hourText: '小时',
	minuteText: '分钟',
	secondText: '秒钟',
	millisecText: '毫秒',
	currentText: '现在',
	closeText: '完成',
	timeFormat: 'h:m:s',
	ampm: false
};
$.timepicker.setDefaults($.timepicker.regional['zh-cn']);
$('#edtDateTime').datetimepicker({
	showSecond: true
	//changeMonth: true,
	//changeYear: true
});



function checkArticleInfo(){
	document.getElementById("edit").action="../cmd.asp?act=ArticlePst<%=IIF(Request.QueryString("type")="Page","&type=Page","")%>";

	if(!editor_api.editor.content.get()){
		alert(str11);
		return false
	}
	
	if(encodeURIComponent(editor_api.editor.content.get()).length>(1024*200)){
		if(!confirm("您的文章内容可能过长，受到IIS的限制可能会出错。是否确认继续？")){
			return false
		}
	}
	
	isSubmit=1;
}

//显示tags
$('#showtags').click(function (event) {  
	event.stopPropagation();  
	var offset = $(event.target).offset();  
	$('#ulTag').css({ top: offset.top + $(event.target).height()+20+ "px", left: offset.left});  
	$('#ulTag').slideDown("fast"); 		
	if(tag_loaded==false){$.getScript('../function/c_admin_js.asp?act=tags');tag_loaded=true;}
});  



function AddKey(i) {
	var strKey=$('#edtTag').val();
	var strNow=","+i
	if(strKey==""){
		strNow=i
	}
	if(strKey.indexOf(strNow)==-1){
		strKey=strKey+strNow;
	}
	$('#edtTag').val(strKey);
}
function DelKey(i) {
	var strKey=$('#edtTag').val();
	var strNow="{"+i+"}"
	if(strKey.indexOf(strNow)!=-1){
		strKey=strKey.substring(0,strKey.indexOf(strNow))+strKey.substring(strKey.indexOf(strNow)+strNow.length,strKey.length)
	}
	$('#edtTag').val(strKey);
}

//提取摘要
function AutoIntro() {
	var s=editor_api.editor.content.get();
	editor_api.editor.intro.put("");
	if(s.indexOf("<hr class=\"more\" />")>-1||s.indexOf("<hr class=\"more\"/>")>-1){
		editor_api.editor.intro.put(editor_api.editor.content.get().split("<hr class=\"more\" />")[0]);
		editor_api.editor.intro.put(editor_api.editor.content.get().split("<hr class=\"more\"/>")[0]);
	}else{
		s="";
		var ss=editor_api.editor.content.get().split("</p>");
		for (var t in ss){
			if(s.length<<%=ZC_TB_EXCERPT_MAX%>){
				s+=ss[t].concat("</p>");
			}
		}
		editor_api.editor.intro.put(s);
	}

	$("#divIntro").show();editor_api.editor.intro.focus();
	$('html,body').animate({scrollTop:$('#divIntro').offset().top},'fast');
}

<%

%>
//文章编辑提交区随动JS开始
var oDiv=document.getElementById("divFloat");
var H=0;var Y=oDiv;
while(Y){H+=Y.offsetTop;Y=Y.offsetParent;};
$(window).bind("scroll resize",function(){
	var s=document.body.scrollTop||document.documentElement.scrollTop;
	if(s>H){
		$("#divFloat").addClass("boxfloat");
	}
	else{
		$("#divFloat").removeClass("boxfloat");
	}   
});
<% 

%>
//选择模板
function selectlogtemplate(c){
<%
	Dim Category
	For Each Category in Categorys
		If IsObject(Category) Then
			Response.Write "	if(c=="&Category.ID&"){if('"&Category.LogTemplate&"'!=''){selectlogtemplatesub('"&Category.LogTemplate&"')}else{ selectlogtemplatesub('"&Category.GetDefaultLogTemplateName&"') }};" & vbCrlf
		End If
	Next

%>
}
function selectlogtemplatesub(a){
	$("#cmbTemplate").find("option[value='"+a+"']").attr("selected","selected");
	$("#edtTemplate").val(a);
}

</script> 
        <!--#include file="admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>