<%UBB4ZB18_()
Function UBB4ZB18_getHTML4Article(n,v)
	v(5)=UBB4ZB18(v(5))
	v(4)=UBB4ZB18(v(4))
End Function

Function UBB4ZB18_getHTML4Comment(n,v)
	v(7)=UBB4ZB18(v(7))
End Function
%>
<script language="javascript" runat="server">
function UBB4ZB18_(){RegisterPlugin("UBB4ZB18","UBB4ZB18.ActivePlugin");}
function UBB4ZB18(txt){return txt.UBBtoHTML();}
UBB4ZB18.ActivePlugin=function (){
	Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","UBB4ZB18_getHTML4Article");
	Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_TemplateTags","UBB4ZB18_getHTML4Comment");
	Add_Action_Plugin("Action_Plugin_Edit_ueditor_Begin","ZC_UBB_ENABLE=True")
}
String.prototype.UBBtoHTML=function(){
	ZC_UBB_ENABLE=true;
	return UBBCode(this.toString(),"[link][email][font][code][face][image][flash][typeset][media][autolink][link-antispam]");
}
</script>