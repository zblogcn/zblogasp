﻿    </div>
  </div>
</div>

			</div>
<script type="text/javascript">
// <![CDATA[
$(document).ready(function(){ 

	$("#avatar").attr("src","<%=BlogUser.Avatar%>");
	$("img[width='16']").each(function(){if($(this).parent().is("a")){$(this).parent().addClass("button")}});

	if($("p.hint_green:visible").length>0){
		$("p.hint_green:visible").eq(0).delay(1500).hide(1500,function(){});
	}

	$("input[type='file']").click(function(){

		if(/IEMobile|WPDesktop/g.test(navigator.userAgent)&&$(this).val()==""){
			alert('<%=ZVA_ErrorMsg(65)%>')
		}
	})
});
// ]]>
</script>
<%=Response_Plugin_Admin_Footer%>
</body>
</html>