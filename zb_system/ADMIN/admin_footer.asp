    </div>
  </div>
</div>

			</div>
<script type="text/javascript">
// <![CDATA[
$(document).ready(function(){ 

	$("#nav_category").before("<li class='split'><hr/></li>");
	$("#nav_themes").before("<li class='split'><hr/></li>");
	$("#nav_user").before("<li class='split'><hr/></li>");

	$("#avatar").attr("src","<%=BlogUser.Avatar%>");
	$("img[width='16']").each(function(){if($(this).parent().is("a")){$(this).parent().addClass("button")}});
});
// ]]>
</script>
<%=Response_Plugin_Admin_Footer%>
</body>
</html>