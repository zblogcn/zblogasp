    </div>
  </div>
</div>

			</div>
<script type="text/javascript">
// <![CDATA[


function changeCheckValue(obj){

	$(obj).toggleClass('imgcheck-on');

	if($(obj).hasClass('imgcheck-on')){
		$(obj).prev('input').val('True');
	}else{
		$(obj).prev('input').val('False');
	}

}




function notify(s){
	if (window.webkitNotifications) {
		if (window.webkitNotifications.checkPermission() == 0) {
			var zb_notifications = window.webkitNotifications.createNotification('<%=GetCurrentHost%>zb_system/IMAGE/ADMIN/logo-16.png', '<%=ZC_MSG257%>', s);
			zb_notifications.onclick = function() {window.parent.focus();this.close();}
			zb_notifications.replaceId = 'Meteoric';
			zb_notifications.show();
		} else {
			window.webkitNotifications.requestPermission(notify);
		}
	} 
}



// Content box tabs:
$('.content-box .content-box-content div.tab-content').hide(); // Hide the content divs
$('ul.content-box-tabs li a.default-tab').addClass('current'); // Add the class "current" to the default tab
$('.content-box-content div.default-tab').show(); // Show the div with class "default-tab"

$('.content-box ul.content-box-tabs li a').click( // When a tab is clicked...
	function() { 
		$(this).parent().siblings().find("a").removeClass('current'); // Remove "current" class from all tabs
		$(this).addClass('current'); // Add class "current" to clicked tab
		var currentTab = $(this).attr('href'); // Set variable "currentTab" to the value of href of clicked tab
		$(currentTab).siblings().hide(); // Hide all content divs
		$(currentTab).show(); // Show the content div with the id equal to the id of clicked tab
		return false; 
	}
);



$(document).ready(function(){ 

	bmx2table();

	if($('.SubMenu').find('span').length>0){
		//if($('#leftmenu').find('li.on').length>0){
		//	$('#leftmenu li.on').after('<li class="sub">'+$('.SubMenu').html()+'</li>');
		//}else{
			$('.SubMenu').show();
		//}
	}

	$("#avatar").attr("src","<%=BlogUser.Avatar%>");

	if(!(($.browser.msie)&&($.browser.version)=='6.0')){
		$('input.checkbox').css("display","none");
		$('input.checkbox[value="True"]').after('<span class="imgcheck imgcheck-on"></span>');
		$('input.checkbox[value="False"]').after('<span class="imgcheck"></span>');
	}else{
		$('input.checkbox').attr('readonly','readonly');
		$('input.checkbox').css('cursor','pointer');
		$('input.checkbox').click(function(){  if($(this).val()=='True'){$(this).val('False')}else{$(this).val('True')} })
	}

	$('span.imgcheck').click(function(){changeCheckValue(this)})

});

// ]]>
</script>
<%=Response_Plugin_Admin_Footer%>
</body>
</html>