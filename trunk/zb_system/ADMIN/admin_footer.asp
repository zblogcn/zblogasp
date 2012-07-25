    </div>
  </div>
</div>

			</div>
<script language="JavaScript" type="text/javascript">
// <![CDATA[


function bmx2table(){

	//斑马线
	var tables=document.getElementsByTagName("table");
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");
		var b=false;
		cells[0].className="color1";
		for (var i = 1; i < cells.length; i++){
			if(b){
				cells[i].className="color2";
				b=false;
				cells[i].onmouseover=function(){
					this.className="color4";
				}
				cells[i].onmouseout=function(){
					this.className="color2";
				}
			}
			else{
				cells[i].className="color3";
				b=true;
				cells[i].onmouseover=function(){
					this.className="color4";
				}
				cells[i].onmouseout=function(){
					this.className="color3";
				}
			};
		};
	};
};

$(document).ready(function(){ 

	bmx2table();

	if($('.SubMenu').find('span').length==0){$('.SubMenu').hide()};




	$("#avatar").attr("src","<%="http://www.gravatar.com/avatar/"& MD5(BlogUser.Email) &"?s=40&d="& Server.urlEncode(GetCurrentHost & "ZB_SYSTEM/image/admin/avatar.png")%>");



});
// ]]>
</script>
<%=Response_Plugin_Admin_Footer%>
</body>
</html>