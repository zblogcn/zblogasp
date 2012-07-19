    </div>
  </div>
</div>

			</div>
<script language="JavaScript" type="text/javascript">
// <![CDATA[

$(document).ready(function(){ 

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
	}


	if($('.SubMenu').find('span').length==0){$('.SubMenu').hide()}


});
// ]]>
</script>
</body>
</html>