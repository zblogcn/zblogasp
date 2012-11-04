function gravatarNow(){
     var emailMD5=hex_md5($("#inpEmail").val());
     var $obj=$("#gravatarNow>img");
     $obj.attr("src",$obj.attr("src").replace(/avatar\/[^?]*\?s=/i,"avatar/"+emailMD5+"?s=")); 
	 }
