function getGravatarNow(){
     var emailMD5=MD5($("#inpEmail").val());
     var $obj=$("#gravatarNow>img");
     $obj.attr("src",$obj.attr("src").replace(/avatar\/[^?]*\?s=/i,"avatar/"+emailMD5+"?s=")); 
	 }
