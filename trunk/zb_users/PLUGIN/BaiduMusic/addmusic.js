function music_Ok(){
	var gs=EncodeUtf8(document.getElementById("music_gs").value)
	var gm=EncodeUtf8(document.getElementById("music_name").value)
	var bfq="<embed width='500' height='74' type='application/x-shockwave-flash' src='http://box.baidu.com/widget/flash/song.swf?name="+gm+"&artist="+gs+"&autoPlay=true'></embed>"	
	
	if(editor.hasContents() == false){
		editor.setContent(bfq)
	}else{
		var uenr=editor.getContent()
		bfq=uenr + bfq
		editor.setContent(bfq)
	};
	
};
