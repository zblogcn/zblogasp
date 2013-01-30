<script language="javascript" runat="server" >
function detect_platform(useragent)
{
	var json;
	json=detect_device(useragent);
	if(json.filename.length>0)
	{
		return json;
	}
	else 
	{	
		json=detect_os(useragent);
		if(json.filename.length > 0)
		{
			return json;
		}
		else{
			json["filename"]="Unknown";
			json["link"]="#";
			json["code"]="null";
			json["fullfilename"]="16/os/null.png";
			json["folder"]="os";
			json["ver"]="";
		}
	}
		return json
}

</script>