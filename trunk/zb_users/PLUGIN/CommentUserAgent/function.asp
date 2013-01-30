<script language="javascript" runat="server">
var commentuseragent={}
commentuseragent["functions"]={
	"submenu":function(id){
        var json = {
            name: ["main.asp"]
            ,
            cls: ["m-left", "m-left", "m-left"]
            ,
            text: ["首页"]
            ,
            level: [5]

        };
        var str = "";
        for (var i = 0; i < json.name.length; i++) {
            if (BlogUser.Level <= json.level[i]) {
                str += MakeSubMenu(json.text[i], json.name[i], json.cls[i] + (id == i ? " m-now ": ""), false)
            }

        }
        return str

	}

};




</script>
<!-- #include file="detect_device.asp"-->
<!-- #include file="detect_os.asp"-->
<!-- #include file="detect_webbrowser.asp"-->
<!-- #include file="detect_browser_ver.asp"-->
<!-- #include file="detect_platform.asp"-->
