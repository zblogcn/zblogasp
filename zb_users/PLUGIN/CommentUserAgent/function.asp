<script language="javascript" runat="server">
var commentuseragent={}
commentuseragent["functions"]={
	"submenu":function(id){
        var json = {
            name: ["main.asp?list=0","main.asp?list=1","main.asp?list=2"]
            ,
            cls: ["m-left", "m-left","m-right"]
            ,
            text: ["首页","说明","测试"]
            ,
            level: [1,1,1]

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
