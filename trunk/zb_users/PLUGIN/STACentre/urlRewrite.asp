<%
class urlRewrite
	private d
	private d2
	private d3
	private reg
	private key
	public genre
	private sub class_Initialize()
		set d = createObject("Scripting.Dictionary")
		call d.add("{%host%}","(.*?)")
		call d.add("{%post%}","[a-z,\d]+")
		call d.add("{%category%}","[a-z,\d]+")
		call d.add("{%alias%}","([a-z,\d]+)")
		call d.add("{%user%}","(.+?)")
		call d.add("{%year%}","(\d{4})")
		call d.add("{%month%}","(\d{2})")
		call d.add("{%day%}","(\d{2})")
		call d.add("{%id%}","(\d+)")
		call d.add("{%date%}","(\d{4}-\d{2}-\d{2})")
		set d2 = createObject("Scripting.Dictionary")
		call d2.add("ZC_ARTICLE_REGEX",ZC_ARTICLE_REGEX)
		call d2.add("ZC_PAGE_REGEX",ZC_PAGE_REGEX)
		call d2.add("ZC_CATEGORY_REGEX",ZC_CATEGORY_REGEX)
		call d2.add("ZC_USER_REGEX",ZC_USER_REGEX)
		call d2.add("ZC_TAGS_REGEX",ZC_TAGS_REGEX)
		call d2.add("ZC_DATE_REGEX",ZC_DATE_REGEX)
		call d2.add("ZC_DEFAULT_REGEX",ZC_DEFAULT_REGEX)
		set d3 = createObject("Scripting.Dictionary")
		call d3.add("ZC_ARTICLE_REGEX","$1/view\.asp\?id=$2")
		call d3.add("ZC_PAGE_REGEX","$1/view\.asp\?id=$2")
		call d3.add("ZC_CATEGORY_REGEX","$1/catalog\.asp\?cate=$2")
		call d3.add("ZC_USER_REGEX","$1/catalog\.asp\?auth=$2")
		call d3.add("ZC_TAGS_REGEX","$1/catalog\.asp\?tags=$2")
		call d3.add("ZC_DATE_REGEX","$1/catalog\.asp\?date=$2")
		call d3.add("ZC_DEFAULT_REGEX","$1/catalog\.asp\?page=$2")
		set reg = new regExp
			reg.ignoreCase = true
			reg.global = true
	end sub
	private sub class_Terminate()
		set d = nothing
		set d2 = nothing
		set d3 = nothing
		set reg = nothing
	end sub
	public sub display()
		if ZC_STATIC_MODE = "REWRITE" then
			dim ary,j,u,u2,s
			ary = array()
			for each u in d2
				j = ubound(ary) + 1
				reDim preserve ary(j)
				ary(j) = d2(u)
				if u = "ZC_DEFAULT_REGEX" then ary(j) = replace(ary(j),"default.html","default_(\d+).html")
				for each u2 in d
					ary(j) = replace(ary(j),u2,d(u2))
				next
				call search(ary(j),d3(u))
				if u = "ZC_CATEGORY_REGEX" or u = "ZC_USER_REGEX" or u = "ZC_TAGS_REGEX" then
					j = ubound(ary) + 1
					reDim preserve ary(j)
					ary(j) = d2(u)
					if inStr(ary(j),"{%alias%}") then
						ary(j) = replace(ary(j),"{%alias%}","{%alias%}{%page%}",1,1,1)
					elseif inStr(ary(j),"{%id%}") then
						ary(j) = replace(ary(j),"{%id%}","{%id%}{%page%}",1,1,1)
					end if
					for each u2 in d
						ary(j) = replace(ary(j),u2,d(u2))
					next
					ary(j) = replace(ary(j),"{%page%}","_(\d+)")
					call search(ary(j),d3(u))
				end if
			next
			call create(ary)
		end if
	end sub
	private sub search(byref pattern,byref s)
		select case genre
			case "iis7":
				reg.pattern = "\$(\d+)"
				s = reg.replace(s,"{R:$1}")
			case "isapi2":
			case "isapi3":
		end select
		pattern = "^"&replace(pattern,"/default.html","(/)?")&"$" & chr(32) & s
	end sub
	private sub create(ary)
		dim s,i
		select case genre
			case "iis7":
				s = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf
				s = s & "<configuration>" & vbcrlf
				s = s & "    <system.webServer>" & vbcrlf
				s = s & "        <rewrite>" & vbcrlf
				s = s & "            <rules>" & vbcrlf
				for i = lbound(ary) to ubound(ary)
					ary(i) = split(ary(i),chr(32))
					s = s & "                <rule name=""Imported Rule "&i&""" stopProcessing=""true"">" & vbcrlf
					s = s & "                    <match url="""&ary(i)(0)&""" ignoreCase=""false"" />" & vbcrlf
					s = s & "                    <action type=""Rewrite"" url="""&ary(i)(1)&""" />" & vbcrlf
					s = s & "                </rule>" & vbcrlf
				next
				s = s & "            </rules>" & vbcrlf
				s = s & "        </rewrite>" & vbcrlf
				s = s & "    </system.webServer>" & vbcrlf
				s = s & "</configuration>" & vbcrlf
			case "isapi2":
				s = "[ISAPI_Rewrite]" & vbcrlf
				for i = lbound(ary) to ubound(ary)
					s = s & "RewriteRule "&ary(i)& vbcrlf
				next
			case "isapi3":
			for i = lbound(ary) to ubound(ary)
				if instr(ary(i),"$1/catalog\.asp\?tags=$2") > 0 then ary(i) = ary(i) & " [NU]"
				s = s & "RewriteRule "&ary(i)& vbcrlf
			next
		end select
		Response.Write(s)
	end sub
end class
dim ur
set ur = new urlRewrite
	ur.genre = "iis7"
	ur.display()
	ur.genre = "isapi2"
	ur.display()
	ur.genre = "isapi3"
	ur.display()
set ur = nothing
%>