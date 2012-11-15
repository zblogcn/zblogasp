<%
'///////////////////////////////////////////////////////////////////////////////
'//             Z-Blog 2.0
'// 作　　者:    未寒 & seanloo
'// 技术支持:    im@imzhou.com
'// 英文名称:    MusicLink
'// 备　　注:    
'///////////////////////////////////////////////////////////////////////////////
'注册插件
Call RegisterPlugin("MusicLink","ActivePlugin_MusicLink")
'具体的接口挂接

Function ActivePlugin_MusicLink() 
	Dim objConfig,AutoPlay,Player
	Set objConfig=New TConfig
	objConfig.Load("MusicLink")
	If objConfig.Exists("Version")=False Then
		objConfig.Write "Version","0.1"
		objConfig.Write "AutoPlay","True"
		objConfig.Write "Player","yige"
		objConfig.Save
	Else
		AutoPlay=objConfig.Read("AutoPlay")
		Player=objConfig.Read("Player")
	End If

	'网站管理加上二级菜单项
	Dim YM_MusicLink
	YM_MusicLink = "<div style='border-style:dashed;border-color:#AAA;border-width: 2px;width:570px;height:90px' ><script type='text/javascript' src='" & ZC_BLOG_HOST & "zb_users/PLUGIN/MusicLink/addmusic.js'></script><p><span class='editinputname' style='background-color:#ffffff;color:#ff0000;'>插入音乐</span></p><div id='music_key' ><p><span class='editinputname'>歌手：</span><input type='text' id='music_gs' name='music_gs' onfocus=""if(this.value=='戴佩妮') this.value=''"" value='戴佩妮' onblur=""if(this.value=='')this.value=this.defaultValue;""/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class='editinputname'>歌名：</span><input style='width:175px' type='text' id='music_name' name='music_name' onfocus=""if(this.value=='光着我的脚丫子') this.value=''"" value='光着我的脚丫子' onblur=""if(this.value=='')this.value=this.defaultValue;""/><input  value='插入' type='button' class='buttons' onclick='music_Ok();'/></p>  <input name='auto' id='auto' type='hidden' value='" &AutoPlay & "'/> <input name='player' id='player' type='hidden' value='" &Player & "'/> </div></div>"
	
	Call Add_Response_Plugin("Response_Plugin_Edit_Form",YM_MusicLink)
	
End Function

Function InstallPlugin_MusicLink()
	On Error Resume Next
	Call SetBlogHint_Custom("‼ 提示:[Baidu音乐插件]已启用.")
	Err.Clear
End Function

Function UninstallPlugin_MusicLink()
	On Error Resume Next
	Call SetBlogHint_Custom("‼ 提示:[Baidu音乐插件]已禁用.")
	Err.Clear
End Function
%>