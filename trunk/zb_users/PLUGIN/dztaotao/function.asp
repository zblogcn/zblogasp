<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    myllop-大猪
'// 版权所有:    www.izhu.org
'// 技术支持:    myllop#gmail.com
'// 程序名称:    留言增加gravatar头像
'// 英文名称:    dztaotao
'// 开始时间:    2009-5-10
'// 最后修改:    
'// 备    注:    only for zblog1.8
'///////////////////////////////////////////////////////////////////////////////


'生成随机数
Function RndNumber(MaxNum,MinNum)
 Randomize 
 RndNumber=int((MaxNum-MinNum+1)*rnd+MinNum)
 RndNumber=RndNumber
End Function


'随机生成名字
function rndName(rndNum)
 select case rndNum
 	case "1": rndName = "春香"'路人甲
	case "2": rndName = "秋香"'路人甲
	case "3": rndName = "夏香"'路人甲
	case "4": rndName = "冬香"'路人甲
	case "5": rndName = "华文"'路人甲
	case "6": rndName = "华武"'路人甲
	case "7": rndName = "华安"'路人甲
	case "8": rndName = "东淫"'路人甲
	case "9": rndName = "西贱"'路人甲
	case "10": rndName = "南荡"'路人甲
	case "11": rndName = "北色"'路人甲
 end select
 rndName = rndName
end function
%>