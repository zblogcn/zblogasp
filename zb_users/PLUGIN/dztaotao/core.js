$(document).ready(function() {
	$("#uploadify").uploadify({
		'uploader': 'uploadify.swf',					//指定上传控件的主体文件，默认'uploader.swf'
		'script': 'upload.asp',									//指定服务器端上传处理文件
		'cancelImg': 'images/cancel.jpg',						//指定取消上传的图片，默认'cancel.png'
		'buttonImg': 'images/button.jpg',			//指定浏览文件按钮的图片
		//'fileDesc' : '图片文件',				//出现在上传对话框中的文件类型描述
		//'fileExt' : '*.jpg;*.jpeg;*.gif;*.bmp;*.png',						//控制可上传文件的扩展名，启用本项时需同时声明fileDesc
		'sizeLimit': 2000000, 								//控制上传文件的大小，单位byte		服务器默认只支持30MB，修改服务器设置请查看相关资料
		//'simUploadLimit' :5,								//多文件上传时，同时上传文件数目限制
		'buttonText':'choose',								//按钮显示文字,不支持中文，要用中文直接用背景图片cancelImg设置
		//'folder': 'upload',									//要上传到的服务器路径(PS：已在服务端设置)
		'queueID': 'fileQueue',								//队列
		'fileDataName': 'Filedata',						//提交文件域名称
		'auto': true,												//选定文件后是否自动上传，默认false
		'multi': false,												//是否允许同时上传多文件，默认false
		'method':'post',											//提交方式 post or get
		//'scriptData'  : {'firstName':'Ronnie','age':30},							//提交自定义数据
		onComplete:function(event,queueID,fileObj,response,data){				//上传成功执行
		
			//返回服务端JSON数据,可在服务端修改返回数据类型
			var filetext=eval("(" + response + ")")					//解析JSON数据
			//alert(response);
			switch   ( filetext.err )   {   
        case	0	:
					document.getElementById('u_img').value = filetext.savename;
					document.getElementById('s_img').value = filetext.s_savename;
					/*$("#RequestText").prepend('保存时间：' + filetext.time + "<br /><br />");
					$("#RequestText").prepend('文件类型：' + filetext.ext + "<br />");
					$("#RequestText").prepend('文件大小：' + filetext.size + "<br />");
					$("#RequestText").prepend('保存路径：' + filetext.path + "<br />");
					$("#RequestText").prepend('保存文件名：' + filetext.savename + "<br />");
					$("#RequestText").prepend('缩略图文件名：' + filetext.s_savename + "<br />");
					$("#RequestText").prepend('<br />文件名：' + filetext.name + "<br />");
					$("#RequestText").prepend('<img src=' + filetext.path + '/' +filetext.savename+">");
					*/
					break ;
        case	1	:
          $("#RequestText").prepend('<br />文件因过大而未被保存<br />');
					break ;
				case	2	:
          $("#RequestText").prepend('<br />文件类型因不匹配而未被保存<br />');
					break ;
				case	3	:
          $("#RequestText").prepend('<br />文件因过大并且类型不匹配而未被保存<br />');
					break ;
				case	-1	:
          $("#RequestText").prepend('<br />没有文件上传<br />');
					break ;
			}
			
			
			//返回uploadify数据
			//$("#RequestText").append('文件名：' + fileObj.name + "<br />");
			//$("#RequestText").append('文件大小：' + fileObj.size + "<br />");
			//$("#RequestText").append('创建时间：' + fileObj.creationDate + "<br />");
			//$("#RequestText").append('最后修改时间：' + fileObj.modificationDate + "<br />");
			//$("#RequestText").append('文件类型：' + fileObj.type + "<br />");
			
		},
	   
	});
});  


//显示发表淘淘表单
function showDialog()
{
	var sdialog = document.getElementById('dialog');
	sdialog.style.display = 'block';
}

//关闭淘淘表单
function closeDialog()
{
	document.getElementById('s_content').value == '';
	var cdialog = document.getElementById('dialog');
	cdialog.style.display = 'none';
}

//显示分享按钮
function shareLayer(sbtn)
{
	var share_btn = document.getElementById('share-'+sbtn);
	share_btn.style.display = 'block';
}

//打开回复列表
function showReply(sbid)
{
	var Reply_id = document.getElementById('item-comment-'+sbid);
	if(Reply_id.style.display == 'none')
	{
		Reply_id.style.display = 'block';
	}
	else{
		Reply_id.style.display = 'none';
	}
}

//插入新增淘淘
function insertTaotao(i_msg,i_content,i_user,i_site,i_img,i_s_img)
{
	var t_time = new Date();//"2011-4-8 11:33:33";
	var t_datetime = t_time.getFullYear()+'-'+t_time.getMonth()+'-'+t_time.getDate()+' '+t_time.toLocaleTimeString();
	var t_idd = i_msg;
	var tt_img ="";
	if(i_s_img !="" && i_img !=""){tt_img = "<a href='upload/"+i_img+"' rel='upload/"+i_img+"' class='miniImg artZoom'><img src='upload/"+i_s_img+"'></a>";}
	
	$("#taotao").prepend("<div class='item' id='item-"+t_idd+"'><div class='item-list'>    	<div class='list-text' id='listText-"+t_idd+"'>"+ i_content +"<br>"+tt_img+"</div>                <div class='list-text'>                             <div class='list-interaction'> <div id='shareLayer' class='share-layer'><dl class='item-share'><dt>分享到:</dt><dd><a href='http://service.weibo.com/share/share.php?url=http://www.izhu.org/plugin/dztaotao/view.asp?id="+i_msg+"&type=3&count=&appkey=&title="+encodeURIComponent("大猪淘淘——")+ encodeURIComponent(i_content) +"&pic="+i_img+"&ralateUid=&rnd=1337756006442' target='_blank' title='转帖到新浪微博' id='share_sina' class='btn-share-sina'></a></dd><dd><a href='http://share.renren.com/share/buttonshare.do?link=http://www.izhu.org/plugin/dztaotao/view.asp?id="+i_msg+"&title="+encodeURIComponent("大猪淘淘——")+ encodeURIComponent(i_content)  +"' target='_blank' title='转帖到人人网' class='btn-share-rr'></a></dd><dd><a href='###' title='转帖到开心网' id='share_kx' class='btn-share-kx'></a></dd><dd><a href='http://share.v.t.qq.com/index.php?c=share&a=index&appkey=&site=http://www.izhu.org/&title="+encodeURIComponent("大猪淘淘——")+ encodeURIComponent(i_content) +"&url=http://www.izhu.org/plugin/dztaotao/view.asp?id="+i_msg+"' target='_blank' title='推荐到QQ微博' id='share_tqq' class='btn-share-tqq'></a></dd></dl></div></div>                                   <div class='clear'></div></div></div>        <div class='item-infor'>    	<div class='infor-text'><img src='/PLUGIN/dztaotao/images/default.jpg'> <span>"+ i_user +"</span> <span>"+ t_datetime +" 发布</span></div>        <div class='infor-set'><a href='javascript:;' onfocus='this.blur()' class='btn-up' onclick='dingUp("+t_idd+")'>称赞</a> <span class='scroe-up highlight' id='ding_"+t_idd+"'>0</span> <a href='javascript:;' onfocus='this.blur()' class='btn-down' onclick='dingDown("+t_idd+")'>鄙视</a> <span id='tread_"+t_idd+"' class='scroe-down highlight'>0</span> | <a href='javascript:;' title='点击展开评论' onfocus='this.blur()' id='commtent-"+t_idd+"' class='comment' onclick='showReply("+t_idd+")'>评论(0)</a></div></div>        <div id='item-comment-"+t_idd+"' style='display:none' class='item-comment'>        <div class='clear'></div>                         <div style='padding: 10px 10px 0pt;' class='blue-con' id='blueCon-"+t_idd+"'>            <table border='0'><tbody><tr>            <td><div id='shortcut-key"+t_idd+"'></div></td>            </tr>            <tr>            <td><textarea name='r_content_"+t_idd+"' class='comment-textarea' id='r_content_"+t_idd+"'></textarea></td>            </tr>               <tr style='display:none'>            <td>昵称：<input type='text' id='r_username_"+t_idd+"' name='r_username_"+t_idd+"'>    邮箱：<input type='text' id='r_email_"+t_idd+"' name='r_email_"+t_idd+"'>    网址：<input type='text' id='r_site_"+t_idd+"' name='r_site_"+t_idd+"'></td>            </tr>            </tbody>            </table>            <div class='discuss-login'><a onclick='postCmt("+t_idd+")' href='javascript:;' class='btn-send' id='send-"+t_idd+"'>发表评论</a><span class='comments-leave'>您还可以输入<strong class='highlight' id='lwords-"+t_idd+"'>400</strong>个字符</span></div>        </div>            <div class='comment-msg' id='msg-"+t_idd+"'></div>                      <div class='comment-list' id='comments-"+t_idd+"'>            <div id='newInsertCmt"+t_idd+"'></div>        </div>            <div class='comment-all' id='all-"+t_idd+"'>共有0条评论 | <a onclick='showReply("+t_idd+")' href='javascript:;'>收起评论</a> | <a href='view.asp?id="+t_idd+"'>更多</a></div>    </div></div>");
	//$("#taotao").append("<div id='item-001' class='item' style='background:#FFC;'><div class='item-list'><div id='listText-new001' class='list-text'>"+ i_content +"</div><div class='list-text'><div style='visibility:hidden' class='tag'></div><div class='list-interaction'><a id='share-item-new001' class='btn-share' href='###'>分享</a></div><div class='clear'></div></div></div><div class='item-infor'><div class='infor-text'><img src='images/default.jpg'> <span>"+ i_user +"</span> <span>"+ t_datetime +" 发布</span></div><div class='infor-set'><a class='btn-up' onfocus='this.blur()' href='###'>称赞</a> <span class='scroe-up highlight'>0</span> <a  class='btn-down' onfocus='this.blur()' href='###'>鄙视</a> <span class='scroe-down highlight'>0</span> | <a class='comment' id='commtent-new001' onfocus='this.blur()' title='点击展开评论' href='###'>评论(0)</a></div></div><div class='item-comment' id='item-comment-new001' style='display: none;'><img height='8' width='14' class='item-comment-title-img' src='images/item-comment-arrow.png'><div class='clear'></div>      <div id='all-new001' class='comment-all'>共有0条评论</div></div></div>");
	//document.getElementById('newInsert').style.display = 'block';
	
}

//插入新增加的评论
function insertCmt(imsg,i_content,i_user,i_site,TID)
{
	var t_time = new Date();//"2011-4-8 11:33:33";
	var t_datetime = t_time.getFullYear()+'-'+t_time.getMonth()+'-'+t_time.getDate()+' '+t_time.toLocaleTimeString();
	$("#comments-"+TID).prepend("<div class='item' id='jitem-"+imsg+"'><div class='comment-box'><a class='discuss-pic' href='"+i_site+"'><img width='32' height='32' src='http://passport.maxthon.cn/_image/avatar-demo.png'></a><div class='discuss-con'><div class='con-bar dash-boder'><a class='name' href='http://haha.mx/user/4141252'>"+i_user+"</a><span class='time'>"+t_datetime+"发表</span> </div><p>"+i_content+"</p></div><div class='clear'></div></div></div>");
	
}

//清空评论框
function clearCmt(TID)
{
	$("#r_content_"+TID).val("");
	//$("#r_username_"+TID).val("");
	//$("#r_email_"+TID).val("");
	//$("#r_site_"+TID).val("");
}

//随机产生评论用户名
function r_random_user(TID)
{
	var userID = Math.floor(Math.random()*10+2);
	var userName = "";
	if(userID == 1){
		userName = '春香';
	}else if(userID == 2){
		userName = '秋香';
	}else if(userID == 3){
		userName = '夏香';
	}else if(userID == 4){
		userName = '冬香';
	}else if(userID == 5){
		userName = '华文';
	}else if(userID == 6){
		userName = '华武';
	}else if(userID == 7){
		userName = '华安';
	}else if(userID == 8){
		userName = '东淫';
	}else if(userID == 9){
		userName = '西贱';
	}else if(userID == 10){
		userName = '南荡';
	}else if(userID == 11){
		userName = '北色';
	}else{
		userName = '灭绝师太';
	}
	document.getElementById('r_username_'+TID).value = userName;
}


//获取评论提交内容
function r_getString(TID)
{
	var r_content = encodeURIComponent($("#r_content_"+TID).val());
	var r_username = encodeURIComponent($("#r_username_"+TID).val());
	var r_email = encodeURIComponent($("#r_email_"+TID).val());
	var r_site = $("#r_site_"+TID).val();
	var qString = "c="+r_content+"&u="+r_username+"&s="+r_site+"&e="+r_email;
	return qString;
}

//提交评论
function postCmt(TID)
{
	//alert(r_getString(TID))
	$("#shortcut-key"+TID).html("<div style='text-align:center;color:#090;'><img src='images/load.gif' border='0'>正在提交中...</div>");
	$.ajax({
		type:	"POST",
		url:	"r.asp?t=r&tid="+TID,
		data:	r_getString(TID),
		success:	function(msg){
			$("#shortcut-key"+TID).html(msg);
			if(msg == 0){
				$("#shortcut-key"+TID).html("<div style='text-align:center;color:red;'>信息要写全才能提交哦！</div>");
			}else if(msg == -4){
				$("#shortcut-key"+TID).html("<div style='text-align:center;color:red;'>添加失败了，悲剧！</div>");
			}else if(msg == -111){
				$("#shortcut-key"+TID).html("<div style='text-align:center;color:red;'>你已经评论过这条见鬼的信息了，不要再评论了！</div>");
			}else if(msg >0){
				$("#shortcut-key"+TID).html("<div style='text-align:center;color:red;'>添加成功</div>");
				insertCmt(msg,$("#r_content_"+TID).val(),$("#r_username_"+TID).val(),$("#r_site_"+TID).val(),TID);
				clearCmt(TID);
			}
			
			
		}
	});
}

//获取顶票
function ding_rread(TID)
{
	var qString = "tid="+TID;
	return qString;
}


//提交淘淘支持票
function dingUp(TID)
{
		$.ajax({
		   type: "POST",
		   url:	"r.asp?t=dingup",//顶
		   data:	ding_rread(TID),
		   success:	function(msg){
			   if(msg > 0)
			   {
				   $("#ding_"+TID).html(msg);
			   }
			   else
			   {
				   alert('您已经搞过了，不要不停的搞好不好，人家会受不鸟的!');
			   }
		   }
		});

}

//提交淘淘反对票
function dingDown(TID)
{
		$.ajax({
		   type: "POST",
		   url:	"r.asp?t=dingdown",//踩
		   data:	ding_rread(TID),
		   success:	function(msg){
			   if(msg > 0)
			   {
				   $("#tread_"+TID).html(msg);
			   }
			   else
			   {
				   alert('您已经搞过了，不要不停的搞好不好，人家会受不鸟的!');
			   }
		   }
		});

}


//获取提交内容
function getString()
{
	var content = encodeURIComponent($("#s_content").val());
	var username = encodeURIComponent($("#username").val());
	var site = $("#s_site").val();
	var img = $("#u_img").val();
	var s_img = $("#s_img").val();
	var qString = "c="+content+"&u="+username+"&s="+site+"&img="+img+"&s_img="+s_img;
	return qString;
}

//提交淘淘验证
function subInfo()
{
	b = $("#s_content").val();
	if(b.indexOf('[IMG]图片地址[/IMG]')>-1)
	{
		alert('图片地址不正确啊');
		return;
	}
	//$("#abc").html(getString());
	//alert(getString());
	$("#msg").html("<div style='text-align:center;color:#090;'>正在提交中...</div>");
	$("#msg").html("<img src='images/load.gif' border='0'>");
	$.ajax({
		   type: "POST",
		   url:	"r.asp?t=p",
		   data:	getString(),
		   success:	function(msg){
			   if(msg == 0){
				   $("#msg").html("<div style='text-align:center;color:red;'>信息要写全才能提交哦！</div>");
			   }else if(msg == -1){
				   $("#msg").html("<div style='text-align:center;color:red;'>您添加的信息已经有了哦！</div>");
			   }else if(msg == -4){
				   $("#msg").html("<div style='text-align:center;color:red;'>添加失败了，悲剧！</div>");
			   }else if(msg > 0){
				   $("#msg").html("<div style='text-align:center;color:red;'>添加成功</div>");
				   closeDialog();
				   insertTaotao(msg,$("#s_content").val(),$("#username").val(),$("#s_site").val(),$("#u_img").val(),$("#s_img").val());
	document.getElementById('u_img').value = "";
	document.getElementById('s_img').value = "";
	document.getElementById('s_content').value = "";
	$("#msg").html(" ");

			   }else{
				   $("#msg").html("<div style='text-align:center;color:red;'>出错了，请重试！</div>");
			   }
			   
			   
			   			   
		   }
		});
}


//设置共享打开的窗口
function open_share(sName,sUrl)
{
	var vlink = encodeURIComponent(document.location); // 文章链接
	var title = encodeURIComponent(document.title.substring(0,76)); // 文章标题
	var source = encodeURIComponent('网站名称'); // 网站名称
	var windowName = 'share'; // 子窗口别称
	var site = 'http://www.example.com/'; // 网站链接

	if(sName != "" && sUrl != "")
	{
		if(sName == 'sina'){
			
		}
	}
}


