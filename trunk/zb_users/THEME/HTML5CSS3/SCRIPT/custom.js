/* -----copyright------
 Custom Script for Z-Blog Theme Html5CSS3
 Author: jgpy.cn
 Pub Date: 2013-1-16
 Last Modified: 2013-1-21
----------------------*/

//网站加载完毕后初始化
blog.js.int=function(){
	//修改默认回复语言
	blog.msg.cmt.reply="[回复]";
	//有评论表单就初始化
	blog.form.cmt[0]&&CmtForm(blog.form);
	//判断是否有评论
	if(blog.cmt.box[0]){
		blog.js.getCmt()
	}else{
		blog.cmt.list.hide();
	}
	//判断是否有相关文章
	$("#related").html(function(i,s){
		!$("li",s)[0]&&$(this).hide();
	})
	//不支持HTML5
	NoHTML5();
	//启用当前页高亮
	blog.nav=$("body header nav ul");
	//管理员登录链接
	Admin();
	//文章内容图片居中
	$("article>section img").parent("p").css("text-align","center");
	//标签云
	if(blog.url=="tags.asp"){
		$("a.tag-name").each(function(){
			this.style.fontSize=(parseInt(this.title)+90)+"%"
		})
	}
};
//AJAX评论
blog.js.ajaxCmt=function(){
	var $form=blog.form,
		$btn=$(":submit",$form.cmt),
		tip=blog.msg.cmt,
		$tip=$("#cmttip"),
		$cmt=blog.cmt,
		$rtip=$("#replytip");
	$(".err").removeClass("err");
	$tip.empty().stop().show(200);;
	var val={
		name:$.trim($form.name.val()),
		email:$.trim($form.email.val()),
		homepage:$.trim($form.homepage.val()),
		txt:$.trim($form.txt.val()),
		verify:$.trim($form.verify.val())
	};
	tip.record=blog.user("name")!=val.name?"您的资料已记录":"";
	if(val.txt==""||val.txt.length>blog.msg.max){
		$form.txt.addClass("err").focus();
		$tip.text(tip.msg);
		return false;
	};
	if(!val.name){
		$form.name.addClass("err").focus();
		$tip.text(tip.name);
		return false;
	};
	if(val.email!=""&&!val.email.match("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$")){
		$form.email.addClass("err").focus();
		$tip.text(tip.email);
		return false;
	};
	if($form.verify[0]&&!val.verify){
		$form.verify.addClass("err").focus();
		$tip.text("请输入验证码");
		return false;
	};
	$tip.text("提交中，请稍候……");
	$btn.attr("disabled",true).fadeTo(500,.5);
	$.post($form.cmt[0].action,{
		"inpAjax":true,
		"inpID":$form.id.val(),
		"inpVerify":($form.verify[0]?val.verify:""),
		"inpEmail":val.email,
		"inpName":val.name,
		"inpArticle":val.txt,
		"inpHomePage":val.homepage?val.homepage:blog.host,
		"inpRevID":blog.replyID
	},function(s){
		if((s.search("faultCode")>0)&&(s.search("faultString")>0)){
			tip.err=s.match("<string>.+?</string>")[0].replace("<string>","").replace("</string>","")
			$tip.text(tip.err);
			if(tip.err=="验证码输入错误"){
				$form.verify.val("").addClass("err").focus();
				$form.validcode.trigger("click");
			}
		}else{
			if(blog.replyID!=0){
				$("#cancelreply").trigger("click").after(s);
			}else{
				var $ins=$cmt.begin.siblings("blockquote:last");
				if($ins[0]){
					$ins.after(s);
				}else{
					$cmt.begin.after(s);
					$cmt.list.fadeIn();
				}
			}
			$form.txt.val("").blur();
			if($form.verify[0]){
				$form.verify.val("");
				$form.validcode.trigger("click");
			}
			blog.cookies.set("inpName",val.name,365);
			blog.cookies.set("inpEmail",val.email,365);
			blog.cookies.set("inpHomePage",val.homepage,365);
			$tip.text("发表成功！"+tip.record).delay(5000).fadeOut(500);
		}
		$btn.attr("disabled",false).fadeTo(500,1);
	});
	return false;	
};
//AJAX获取评论成功后回调
blog.js.getCmt=function(){
	$("#comment blockquote").each(function(){
		var id=this.id.substr(3);
		$(">q",this).children().first().before(blog.cmt.reply(id,true).click(function(){ReplyForm(id,this)}))
	});
	blog.cmt.post.show(500);
};
//侧栏加载成功后回调
blog.js.sidebar=function(){
	$("view,count").each(function(){
        $(this).replaceWith($(this).text());
     });
	$("#btnPost").addClass("btn");
	try{
		if(!$.trim($("aside#extra").html())) $("aside#extra").html("<dl><dt>温馨提示</dt><dd>请在后台【侧栏管理】分配部分侧栏模块到【侧栏2】！</dd></dl>");
	}catch(e){}
};
//创建回复评论
function ReplyForm(id,self){
	if($("#cancelreply")[0]){
		$("#cancelreply").hide(500,function(){
			$(this).remove();
		}).prev("ins").show(500);
	}
	var $cmt=blog.cmt,
		$form=blog.form,
		$cancel=$("<b id='cancelreply'>[取消回复]</b>").click(function(){
			$(this).hide(500,function(){
				$(this).remove();
			}).prev("ins").show(500);
			$("#postreply").hide(500,function(){
				$cmt.post.show(500).find("dd").append($form.cmt.hide().show(500));
				$(this).remove();
			});
			$cmt.list.find("nav").show(500);
			blog.replyID=0;
		});
	$("#postreply")[0]&&$form.cmt.unwrap();
	$(self).hide(500).after($cancel);
	$form.cmt.wrap("<div id='postreply'/>").parent().hide(500,function(){
		$(this).show(500,function(){
			$form.txt.focus();
		}).insertAfter($cancel);
	})
	$cmt.post.hide(500);
	$cmt.list.find("nav").hide(500);
};
//初始化评论表单
function CmtForm($form){
	var $avatar=$("<figure><img src='"+blog.user("avatar")+"' alt='"+blog.user("name")+"的头像' border='0'/></figure>").insertBefore($form.cmt),
		$name=$("<b>"+blog.user("name")+"</b>").appendTo($avatar),
		$notme=$("<ins style='cursor:pointer'>[不是我]</ins>").click(function(){
			$form.name.val("").parents("p").show(500);
			$form.email.val("").parents("p").show(500);
			$form.homepage.val("").parents("p").show(500);
			$(this).fadeOut(500);
			$name.text("无名氏");
			$avatar.find("img").attr("src","http://gravatar.com/avatar/");
		});
	if(!blog.user("id"))$avatar.append($notme);
	$form.cmt.find(":submit").after("<b id='cmttip'/>");
	$form.txt.blur(function(){
		$.support.opacity&&$(this).css("background",this.value!=""?"#fff":"");
	}).attr({title:"",tabindex:$.support.opacity?1:6}).blur();
	$form.email.blur(function(){
		var val=this.value;
		val&&$.getScript(blog.sys+"script/md5.js",function(){
			$avatar.find("img").attr("src","http://gravatar.com/avatar/"+MD5(val));
		})
	}).blur();
	$form.name.keyup(function(){
		$name.text(this.value?this.value:"无名氏")
	}).blur(function(){
		$name.text(this.value?this.value:"无名氏")
	})
	$form.name.val()&&$form.name.parents("p").hide();
	$form.email.val()&&$form.email.parents("p").hide();
	$form.homepage.val()&&$form.homepage.parents("p").hide();
};
//管理链接
function Admin(){
	if($("#divContorPanel")[0]) return;
	if(blog.user("level")=="管理员"){
		$("body>footer h4:last").append(" [<a target='_blank' href='"+blog.sys+"cmd.asp?act=login'>"+blog.msg._248+"</a>]");
	}
}
//不支持HTML5的浏览器提示
function NoHTML5(){
	var IEstr=$.support.opacity?"":"<big>珍爱生命，远离IE！或升级到IE9+吧！</big>";
	$("body").prepend("<hgroup><h3 id='nohtml5'>"+IEstr+"对不起！您的浏览器不支持HTML5，像iPhone5和WIN8一样，HTML5，你值得拥有！请升级你的浏览器到最新版本，以获得更佳浏览体验，感谢您对互联网的贡献及对<big>HTML5</big>的认可！</h3></hgroup>");
	$("#nohtml5").css("display",function(i,s){
		if(s=="none")$(this).parent().remove();
	})
}