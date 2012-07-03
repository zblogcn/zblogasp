var objTextbox;
function FillText(objHTML,strText) {
	if(strText==""){return("")}
	if(!document.getElementById(objHTML)){return false;}
	var obj=document.getElementById(objHTML);
	obj.value=strText;
	obj.focus();
}

function makeGlobalHint(sType,sText){
	sType=sType.toLowerCase();
	var sBgColor;
	switch(sType){
		case 'error' :
			sBgColor='ffd0d5';
			break
		case 'help' :
			sBgColor='b9d4f8';
			break
		case 'done' :
			sBgColor='e5ffdf';
			break
		default :
			sBgColor='b9d4f8';
			break
	}
	StopClosing();
	var $obj=$('#globalHint');
	$obj.fadeOut('normal',function(){
		$obj.css('background-color','#'+sBgColor);
		$obj.children('span').html(sText);
		$obj.fadeIn('normal');
		closingGlobalHint();
	});
}
var timer;
function closingGlobalHint(){
	timer=setTimeout(function(){$('#globalHint').fadeOut('normal');},6000);
}
function StopClosing(){
	clearTimeout(timer);
}
function getValueSet(){
	$("input.checkValue").each(function(){
		if($(this).val().toLowerCase()==="true"){
			$(this).parent("td").children("a").children("span.enable").addClass("checked");
			$(this).parent("td").children("a").children("span.disable").removeClass("checked");
		}else{
			$(this).parent("td").children("a").children("span.enable").removeClass("checked");
			$(this).parent("td").children("a").children("span.disable").addClass("checked");
		}
	});
}


function toggleCheckbox(objToggle){

	//toggle
	var $objSet=$(objToggle).parent('td').children('input.checkValue');
	if($objSet.val().toLowerCase()==='true'){
		$objSet.val('False');
		$(objToggle).children('span.enable').removeClass('checked');
		$(objToggle).children('span.disable').addClass('checked');
	}else{
		$objSet.val('True');
		$(objToggle).children('span.enable').addClass('checked');
		$(objToggle).children('span.disable').removeClass('checked');
	}

	//check directory
	if($objSet.attr('class').indexOf('relValue')!==-1){
		var $objCheck=$(objToggle).parent('td').parent('tr').children('td').children('input.inputValue');
		var strCheck=$objCheck.val();
		var sType=$objCheck.attr('name');
		if($objSet.val().toLowerCase()==='true'){
			if(strCheck.indexOf("{%alias%}")===-1 && strCheck.indexOf("{%id%}")===-1 && strCheck.indexOf("{%name%}")===-1){
				makeGlobalHint('error','路径设置中必须含有 {%alias%},{%id%},{%name%} 三者之一, 已自动补全路径设置.');
				if(sType.toLowerCase().indexOf('categorys')!==-1){strCheck+='/{%alias%}';}
				if(sType.toLowerCase().indexOf('tags')!==-1){strCheck+='/{%name%}';}
				if(sType.toLowerCase().indexOf('authors')!==-1){strCheck+='/{%alias%}';}
				if(sType.toLowerCase().indexOf('archives')!==-1){strCheck+='/{%name%}';}
				strCheck=formatValue(strCheck);
				$objCheck.val(strCheck);
				$objCheck.focus();
			}
		}
	}

	//val changed
	ValueChanged();

}

function checkValue(idCheck){

	if(!document.getElementById(idCheck)){return false;}

	//check alias
	var $objCheck=$('#'+idCheck);
	var strCheck=$objCheck.val();
	if(strCheck.indexOf("{%alias%}")===-1 && strCheck.indexOf("{%id%}")===-1 && strCheck.indexOf("{%name%}")===-1){
		var $objSet=$objCheck.parent('td').parent('tr').children('td').children('input.relValue');
		if($objSet.val().toLowerCase()==='true'){
			makeGlobalHint('error','路径设置没有 {%alias%},{%id%},{%name%} 三者之一, 匿名路径已停用.');
			$objSet.val('false');
			$objSet.parent("td").children("a").children("span.enable").removeClass("checked");
			$objSet.parent("td").children("a").children("span.disable").addClass("checked");
		}
	}

	//format value
	strCheck=formatValue(strCheck);
	$objCheck.val(strCheck);
	$objCheck.focus();

	//val changed
	ValueChanged();


}

function formatValue(strValue){
	strValue=strValue.replace(/\\/g,'/');
	strValue=strValue.replace(/[\/]+/g,'/');
	if(strValue.match(/[#:\?\*\"<>\s|]/i)!==null){makeGlobalHint('error','不要使用 #,:,?,*,\",<,> 和空格等非法字符.');}
	strValue=strValue.replace(/[#:\?\*\"<>\s|]/g,'');
	return strValue;
}

//ajax preview
var loadingTimer;
function Preview(sType){

	ExportLoading();

	var sDate = new Date();
	var sUrl = 'build.asp?act=view&type='+sType+'&rndtm='+sDate.getTime();

	switch(sType)
		{
		case 'Categorys':
			$.post(sUrl,
				{
				"STACentre_Dir_Categorys_Enable": $('#STACentre_Dir_Categorys_Enable').val(),
				"STACentre_Dir_Categorys_Regex": $('#STACentre_Dir_Categorys_Regex').val(),
				"STACentre_Dir_Categorys_Anonymous": $('#STACentre_Dir_Categorys_Anonymous').val(),
				"STACentre_Dir_Categorys_FCate": $('#STACentre_Dir_Categorys_FCate').val()
				},
				function(data){
				ExportPreview(sType,data);
			});
			break
		case 'Tags':
			$.post(sUrl,
				{
				"STACentre_Dir_Tags_Enable": $('#STACentre_Dir_Tags_Enable').val(),
				"STACentre_Dir_Tags_Regex": $('#STACentre_Dir_Tags_Regex').val(),
				"STACentre_Dir_Tags_Anonymous": $('#STACentre_Dir_Tags_Anonymous').val()
				},
				function(data){
				ExportPreview(sType,data);
			});
			break
		case 'Authors':
			$.post(sUrl,
				{
				"STACentre_Dir_Authors_Enable": $('#STACentre_Dir_Authors_Enable').val(),
				"STACentre_Dir_Authors_Regex": $('#STACentre_Dir_Authors_Regex').val(),
				"STACentre_Dir_Authors_Anonymous": $('#STACentre_Dir_Authors_Anonymous').val()
				},
				function(data){
				ExportPreview(sType,data);
			});
			break
		case 'Archives':
			$.post(sUrl,
				{
				"STACentre_Dir_Archives_Enable": $('#STACentre_Dir_Archives_Enable').val(),
				"STACentre_Dir_Archives_Regex": $('#STACentre_Dir_Archives_Regex').val(),
				"STACentre_Dir_Archives_Anonymous": $('#STACentre_Dir_Archives_Anonymous').val(),
				"STACentre_Dir_Archives_Format": $('#STACentre_Dir_Archives_Format').val()
				},
				function(data){
				ExportPreview(sType,data);
			});
			break
		default:
			return false;
		}

	return false;
}
function ExportLoading(){
	$('a.previewBtn',$('#setting')).removeClass('crrView');
	$('#preview>p>a.crrView').removeClass('crrView');
	$('#preview>table').html('<tr><td id=\'previewPanel\'>正在生成预览 </td></tr>');
	StartLoadingProgress('previewPanel');
}
function ExportPreview(sType,data){
	StopLoadingProgress();
	$('a.previewBtn[onclick*='+sType+']',$('#setting')).addClass('crrView');
	$('#preview>p>a[onclick*='+sType+']').addClass('crrView');
	$('#preview>table').html(data);
}
function StartLoadingProgress(sId){
	loadingTimer=setInterval(function(){LoadingProgressStatus(sId);},800);
}
function StopLoadingProgress(){
	clearInterval(loadingTimer);
}
function LoadingProgressStatus(sId){
	var $obj=$("#"+sId);
	var text=$obj.text();
	if(text.indexOf(">>>>>")>1){$obj.text(text.replace(">>>>>",""));}else{$obj.text(text+">");}
}


function ValueChanged(){
	//Preview
	$('a.previewBtn',$('#setting')).removeClass('crrView');
	$('#preview>p>a.crrView').removeClass('crrView');
	$('#preview>table').html('<tr><td>请点击菜单来重新生成预览 ↑</td></tr>');
	//Postbtn
	$('#btnPost').removeAttr('disabled');
	//Buildbtn
	$('#btnBuild').attr('disabled','disabled');
}


function SaveSetting(){

	$('#btnPost').attr('disabled','disabled');

	var sDate = new Date();
	$.post('build.asp?act=save&rndtm='+sDate.getTime(),
		{
		"STACentre_Dir_Categorys_Enable": $('#STACentre_Dir_Categorys_Enable').val(),
		"STACentre_Dir_Categorys_Regex": $('#STACentre_Dir_Categorys_Regex').val(),
		"STACentre_Dir_Categorys_Anonymous": $('#STACentre_Dir_Categorys_Anonymous').val(),
		"STACentre_Dir_Categorys_FCate": $('#STACentre_Dir_Categorys_FCate').val(),

		"STACentre_Dir_Tags_Enable": $('#STACentre_Dir_Tags_Enable').val(),
		"STACentre_Dir_Tags_Regex": $('#STACentre_Dir_Tags_Regex').val(),
		"STACentre_Dir_Tags_Anonymous": $('#STACentre_Dir_Tags_Anonymous').val(),

		"STACentre_Dir_Authors_Enable": $('#STACentre_Dir_Authors_Enable').val(),
		"STACentre_Dir_Authors_Regex": $('#STACentre_Dir_Authors_Regex').val(),
		"STACentre_Dir_Authors_Anonymous": $('#STACentre_Dir_Authors_Anonymous').val(),

		"STACentre_Dir_Archives_Enable": $('#STACentre_Dir_Archives_Enable').val(),
		"STACentre_Dir_Archives_Regex": $('#STACentre_Dir_Archives_Regex').val(),
		"STACentre_Dir_Archives_Anonymous": $('#STACentre_Dir_Archives_Anonymous').val(),
		"STACentre_Dir_Archives_Format": $('#STACentre_Dir_Archives_Format').val()
		},
		function(data){
			$('#btnBuild').removeAttr('disabled');
			if(data.toLowerCase()==='true'){
				$('#setting').slideUp('normal');
				$('#preview').slideUp('normal');
				$('#buildStatus').html('<img src=\"point.gif\" style=\"float:left\"/><p>点此重建静态列表!</p>');
				makeGlobalHint('done','静态路径配置已成功修改, 请<a href=\"javascript:void(0);\" onclick=\"$(\'#btnBuild\').click();\">[重建静态列表页]</a>, 并执行文件重建.');
				$('#ShowBlogHint').html('<p class=\"hint hint_blue\"><font color=\"blue\">‼ 提示:需要进行\"<a href=\"../../../ZB_SYSTEM/admin/admin.asp?act=AskFileReBuild\">[文件重建]</a>\".</font></p>');
			}
			if(data.toLowerCase()==='false'){
				makeGlobalHint('done','静态路径配置已成功修改, 但配置数据并未变更.');
			}

	});

	return false;
}


function PageRebuild(iNum){
	$('#btnBuild').attr('disabled','disabled');
	if($('#pregressWrapper').size()===0){$('#buildStatus').html('<div id=\"pregressWrapper\" style=\"float:left;width:380px;height:10px;border:1px solid #99AAFF;\"><div id=\"progressbar\" style=\"height:10px;width:0;background:#6699FA;\"></div></div><div id=\"percent\" style=\"float:left;padding:0 0 0 10px;\">0%</div>');}
	var sdate = new Date();
	$.ajax({
		type: 'GET', dataType: 'html', timeout: 12000,
		url: 'build.asp?act=build&tasknum='+iNum+'&rndtm='+sdate.getTime(),
		error: function(){
			setTimeout(function(){PageRebuild(iNum);},400);
		},
		success: function(data){
			var n=parseInt(data.split('/')[0]);
			var m=parseInt(data.split('/')[1]);
			var o=parseInt((n/m)*380);
			var p=parseInt((n/m)*100);
			if(n>iNum){iNum=n;}
			if(iNum>m){
				setTimeout(function(){
					$('#buildStatus').text('静态列表页生成完毕!');
					makeGlobalHint('done','静态列表页重新生成完毕!');
					$('#btnBuild').removeAttr('disabled');
					$('#setting').slideDown('normal');
					$('#preview').slideDown('normal');
				},400);
			}else{
				//$("#progressbar").width(o);
				$('#progressbar').animate({'width':o+'px'},500); //only can be use in JQ 1.4+
				//$("#pregressWrapper").text(data);
				setTimeout(function(){$('#percent').text(p+'%');},400);
				iNum++;
				setTimeout(function(){PageRebuild(iNum);},400);
			}
		}
	});
	return false;
}