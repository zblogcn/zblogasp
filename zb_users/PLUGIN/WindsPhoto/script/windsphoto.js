//按比例缩放图片
function WindsPhotoResizeImage(objImage,maxWidth,maxHeight) {
try{
  if(maxWidth>0 && maxHeight>0){
   var objImg = $(objImage);
   var objImg = $(objImage);
   if(objImg.width()>maxWidth && objImg.width()>objImg.height() || objImg.height()==objImg.width()){
    objImg.width(maxWidth).css("cursor","pointer")
    objImg.height(objImg.height() / (objImg.width() / maxWidth)).css("cursor","pointer")
   }
  else if(objImg.height()>maxHeight && objImg.height()>objImg.width() || objImg.height()==objImg.width()){
    objImg.height(maxHeight).css("cursor","pointer")
    bjImg.width(objImg.width() / (objImg.height() / maxHeight)).css("cursor","pointer")
   }
  }
}catch(e){};
}

//检查上传表单
function CheckForm()
{
if ((document.upload.file0.value.length == 0) && (document.upload.url.value.length == 0))  {
alert("一个也没填?!");
document.upload.file1.focus();
return false;
}
document.getElementById('upupup').value = '上传中...';
document.getElementById('upupup').disabled=true;
return true;
}

//实现预览图片（仿xspace后台图片上传（本地图片不支持ff））
var maxWidth=170;
var maxHeight=170;
var fileTypes=["jpg","gif","png","bmp","jpeg"];
var outImage="previewField";
var defaultPic="images/nopic.jpg";
var globalPic;

function preview(what){
  var source=what.value;
  var ext=source.substring(source.lastIndexOf(".")+1,source.length).toLowerCase();
  for (var i=0; i<fileTypes.length; i++) if (fileTypes[i]==ext) break;
  globalPic=new Image();
  if (i<fileTypes.length) globalPic.src=source;
  else {
    globalPic.src=defaultPic;
    what.outerHTML = what.outerHTML.replace(/value=\w/g,"");
    alert("图片格式限制: "+fileTypes.join(", "));
  }
  setTimeout("applyChanges()",200);
}

function applyChanges(){
  var field=document.getElementById(outImage);
  var x=parseInt(globalPic.width);
  var y=parseInt(globalPic.height);
  if (x>maxWidth) {
    y*=maxWidth/x;
    x=maxWidth;
  }
  if (y>maxHeight) {
    x*=maxHeight/y;
    y=maxHeight;
  }
  field.style.display=(x<1 || y<1)?"none":"";
  field.src=globalPic.src;
  field.width=x;
  field.height=y;
}

//键盘快捷键翻页
function ToPage(event){
event = event ? event : (window.event ? window.event : null);
if (event.keyCode==37) location=prevpage
if (event.keyCode==38) location=prevpage
if (event.keyCode==39) location=nextpage
if (event.keyCode==40) location=nextpage
}