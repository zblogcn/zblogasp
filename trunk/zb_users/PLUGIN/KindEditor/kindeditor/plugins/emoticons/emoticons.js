/*******************************************************************************
* @author panderman <panderman@163.com>
* @site http://www.xunwee.com/
* @licence http://www.kindsoft.net/license.php
*******************************************************************************/

KindEditor.plugin('emoticons', function (K) {
    var self = this, name = 'emoticons',arrEmots=[],
		path = (self.emoticonsPath || self.pluginsPath + 'emoticons/images/'),
		emoticonspluginsPath = self.pluginsPath + 'emoticons/',
		allowPreview = self.allowPreviewEmoticons === undefined ? true : self.allowPreviewEmoticons;
    self.clickToolbar(name, function () {
        var elements = [],
			menu = self.createMenu({
			    name: name
			}),
			
            html = '<div class="ke-plugin-emoticons">\
                        <link rel="stylesheet" type="text/css" href="' + emoticonspluginsPath + 'emoticon.css" />\
                        <div id="tabPanel" class="neweditor-tab">\
                            <div id="tabMenu" class="neweditor-tab-h">\
                                <div>精选</div>\
                                <div>兔斯基</div>\
                            </div>\
                            <div id="tabContent" class="neweditor-tab-b">\
                                <div id="tab0"></div>\
                                <div id="tab1"></div>\
                            </div>\
                        </div>\
                    <div id="tabIconReview"><img id="faceReview" class="review" src="' + emoticonspluginsPath + '0.gif" /></div></div>\
                    </div>';
        menu.div.append(K(html));
        function removeEvent() {
            K.each(elements, function () {
                this.unbind();
            });
        }

		
			
        /*******************************************表情方法开始******************************************************/
        function initImgBox(box, str, len) {
            if (box.length) return;
            var tmpStr = "", i = 1;
            for (; i <= len; i++) {
                tmpStr = str;
                if (i < 10) tmpStr = tmpStr + '0';
                tmpStr = tmpStr + i + '.gif';
                box.push(tmpStr);
            }
        }
        function $G(id) {
            return document.getElementById(id)
        }
        function InsertSmiley(url) {
            var obj = {
                src: editor.options.emotionLocalization ? editor.options.UEDITOR_HOME_URL + "dialogs/emotion/" + url : url
            };
            obj.data_ue_src = obj.src;
            editor.execCommand('insertimage', obj);
            dialog.popup.hide();
        }
        function over(td, srcPath, posFlag) {
            td.style.backgroundColor = "#ACCD3C";
            $G('faceReview').style.backgroundImage = "url(" + srcPath + ")";
            if (posFlag == 1) $G("tabIconReview").className = "show";
            $G("tabIconReview").style.display = 'block';
        }
        function bindCellEvent(td) {//表情绑定事件
            td.mouseover(function () {
                var obj = K(this);
                var sUrl = obj.attr("sUrl"),
                    posFlag = obj.attr("posflag");
                obj.css({
                    "background-color": "#ACCD3C"
                });
                $G('faceReview').style.backgroundImage = "url(" + sUrl + ")";
                if (posFlag == "1") $G("tabIconReview").className = "show";
                $G("tabIconReview").style.display = 'block';
            });
            td.mouseout(function () {
                var obj = K(this);
                obj.css({
                    "background-color": "#FFFFFF"
                });
                var tabIconRevew = $G("tabIconReview");
                tabIconRevew.className = "";
                tabIconRevew.style.display = 'none';
            });
            td.click(function (e) {
                self.insertHtml('<img src="' + K(this).attr("realUrl") + '" border="0" alt="" />').hideMenu().focus();
            });
        }
        var emotion = {};
        emotion.SmileyPath = path;
        emotion.SmileyBox = { tab0: [], tab1: [], tab2: [], tab3: [], tab4: [], tab5: [], tab6: [] };
        emotion.SmileyInfor = { tab0: [], tab1: [], tab2: [], tab3: [], tab4: [], tab5: [], tab6: [] };
        var faceBox = emotion.SmileyBox;
        var inforBox = emotion.SmileyInfor;
        var sBasePath = emotion.SmileyPath;
        initImgBox(faceBox['tab0'], 'j_00', 84);
        initImgBox(faceBox['tab1'], 't_00', 40);
   
        //大对象
        FaceHandler = {
            imageFolders: { tab0: 'jx2/', tab1: 'tsj/', tab2: 'ldw/', tab3: 'bobo/', tab4: 'babycat/', tab5: 'face/', tab6: 'youa/' },
            imageWidth: { tab0: 35, tab1: 35, tab2: 35, tab3: 35, tab4: 35, tab5: 35, tab6: 35 },
            imageCols: { tab0: 11, tab1: 11, tab2: 11, tab3: 11, tab4: 11, tab5: 11, tab6: 11 },
            imageColWidth: { tab0: 3, tab1: 3, tab2: 3, tab3: 3, tab4: 3, tab5: 3, tab6: 3 },
            imageCss: { tab0: 'jd', tab1: 'tsj', tab2: 'ldw', tab3: 'bb', tab4: 'cat', tab5: 'pp', tab6: 'youa' },
            imageCssOffset: { tab0: 35, tab1: 35, tab2: 35, tab3: 35, tab4: 35, tab5: 25, tab6: 35 },
            tabExist: [0, 0, 0, 0, 0, 0, 0]
        };
        function switchTab(index) {
            if (FaceHandler.tabExist[index] == 0) {
                FaceHandler.tabExist[index] = 1;
                createTab('tab' + index);
            }
            //获取呈现元素句柄数组
            var tabMenu = $G("tabMenu").getElementsByTagName("div"),
                        tabContent = $G("tabContent").getElementsByTagName("div"),
                        i = 0,
                        L = tabMenu.length;
            //隐藏所有呈现元素
            for (; i < L; i++) {
                tabMenu[i].className = "";
                tabContent[i].style.display = "none";
            }
            //显示对应呈现元素
            tabMenu[index].className = "on";
            tabContent[index].style.display = "block";
        }
        function createTab(tabName) {
            var faceVersion = "?v=1.1", //版本号
                tab = $G(tabName), //获取将要生成的Div句柄
                imagePath = sBasePath + FaceHandler.imageFolders[tabName], //获取显示表情和预览表情的路径
                imageColsNum = FaceHandler.imageCols[tabName], //每行显示的表情个数
                positionLine = imageColsNum / 2, //中间数
                iWidth = iHeight = FaceHandler.imageWidth[tabName], //图片长宽
                iColWidth = FaceHandler.imageColWidth[tabName], //表格剩余空间的显示比例
                tableCss = FaceHandler.imageCss[tabName],
                cssOffset = FaceHandler.imageCssOffset[tabName],
                textHTML = ['<table class="smileytable" cellpadding="1" cellspacing="0" align="center" style="border-collapse:collapse;" border="1" bordercolor="#BAC498" width="100%">'],
                i = 0,
                imgNum = faceBox[tabName].length,
                imgColNum = FaceHandler.imageCols[tabName],
                faceImage,
                sUrl,
                realUrl,
                posflag,
                offset,
                infor;
            for (; i < imgNum; ) {
                textHTML.push('<tr>');
                for (var j = 0; j < imgColNum; j++, i++) {
                    faceImage = faceBox[tabName][i];
                    if (faceImage) {
                        sUrl = imagePath + faceImage + faceVersion;
                        realUrl = imagePath + faceImage;
                        posflag = j < positionLine ? 0 : 1;
                        offset = cssOffset * i * (-1) - 1;
                        infor = inforBox[tabName][i];
                        textHTML.push('<td  class="' + tableCss + '" sUrl="' + sUrl + '" posflag="' + posflag + '" realUrl="' + realUrl.replace(/'/g, "\\'") + '" border="1" width="' + iColWidth + '%" style="border-collapse:collapse;" align="center"  bgcolor="#FFFFFF">');
                        textHTML.push('<span  style="display:block;">');
                        textHTML.push('<img  style="background-position:left ' + offset + 'px;" title="' + infor + '" src="' + sBasePath + '0.gif" width="' + iWidth + '" height="' + iHeight + '"></img>');
                        textHTML.push('</span>');
                    } else {
                        textHTML.push('<td width="' + iColWidth + '%"   bgcolor="#FFFFFF">');
                    }
                    textHTML.push('</td>');
                }
                textHTML.push('</tr>');
            }
            textHTML.push('</table>');
            textHTML = textHTML.join("");
            tab.innerHTML = textHTML;
            K("td[class='" + tableCss + "']").each(function () {//表情绑定事件
                bindCellEvent(K(this));
            });
        }
        //初始显示第一个表情目录
        switchTab(0);
        K("div#tabMenu>div").each(function (index) {//标签切换绑定事件
            K(this).click(function () {
                switchTab(index);
            });
        });
        /**********************************************表情方法结束*****************************************************/
    });
});
