///import core
///commands 自动排版
///commandsName  autotypeset
///commandsTitle  自动排版
/**
 * 自动排版
 * @function
 * @name baidu.editor.execCommands
 */

UE.plugins['autotypeset'] = function(){

    this.setOpt({'autotypeset':{
        mergeEmptyline : true,          //合并空行
            removeClass : true,            //去掉冗余的class
            removeEmptyline : false,        //去掉空行
            textAlign : "left",             //段落的排版方式，可以是 left,right,center,justify 去掉这个属性表示不执行排版
            imageBlockLine : 'center',      //图片的浮动方式，独占一行剧中,左右浮动，默认: center,left,right,none 去掉这个属性表示不执行排版
            pasteFilter : false,             //根据规则过滤没事粘贴进来的内容
            clearFontSize : false,           //去掉所有的内嵌字号，使用编辑器默认的字号
            clearFontFamily : false,         //去掉所有的内嵌字体，使用编辑器默认的字体
            removeEmptyNode : false,         // 去掉空节点
            //可以去掉的标签
            removeTagNames : utils.extend({div:1},dtd.$removeEmpty),
            indent : false,                  // 行首缩进
            indentValue : '2em'             //行首缩进的大小
    }});
    var me = this,
        opt = me.options.autotypeset,
        remainClass = {
            'selectTdClass':1,
            'pagebreak':1,
            'anchorclass':1
        },
        remainTag = {
            'li':1
        },
        tags = {
            div:1,
            p:1,
            //trace:2183 这些也认为是行
            blockquote:1,center:1,h1:1,h2:1,h3:1,h4:1,h5:1,h6:1
        },
        highlightCont;
    //升级了版本，但配置项目里没有autotypeset
    if(!opt){
        return;
    }
    function isLine(node,notEmpty){

        if(node && node.parentNode && tags[node.tagName.toLowerCase()]){
            if(highlightCont && highlightCont.contains(node)
                ||
                node.getAttribute('pagebreak')
            ){
                return 0;
            }

            return notEmpty ? !domUtils.isEmptyBlock(node) : domUtils.isEmptyBlock(node);
        }
    }

    function removeNotAttributeSpan(node){
        if(!node.style.cssText){
            domUtils.removeAttributes(node,['style']);
            if(node.tagName.toLowerCase() == 'span' && domUtils.hasNoAttributes(node)){
                domUtils.remove(node,true)
            }
        }
    }
    function autotype(type,html){

        var cont;
        if(html){
            if(!opt.pasteFilter)return;
            cont = me.document.createElement('div');
            cont.innerHTML = html.html;
        }else{
            cont = me.document.body;
        }
        var nodes = domUtils.getElementsByTagName(cont,'*');

          // 行首缩进，段落方向，段间距，段内间距
        for(var i=0,ci;ci=nodes[i++];){
            if(!highlightCont && ci.tagName == 'DIV' && ci.getAttribute('highlighter')){
                highlightCont = ci;
            }
             //font-size
            if(opt.clearFontSize && ci.style.fontSize){
                ci.style.fontSize = '';
                removeNotAttributeSpan(ci)

            }
            //font-family
            if(opt.clearFontFamily && ci.style.fontFamily){
                ci.style.fontFamily = '';
                removeNotAttributeSpan(ci)
            }

            if(isLine(ci)){
                //合并空行
                if(opt.mergeEmptyline ){
                    var next = ci.nextSibling,tmpNode;
                    while(isLine(next)){
                        tmpNode = next;
                        next = tmpNode.nextSibling;
                        domUtils.remove(tmpNode);
                    }

                }
                 //去掉空行，保留占位的空行
                if(opt.removeEmptyline && domUtils.inDoc(ci,cont) && !remainTag[ci.parentNode.tagName.toLowerCase()] ){
                    domUtils.remove(ci);
                    continue;

                }

            }
            if(isLine(ci,true) ){
                if(opt.indent)
                    ci.style.textIndent = opt.indentValue;
                if(opt.textAlign)
                    ci.style.textAlign = opt.textAlign;
//                if(opt.lineHeight)
//                    ci.style.lineHeight = opt.lineHeight + 'cm';


            }

            //去掉class,保留的class不去掉
            if(opt.removeClass && ci.className && !remainClass[ci.className.toLowerCase()]){

                if(highlightCont && highlightCont.contains(ci)){
                     continue;
                }
                domUtils.removeAttributes(ci,['class'])
            }

            //表情不处理
            if(opt.imageBlockLine && ci.tagName.toLowerCase() == 'img' && !ci.getAttribute('emotion')){
                if(html){
                    var img = ci;
                    switch (opt.imageBlockLine){
                        case 'left':
                        case 'right':
                        case 'none':
                            var pN = img.parentNode,tmpNode,pre,next;
                            while(dtd.$inline[pN.tagName] || pN.tagName == 'A'){
                                pN = pN.parentNode;
                            }
                            tmpNode = pN;
                            if(tmpNode.tagName == 'P' && domUtils.getStyle(tmpNode,'text-align') == 'center'){
                                if(!domUtils.isBody(tmpNode) && domUtils.getChildCount(tmpNode,function(node){return !domUtils.isBr(node) && !domUtils.isWhitespace(node)}) == 1){
                                    pre = tmpNode.previousSibling;
                                    next = tmpNode.nextSibling;
                                    if(pre && next && pre.nodeType == 1 &&  next.nodeType == 1 && pre.tagName == next.tagName && domUtils.isBlockElm(pre)){
                                        pre.appendChild(tmpNode.firstChild);
                                        while(next.firstChild){
                                            pre.appendChild(next.firstChild)
                                        }
                                        domUtils.remove(tmpNode);
                                        domUtils.remove(next);
                                    }else{
                                        domUtils.setStyle(tmpNode,'text-align','')
                                    }


                                }


                            }
                            domUtils.setStyle(img,'float',opt.imageBlockLine);
                            break;
                        case 'center':
                            if(me.queryCommandValue('imagefloat') != 'center'){
                                pN = img.parentNode;
                                domUtils.setStyle(img,'float','none');
                                tmpNode = img;
                                while(pN && domUtils.getChildCount(pN,function(node){return !domUtils.isBr(node) && !domUtils.isWhitespace(node)}) == 1
                                    && (dtd.$inline[pN.tagName] || pN.tagName == 'A')){
                                    tmpNode = pN;
                                    pN = pN.parentNode;
                                }
                                var pNode = me.document.createElement('p');
                                domUtils.setAttributes(pNode,{

                                    style:'text-align:center'
                                });
                                tmpNode.parentNode.insertBefore(pNode,tmpNode);
                                pNode.appendChild(tmpNode);
                                domUtils.setStyle(tmpNode,'float','');

                            }


                    }
                }else{
                    var range = me.selection.getRange();
                    range.selectNode(ci).select();
                    me.execCommand('imagefloat',opt.imageBlockLine);
                }



            }

            //去掉冗余的标签
            if(opt.removeEmptyNode){
                if(opt.removeTagNames[ci.tagName.toLowerCase()] && domUtils.hasNoAttributes(ci) && domUtils.isEmptyBlock(ci)){
                    domUtils.remove(ci)
                }
            }
        }
        if(html)
            html.html = cont.innerHTML
    }
    if(opt.pasteFilter){
        me.addListener('beforepaste',autotype);
    }

    me.commands['autotypeset'] = {
        execCommand:function () {
            me.removeListener('beforepaste',autotype);
            if(opt.pasteFilter){
                me.addListener('beforepaste',autotype);
            }
            autotype();
        }

    };

};

