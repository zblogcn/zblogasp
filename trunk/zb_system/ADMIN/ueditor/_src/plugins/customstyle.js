///import core
///commands 自定义样式
///commandsName  CustomStyle
///commandsTitle  自定义样式
UE.plugins['customstyle'] = function() {
    var me = this;
    me.setOpt({ 'customstyle':[
        {tag:'h1', label:'居中标题', style:'font-size:32px;font-weight:bold;line-height:38px;border-bottom:#ccc 2px solid;padding:0 4px 0 0;text-align:center;margin:0 0 20px 0;'},
        {tag:'h1', label:'居左标题', style:'font-size:32px;font-weight:bold;line-height:38px;border-bottom:#ccc 2px solid;padding:0 4px 0 0;text-align:left;margin:0 0 10px 0;'},
        {tag:'span', label:'强调', style:'font-size:16px;font-style:italic;font-weight:bold;color:#000;line-height:18px;'},
        {tag:'span', label:'明显强调', style:'font-size:16px;font-style:italic;font-weight:bold;color:rgb(51, 153, 204);line-height:18px;'}
    ]});
    me.commands['customstyle'] = {
        execCommand : function(cmdName, obj) {
            var me = this,
                    tagName = obj.tag,
                    node = domUtils.findParent(me.selection.getStart(), function(node) {
                        return node.getAttribute('label')
                    }, true),
                    range,bk,tmpObj = {};
            for (var p in obj) {
                tmpObj[p] = obj[p]
            }
            delete tmpObj.tag;
            if (node && node.getAttribute('label') == obj.label) {
                range = this.selection.getRange();
                bk = range.createBookmark();
                if (range.collapsed) {
                    //trace:1732 删掉自定义标签，要有p来回填站位
                    if(dtd.$block[node.tagName]){
                        var fillNode = me.document.createElement('p');
                        domUtils.moveChild(node, fillNode);
                        node.parentNode.insertBefore(fillNode, node);
                        domUtils.remove(node)
                    }else{
                        domUtils.remove(node,true)
                    }

                } else {

                    var common = domUtils.getCommonAncestor(bk.start, bk.end),
                            nodes = domUtils.getElementsByTagName(common, tagName);
                    if(new RegExp(tagName,'i').test(common.tagName)){
                        nodes.push(common);
                    }
                    for (var i = 0,ni; ni = nodes[i++];) {
                        if (ni.getAttribute('label') == obj.label) {
                            var ps = domUtils.getPosition(ni, bk.start),pe = domUtils.getPosition(ni, bk.end);
                            if ((ps & domUtils.POSITION_FOLLOWING || ps & domUtils.POSITION_CONTAINS)
                                    &&
                                    (pe & domUtils.POSITION_PRECEDING || pe & domUtils.POSITION_CONTAINS)
                                    )
                                if (dtd.$block[tagName]) {
                                    var fillNode = me.document.createElement('p');
                                    domUtils.moveChild(ni, fillNode);
                                    ni.parentNode.insertBefore(fillNode, ni);
                                }
                            domUtils.remove(ni, true)
                        }
                    }
                    node = domUtils.findParent(common, function(node) {
                        return node.getAttribute('label') == obj.label
                    }, true);
                    if (node) {

                        domUtils.remove(node, true)

                    }

                }
                range.moveToBookmark(bk).select();
            } else {
                if (dtd.$block[tagName]) {
                    this.execCommand('paragraph', tagName, tmpObj,'customstyle');
                    range = me.selection.getRange();
                    if (!range.collapsed) {
                        range.collapse();
                        node = domUtils.findParent(me.selection.getStart(), function(node) {
                            return node.getAttribute('label') == obj.label
                        }, true);
                        var pNode = me.document.createElement('p');
                        domUtils.insertAfter(node, pNode);
                        domUtils.fillNode(me.document, pNode);
                        range.setStart(pNode, 0).setCursor()
                    }
                } else {

                    range = me.selection.getRange();
                    if (range.collapsed) {
                        node = me.document.createElement(tagName);
                        domUtils.setAttributes(node, tmpObj);
                        range.insertNode(node).setStart(node, 0).setCursor();

                        return;
                    }

                    bk = range.createBookmark();
                    range.applyInlineStyle(tagName, tmpObj).moveToBookmark(bk).select()
                }
            }

        },
        queryCommandValue : function() {
            var parent = utils.findNode(this.selection.getStartElementPath(),null,function(node){return node.getAttribute('label')});
            return  parent ? parent.getAttribute('label') : '';
        },
        queryCommandState : function() {
            return this.highlight ? -1 : 0;
        }
    };
    //当去掉customstyle是，如果是块元素，用p代替
    me.addListener('keyup', function(type, evt) {
        var keyCode = evt.keyCode || evt.which;

        if (keyCode == 32 || keyCode == 13) {
            var range = me.selection.getRange();
            if (range.collapsed) {
                var node = domUtils.findParent(me.selection.getStart(), function(node) {
                    return node.getAttribute('label')
                }, true);
                if (node && dtd.$block[node.tagName] && domUtils.isEmptyNode(node)) {
                        var p = me.document.createElement('p');
                        domUtils.insertAfter(node, p);
                        domUtils.fillNode(me.document, p);
                        domUtils.remove(node);
                        range.setStart(p, 0).setCursor();


                }
            }
        }
    })
};