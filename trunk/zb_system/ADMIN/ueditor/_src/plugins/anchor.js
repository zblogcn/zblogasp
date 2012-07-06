///import core
///commands 锚点
///commandsName  Anchor
///commandsTitle  锚点
///commandsDialog  dialogs\anchor\anchor.html
/**
 * 锚点
 * @function
 * @name baidu.editor.execCommands
 * @param    {String}    cmdName     cmdName="anchor"插入锚点
 */

UE.commands['anchor'] = {
    execCommand:function (cmd, name) {
        var range = this.selection.getRange(),img = range.getClosedNode();
        if (img && img.getAttribute('anchorname')) {
            if (name) {
                img.setAttribute('anchorname', name);
            } else {
                range.setStartBefore(img).setCursor();
                domUtils.remove(img);
            }
        } else {
            if (name) {
                //只在选区的开始插入
                var anchor = this.document.createElement('img');
                range.collapse(true);
                domUtils.setAttributes(anchor,{
                    'anchorname':name,
                    'class':'anchorclass'
                });
                range.insertNode(anchor).setStartAfter(anchor).setCursor(false,true);
            }
        }
    },
    queryCommandState:function () {
        return this.highlight ? -1 : 0;
    }

};

