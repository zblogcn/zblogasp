$(document).ready(function() {
    if (GetCookie("username") && GetCookie("password")) {
        var a,
        c,
        b,
        d = 0;
        if ($('p[codetype=post]').length > 0) {
            d = $('p[codetype=post]').attr('entryid')
        }
        $('p.cloudreamHelperLink').css('display', 'block').css('text-align', 'right').each(function() {
            a = $(this);
            c = a.attr('codetype');
            b = a.attr('entryid');
            if (c == "comment") {
                a.append('[<a class="helperLink" href="' + bloghost + 'zb_system/cmd.asp?act=CommentEdt&amp;revid=0&amp;id=' + b + '&amp;log_id=' + d + '">编辑</a>]&nbsp;[<a class="helperGet" href="' + bloghost + 'zb_system/cmd.asp?act=CommentDel&amp;id=' + b + '&amp;log_id=' + d + '">删除</a>]')
            } else if (c == "postmulti" || c == "post") {
                a.append('[<a class="helperLink" href="' + bloghost + 'zb_system/cmd.asp?act=ArticleEdt&amp;webedit=ueditor&amp;id=' + b + '">编辑</a>]')
            }
        });
        $('a.helperLink').attr("target", "_0");
        $('a.helperGet').click(function(e) {
            if (window.confirm('确认删除？')) {
                $this = $(this);
                $href = $this.attr("href");
                $this = $this.parent().html("[删除中...]");
                $.ajax({
                    url: $href,
                    dataType: "text",
                    complete: function() {
                        $this.html("[已删除]").parent().css("opacity", "0.5")
                    }
                })
            }
            e.preventDefault()
        })
    }
});