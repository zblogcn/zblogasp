<tr>
  <th scope="row"><%=文件注释%></th>
  <td><input name="include_<%=文件名%>" type="file"/></td>
  <td><!--<%=主题调用代码%>--><%=文件名%><input style="float:right" name="copybutton_<%=文件名%>" id="copybutton_<%=文件名%>" value="复制" type="button" bindtag="&lt;#ZC_BLOG_HOST#&gt;zb_users/theme/<%=主题名%>/include/<%=文件名%>" onclick="copydata(this)"/></td>

</tr>
