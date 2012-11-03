<tr>
  <th scope="row"><%=文件注释%></th>
  <td><textarea name="include_<%=文件名%>" style="width:98%"><%=LoadFromFile(BlogPath & "zb_users\theme\<%=主题名%>\include\<%=文件名%>","utf-8")%></textarea></td>
  <td><%=主题调用代码%><input style="float:right" name="copybutton_<%=文件名%>" id="copybutton_<%=文件名%>" value="复制" type="button" bindtag="<%=主题调用代码%>" onclick="copydata(this)"/></td>
</tr>
