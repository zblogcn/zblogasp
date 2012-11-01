<tr>
  <th scope="row"><%=templatetag%></th>
  <td><textarea name="include_<%templatetag_name%>" style="width:98%;height:200px"><%=LoadFromFile(BlogPath & "zb_users\theme\<%=templatename%>\include\<%=templatetag_name%>.html","utf-8")%></textarea></td>
</tr>
