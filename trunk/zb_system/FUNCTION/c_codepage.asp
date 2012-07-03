<%
Sub ErrorHandle
On Error Resume Next
Response.CodePage=65001
Err.Clear
End Sub
Call ErrorHandle
%>