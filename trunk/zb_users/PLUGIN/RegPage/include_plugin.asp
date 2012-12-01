<%
Dim Action_Plugin_RegPage_Begin()
ReDim Action_Plugin_RegPage_Begin(0)
Dim bAction_Plugin_RegPage_Begin
Dim sAction_Plugin_RegPage_Begin

Dim Action_Plugin_RegPage_End()
ReDim Action_Plugin_RegPage_End(0)
Dim bAction_Plugin_RegPage_End
Dim sAction_Plugin_RegPage_End

Dim Action_Plugin_RegSave_End()
ReDim Action_Plugin_RegSave_End(0)
Dim bAction_Plugin_RegSave_End
Dim sAction_Plugin_RegSave_End

Dim Action_Plugin_RegSave_Begin()
ReDim Action_Plugin_RegSave_Begin(0)
Dim bAction_Plugin_RegSave_Begin
Dim sAction_Plugin_RegSave_Begin

Dim Action_Plugin_RegSave_Register
ReDim Action_Plugin_RegSave_Register(0)
Dim sAction_Plugin_RegSave_Register
Dim bAction_Plugin_RegSave_Register

Dim Response_Plugin_RegPage_End
Response_Plugin_RegPage_End=""

Dim Response_Plugin_RegPage_Begin
Response_Plugin_RegPage_Begin=""

Dim sFilter_Plugin_RegPage_Vaild
Function Filter_Plugin_RegPage_Vaild(ByRef Username,ByRef Password,ByRef EMail,ByRef HomePage)

	Dim s,i

	If sFilter_Plugin_RegPage_Vaild="" Then Exit Function

	s=Split(sFilter_Plugin_RegPage_Vaild,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "Username,Password,EMail,HomePage")
	Next

End Function
%>