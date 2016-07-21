<%
Function UserExtDataCheck ()
'Place here ASP code for your external data validation
'Use expression Request("FieldName") to get the value of the FieldName field
'.......................
UserExtDataCheck=True 	'If everything is OK
'UserExtDataCheck=False 'If something wrong
'Session("Message")="Невозможно удалить/отредактировать данные, тк недостаточно прав!!"
End Function
%>