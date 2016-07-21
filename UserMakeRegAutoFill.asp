
<%
'Autofill registration log field values on the client side
'
'Possible functions:
'CanGetFromDoc(FieldName) Is FieldName field name accessible for this registration log from the document record 
'	FieldName - registration log field name the same as the document record field name 
'
'
'Use the expression window.opener.RegForm.FieldName.value to provide the value of FieldName field from the document record
'Use the expression document.forms[0].FieldName.value to provide the value of FieldName field from in registration log
'Example:
'
If CanGetFromDoc("DateReg") Then
%>
//document.forms[0].DateReg.value=window.opener.RegForm.DateReg.value; 
//document.forms[0].DateReg.value='test'; 
<%
End If
%>

