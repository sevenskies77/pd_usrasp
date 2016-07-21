<%
'Autocalculate new DocID field value on the server side
'
'Variables:
'S_DocID - DocID field to be calculated
'
'S_DocIDParent - parent DocID
'S_Department - department
'S_ClassDoc - document category
'
'Possible functions:
'GenNewDocIDIncrement(sTemplate) - Increments string template by 1
' 	Parameters:
'	sTemplate - DocID string template
'
'Example:
'S_DocID=GenNewDocIDIncrement("D02N05S") - returns "D02N06S"
'
'nRecClassDoc(S_ClassDoc) - returns number of records (documents) having S_ClassDoc document category
'
'nRecDependants(S_DocID) - returns number of records (documents) having S_DocID document parent ID 
'
'LeadSymbolNVal(sVal, sSymbol, N) - returns string value sVal added to the size N with the symbols sSymbol
'
'Example:
'sVal=LeadSymbolNVal("5", "0", 3) - returns "005"
'
'LastEnteredDocID_ClassDoc(S_ClassDoc) - returns last entered document ID for defined document category
'
'Example:
'sVal=LastEnteredDocID_ClassDoc(S_ClassDoc)
'
'LastEnteredDocID_Dependent(S_DocIDParent) - returns last entered dependent document ID for defined document ID
'
'Example:
'LastEnteredDocID_Dependent(S_DocIDParent)
'

'S_DocID=GenNewDocIDIncrement(S_DocID)
'S_DocIDAdd=GenNewDocIDIncrement(S_DocIDAdd)

If IsHelpDeskDoc() Then
	S_DocID=GenNewDocIDIncrement(LastEnteredDocID_ClassDoc(S_ClassDoc))
End If

%>
