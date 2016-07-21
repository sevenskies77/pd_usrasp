<%
'Autocalculate Doc field values on the server side
'FieldToCalculate - field name which has to be calculated
'Formula - VBScript-formula valid on the server side
'CurrentClassDoc - current document category 
'
'Possible subroutines:
'UserCalculateValuesDoc FieldToCalculate, Formula
' 	where:
'	FieldToCalculate - existing field name of Docs table, for example, "AmountDoc", "QuantityDoc" etc.
'	!Do not use DocID field name in this subroutine!
'
'Example:
'
'If CurrentClassDoc="Интеграция с бухгалтерской системой" Then
'If CurrentClassDoc="Invoices" Then
	'UserCalculateValuesDoc "AmountDoc", rMoneyFieldTotal
'End If

'Calculation for "Счета-фактуры"
'If CurrentClassDoc="Счета-фактуры" Then
'	UserCalculateValuesDoc "AmountDoc", rSumV
'	UserCalculateValuesDoc "QuantityDoc", rKol
'End If
'
'Calculation for Invoices
'If CurrentClassDoc="Invoices outgoing" Then
'	UserCalculateValuesDoc "AmountDoc", rAmountPlusTax
'	UserCalculateValuesDoc "QuantityDoc", rQty
'End If
'
'Calculation for Credit Card expenses
'If CurrentClassDoc="Expense reports / Credit Card expenses" Then
'	UserCalculateValuesDoc "AmountDoc", rAmountCC
'End If

%>
