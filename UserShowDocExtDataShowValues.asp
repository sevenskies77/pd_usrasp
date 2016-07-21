<%
'Show some calculated user values after external data output
'
'sExtGUID - output data source GUID
'ds current data recordset, ds("MyField") - field MyField value
'
'Use the following functions:
'
'UserShowDocExtDataShowValues UserExpression, sExtDataFieldName
'
'Example:
If MyCStr(ExtDirGUID)="{60586E06-E3F3-4214-9877-8422A001C165}" Then
'to show the expression "Total: "+MyFormatCurrency(rMoneyField) in the MoneyField field table column 
	UserShowDocExtDataShowValues "Total: "+MyFormatCurrency(rMoneyField), "MoneyField"
'to show the expression "<br>SuperTotal: "+MyFormatCurrency(rMoneyFieldTotal) in the MoneyField field table column 
	UserShowDocExtDataShowValues "<br>SuperTotal: "+MyFormatCurrency(rMoneyFieldTotal), "MoneyField"
End If

If sExtGUID="{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}" Then
	UserShowDocExtDataShowValues "<b>"+MyFormatCurrency(rKol)+"</b>", "Количество"
	UserShowDocExtDataShowValues "<b>"+MyFormatCurrency(rSum)+"</b>", "Сумма"
	UserShowDocExtDataShowValues "<b>"+MyFormatCurrency(rSumN)+"</b>", "Сумма налога, руб"
	UserShowDocExtDataShowValues "<b>Итого: "+MyFormatCurrency(rSumV)+"</b>", "Стоимость товаров (работ, услуг), всего с учетом налога, руб"
End If

'Calculation for Invoices
If sExtGUID="{C3DB86C0-0F73-FFBA-DF8D-9F57ED31707A}" Then
	UserShowDocExtDataShowValues "<b>"+MyCStr(rQty)+"</b>", "Qty"
	UserShowDocExtDataShowValues "<b>"+MyFormatCurrency(rAmount)+"</b>", "Amount"
	UserShowDocExtDataShowValues "<b>"+MyFormatCurrency(rTax)+"</b>", "Tax"
	UserShowDocExtDataShowValues "<b>Total: "+MyFormatCurrency(rAmountPlusTax)+"</b>", "AmountPlusTax"
End If

'Calculation for Credit Card expenses
If sExtGUID="{A6CA53D4-6AB6-0D48-D543-971BD07A16B3}" Then
	UserShowDocExtDataShowValues "<b>"+MyFormatCurrency(rAmountCC)+"</b>", "Amount"
End If

%>
