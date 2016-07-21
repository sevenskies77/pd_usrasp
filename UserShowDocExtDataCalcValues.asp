<%
'Calculate some user values during external data output
'
'sExtGUID - output data source GUID
'ds current data recordset, ds("MyField") - field MyField value
'
'Use the following functions:
'
'UserShowDocExtDataCalcValues UserVar, CalcExprBeforeLoop, CalcExprInsideLoop, CalcExprAfterLoop
'
'Example:
If sExtGUID="{60586E06-E3F3-4214-9877-8422A001C165}" or sExtGUID="{60586E06-E3F3-4214-9877-8422A001C167}" Then
'to calculate total value for field MoneyField and to store the total value in rMoneyField variable
	UserShowDocExtDataCalcValues rMoneyField, 0, "rMoneyField+ds(""MoneyField"")", ""
'to calculate super total value for field MoneyField and to store the total value in rMoneyFieldTotal variable
	UserShowDocExtDataCalcValues rMoneyFieldTotal, "", "", "rMoneyFieldTotal+rMoneyField"
'or - to calculate average value
'	UserShowDocExtDataCalcValues rMoneyField, 0, "rMoneyField+ds(""MoneyField"")", rMoneyField/ds.RecordCount
End If
'

If sExtGUID="{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}" Then
'to calculate total value for field Количество and to store the total value in rKol variable
	UserShowDocExtDataCalcValues rKol, 0, "rKol+ds(""Количество"")", ""
	UserShowDocExtDataCalcValues rSum, 0, "rSum+ds(""Сумма"")", ""
	UserShowDocExtDataCalcValues rSumN, 0, "rSumN+ds(""Сумма налога, руб"")", ""
	UserShowDocExtDataCalcValues rSumV, 0, "rSumV+ds(""Стоимость товаров (работ, услуг), всего с учетом налога, руб"")", ""
End If

'Calculation for Invoices
If sExtGUID="{C3DB86C0-0F73-FFBA-DF8D-9F57ED31707A}" Then
	UserShowDocExtDataCalcValues rQty, 0, "rQty+ds(""Qty"")", ""
	UserShowDocExtDataCalcValues rAmount, 0, "rAmount+ds(""Amount"")", ""
	UserShowDocExtDataCalcValues rTax, 0, "rTax+ds(""Tax"")", ""
	UserShowDocExtDataCalcValues rAmountPlusTax, 0, "rAmountPlusTax+ds(""AmountPlusTax"")", ""
End If

'Calculation for Credit Card expenses
If sExtGUID="{A6CA53D4-6AB6-0D48-D543-971BD07A16B3}" Then
	UserShowDocExtDataCalcValues rAmountCC, 0, "rAmountCC+ds(""Amount"")", ""
'AddLogD "rAmount:"+CStr(rAmountCC)
End If

%>
