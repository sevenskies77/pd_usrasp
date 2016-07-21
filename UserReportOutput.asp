<%
'Define some user preport outputs

'Possible variables:
'sReportGUID - Report GUID
'
'AddFieldSumOrder1(1),...,AddFieldSumOrder1(32) - summary values of additional calculated fields 1,...,32 for order 1
'AddFieldSumOrder2(1),...,AddFieldSumOrder2(32) - summary values of additional calculated fields 1,...,32 for order 2
'AddFieldSumTotal(1),...,AddFieldSumTotal(32) - total summary values of additional calculated fields 1,...,32
'
'sReportStage - use this variable to check the currenr report running stage, 
'possible variable values: 
'"BEGIN" 		- begin of the report
'"LOOP" 		- inside the report loop, use here the expression dsDoc("<FieldName>") to get the field value of the field name <FieldName> of the table Docs
'"SUMORDER1" 	- inside the output for the order 1
'"SUMORDER2" 	- inside the output for the order 2
'"SUMTOTAL" 	- inside the output for the total sum
'Example:

If sReportGUID="{041D872E-E4ED-66AA-2EF5-74B08D9D9831}" Then 'Contracts
Select Case sReportStage
    Case "BEGIN"
		'User variables initialization
		MyVar1=0
		MyVar2=0
    Case "LOOP"
		'Calculace user variables 
		If dsDoc("ClassDoc")="Invoices" Then
			MyVar1=MyVar1+dsDoc("AmountDoc") 'Calculate summary invoices amount
		End If
		If dsDoc("ClassDoc")<>"Invoices" Then
			MyVar2=MyVar2+dsDoc("AmountDoc") 'Calculate summary amount without invoices 
		End If
    Case "SUMORDER1"
    Case "SUMORDER2"
    Case "SUMTOTAL"
		Response.Write "Total invoices amount:"+MyFormatCurrency(MyVar1)
		Response.Write "<br>Total amount without invoices :"+MyFormatCurrency(MyVar2)
End Select
End If	

'Ph- 20081117 - Убрать нумерацию колонок
bColCounter = False

'Ограничение на число выводимых записей в отчете TOP_20_longest_approvals_list делаем фильтром
If sReportGUID="{66B2774B-07A1-4CEF-895E-BAED8351D709}" Then
  If sReportStage = "LOOP" Then
    If nPrinted = 20 Then
      sShowThisRecord = "X" 'Завершить цикл по документам
    End If
  End If
End If

'Запрос №23 - СТС - start
'Отсечечение из лога использования правил строк, относящихся к предыдущим сохранениям документа. Остаются только строки, отвечающие за последнее сохранение (по дате создания)
If sReportGUID = "{DE37E6E8-58E7-468E-A635-1F3BE63BBE9C}" Then
  If sReportStage = "BEGIN" Then
    SaveDateCreation = VAR_BeginOfTimes
  End If
  If InStr(sReportStage, "LOOP") = 1 Then
    If SaveDateCreation <> VAR_BeginOfTimes and dsDoc("DateCreation") <> SaveDateCreation Then
      sShowThisRecord = "X" 'Завершить цикл, т.е. показываем только записи с одной датой (последний блок)
	Else
	  SaveDateCreation = dsDoc("DateCreation")
    End If
  End If
End If
'Запрос №23 - СТС - end
'{ph - 20120610
'Documents in the Paydox archive
'Переключение проверки доступа на архивную БД
If sReportGUID = "{3D00D517-6264-4A36-A4E5-2AB33FD5EB77}" Then
	If sReportStage = "BEGIN" Then
		sSaveSessionArchive = Session("Archive")
		Session("Archive") = "YES"
	End If
	If sReportStage = "END" Then
		Session("Archive") = sSaveSessionArchive
	End If
End If
'ph - 20120610}

%>
