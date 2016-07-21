<%
Sub UserMSWordBookmarkProcessing (oCurrentRecordset)
'
'Place here ASP code for your MS Word bookmark processing
'This subroutine is called for every document recordset, for every recordset record and for every record field 
'Use CASE statements to process all the values to be output
'Use the following variables in your code:
'oCurrentRecordset		- current recordset object (for example, use oCurrentRecordset.EOF expression to check end of recordset)
'CurrentClassDoc 			- current document type (ClassDoc field of the document record)
'CurrentDocID 			- current document ID (DocID field of the document record)
'CurrentRSNumber 			- current recordset number (starting at 1), 9999 - recordset number for dependent document 
'CurrentRSRecordNumber	- current recordset record number (starting at 1)
'CurrentRSRecordCount	- number of records in the current recordset 
'CurrentRSFieldName		- current recordset field name
'CurrentRSFieldValue		- current recordset field value
'nDataSources				- number of recordsets 
'
'Use the following subroutines in your code:
'
'MSWordInsertRowInTable iTable - inserts row in MS Word table
'Parameters: iTable	- table number in the MS Word document, 0 for the last table
'The table has to be already created 
'
'MSWordInsertBookmarkText Text, BookmarkName - inserts text after the bookmark in the current MS Word document 
'Parameters: BookmarkName - bookmark name
'			   Text - text to be inserted
'
'MSWordAddBookmarksToTable Bookmarks, iTable, iRow, iColStart - adds/moves bookmarks to the specified table in the MS Word document
'Parameters: Bookmarks - array of bookmark names
'			   iTable	- table number in the MS Word document, 0 for the last table
'			   iRow	- row number in the table, 0 for the last row
'			   iColStart	- starting column number in the table
'
'The table has to be already created
'
'Use AddLogD sub to output some debug values to debug log
'AddLogD "CurrentClassDoc: "+CurrentClassDoc
'AddLogD "CurrentRSNumber: "+CStr(CurrentRSNumber)
'AddLogD "CurrentRSRecordCount: "+CStr(CurrentRSRecordCount)
'AddLogD "CurrentRSRecordNumber: "+CStr(CurrentRSRecordNumber)
'AddLogD "CurrentRSFieldName: "+CurrentRSFieldName
'AddLogD "CurrentRSFieldName: "+CurrentRSFieldName

'Example:

'MSWordInsertRowInTable 4

Select Case CurrentClassDoc
	'Invoices outgoing MS Word bookmark processing
    Case "Accounts Receivable" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1 'Recordset N0.1 to be processed - Invoice - shipping/billing details
				Select Case CurrentRSFieldName
    				Case "ShipToName" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToName"
    				Case "ShipToAddress" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToAddress"
    				Case "ShipToCity" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToCity"
    				Case "ShipToState" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToState"
    				Case "ShipToZipCode" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToZipCode"
    				Case "ShipToPhone" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToPhone"
    				Case "BillToName" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "BillToName"
    				Case "BillToAddress" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "BillToAddress"
    				Case "BillToCity" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "BillToCity"
    				Case "BillToState" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "BillToState"
    				Case "BillToZipCode" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "BillToZipCode"
    				Case "BillToPhone" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "BillToPhone"
    				Case "ShipToAddress" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToAddress"
    				Case "ShipToAddress" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToAddress"
    				Case "ShipToAddress" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShipToAddress"
    				Case "Shipping method" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ShippingMethod"
			    			
				End Select
    		Case 2  'Recordset N0.2 -  other recordset to be processed - Invoice - positions
				Select Case CurrentRSFieldName
    				Case "ItemCode" ' - field to be processed
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Or CurrentRSRecordNumber=1 Then
				    			MSWordAddBookmarksToTable Array("ItemCode","ItemNumber","ItemDescription","Qty","Price","Amount","Comment"), 1, 0, 1
			    			End If
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemNumber" ' - field to be processed
			    			MSWordInsertBookmarkText CStr(CurrentRSFieldValue), CurrentRSFieldName
    				Case "ItemDescription" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "Qty" ' - field to be processed
			    			MSWordInsertBookmarkText CStr(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber=1 Then
			    				Session("rItems")=0
			    			End If
			    			Session("rItems")=Session("rItems")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText MyCStr(Session("rItems")), "Items"
			    			End If
    				Case "Price" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), CurrentRSFieldName
    				Case "Amount" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber=1 Then
			    				Session("rSubTotal")=0
			    			End If
			    			Session("rSubTotal")=Session("rSubTotal")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText MyFormatCurrency(Session("rSubTotal")), "SubTotal"
			    			End If
    				Case "Tax" ' - field to be processed
			    			If CurrentRSRecordNumber=1 Then
			    				Session("rTax")=0
			    			End If
			    			Session("rTax")=Session("rTax")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText MyFormatCurrency(Session("rTax")), "TotalTax"
				    			MSWordInsertBookmarkText MyFormatCurrency(Session("rSubTotal")+Session("rTax")), "Total"
			    			End If
    				Case "Comment" ' - field to be processed
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
				    			MSWordInsertRowInTable 1
			    			End If
				End Select
		End Select
    Case "Expense Reports / Credit Card Expenses" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1  'Recordset N0.1 -  other recordset to be processed
				Select Case CurrentRSFieldName
    				Case "CCNumber" ' - field to be processed
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Or CurrentRSRecordNumber=1 Then
				    			MSWordAddBookmarksToTable Array("CCNumber","ItemsCharged","Explanation","DateTime","Amount"), 1, 0, 1
			    			End If
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemsCharged" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ItemsCharged"
    				Case "Explanation" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "DateTime" ' - field to be processed
			    			MSWordInsertBookmarkText MyDate(CurrentRSFieldValue), CurrentRSFieldName
    				Case "Amount" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber=1 Then
			    				Session("rTotal")=0
			    			End If
			    			Session("rTotal")=Session("rTotal")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText MyFormatCurrency(Session("rTotal")), "Total"
			    			End If
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Then
				    			MSWordInsertRowInTable 1
			    			End If
				End Select
		End Select

    Case "Inventory Acknowledgements" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1  'Recordset N0.1 -  other recordset to be processed  - Invoice - positions
				Select Case CurrentRSFieldName
    				Case "ItemNumber" ' - field to be processed
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Or CurrentRSRecordNumber=1 Then
				    			MSWordAddBookmarksToTable Array("ItemNumber","ItemCode","ItemDescription","Price","Qty"), 1, 0, 1
			    			End If
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemCode" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemDescription" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "Qty" ' - field to be processed
			    			MSWordInsertBookmarkText CStr(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber=1 Then
			    				Session("nTotalItems")=0
			    			End If
			    			Session("nTotalItems")=Session("nTotalItems")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText CStr(Session("nTotalItems")), "TotalItems"
			    			End If
    				Case "Price" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Then
				    			MSWordInsertRowInTable 1
			    			End If
				End Select
		End Select

    Case "Parts Orders (Outgoing)" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1  'Recordset N0.1 -  other recordset to be processed  - Invoice - positions
				Select Case CurrentRSFieldName
    				Case "ItemNumber" ' - field to be processed
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Or CurrentRSRecordNumber=1 Then
				    			MSWordAddBookmarksToTable Array("ItemNumber","ItemCode","ItemDescription","Qty", "Comment"), 1, 0, 1
			    			End If
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemCode" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemDescription" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "Qty" ' - field to be processed
			    			MSWordInsertBookmarkText CStr(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber=1 Then
			    				Session("nTotalItems")=0
			    			End If
			    			Session("nTotalItems")=Session("nTotalItems")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText CStr(Session("nTotalItems")), "TotalItems"
			    			End If
    				Case "Comment" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Then
				    			MSWordInsertRowInTable 1
			    			End If
				End Select
		End Select

    Case "Parts Orders" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1  'Recordset N0.1 -  other recordset to be processed  - Invoice - positions
				Select Case CurrentRSFieldName
    				Case "ItemNumber" ' - field to be processed
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Or CurrentRSRecordNumber=1 Then
				    			MSWordAddBookmarksToTable Array("ItemNumber","ItemCode","ItemDescription","Qty", "Comment"), 1, 0, 1
			    			End If
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemCode" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "ItemDescription" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
    				Case "Qty" ' - field to be processed
			    			MSWordInsertBookmarkText CStr(CurrentRSFieldValue), CurrentRSFieldName
			    			If CurrentRSRecordNumber=1 Then
			    				Session("nTotalItems")=0
			    			End If
			    			Session("nTotalItems")=Session("nTotalItems")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText CStr(Session("nTotalItems")), "TotalItems"
			    			End If
    				Case "Comment" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, CurrentRSFieldName
			    			If CurrentRSRecordNumber<>CurrentRSRecordCount Then
				    			MSWordInsertRowInTable 1
			    			End If
				End Select
		End Select

    'Case "???" ' - other document type to be processed
    		'....
    Case "Счета-фактуры" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1 'Recordset N0.1 to be processed
				Select Case CurrentRSFieldName
    				Case "Наименование товара" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Naimenovanie"
    				Case "Единица измерения" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "EdIzm"
    				Case "Количество" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Kol"
    				Case "Цена (тариф за ед изм) руб" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), "Cena"
    				Case "Сумма" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), "Sum"
    				Case "В тч акциз" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), "Akciz"
    				Case "Налоговая ставка" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), "NalogStavka"
    				Case "Сумма налога, руб" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), "SumNalog"
    				Case "Стоимость товаров (работ, услуг), всего с учетом налога, руб" ' - field to be processed
			    			MSWordInsertBookmarkText MyFormatCurrency(CurrentRSFieldValue), "Stoimost"
			    			If CurrentRSRecordNumber=1 Then
			    				Session("rTotalSum")=0
			    			End If
			    			Session("rTotalSum")=Session("rTotalSum")+CurrentRSFieldValue
			    			If CurrentRSRecordNumber=CurrentRSRecordCount Then
				    			MSWordInsertBookmarkText MyFormatCurrency(Session("rTotalSum")), "Total"
				    			MSWordInsertBookmarkText oPayDox.AmountByWords(Session("rTotalSum"), "RUR"), "TotalWords"
			    			End If
    				Case "Страна происхождения" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Strana"
    				Case "Номер грузовой таможенной декларации" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Tamoz"
			    			MSWordInsertRowInTable 2
			    			MSWordAddBookmarksToTable Array("Naimenovanie","EdIzm","Kol","Cena","Sum","Akciz","NalogStavka","SumNalog","Stoimost","Strana","Tamoz"), 2, 0, 1
    				'Case "Номер строки" ' - field to be processed
				End Select
    		Case 2 'Recordset N0.2 to be processed
				Select Case CurrentRSFieldName
    				Case "Продавец" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Prodavec"
    				Case "Адрес" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "AdresProdavca"
    				Case "ИНН продавца" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "INNProdavca"
    				Case "Грузоотправитель" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Gruzootpravitel"
    				Case "Адрес грузоотправителя" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "AdresGruzootpravitela"
    				Case "Грузополучатель" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Gruzopoluchatel"
    				Case "Адрес грузополучателя" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "AdresGruzopoluchatela"
    				Case "К платежно-расчетному документу" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "KPlatRasDoc"
    				Case "Покупатель" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Pokupatel"
    				Case "Адрес покупателя" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "AdresPokupatela"
    				Case "ИНН покупателя" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "INNPokupatela"
    				Case "Руководитель организации (предприятия)" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Rukovoditel"
    				Case "Главный бухгалтер" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "GlavBuh"
    				Case "Выдал (ФИО)" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "FIO"
				End Select
		End Select
    Case "Договора" ' - пример для выдачи списка подчиненных документов в MS Word
		Select Case CurrentRSNumber 
    		Case 1 'Recordset N0.1 to be processed
				Select Case CurrentRSFieldName
    				Case "Наименование товара" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "Naimenovanie"
				End Select
		End Select

    Case "Пропуска / Разовые пропуска для списка" ' - document type to be processed
		Select Case CurrentRSNumber 
    		Case 1 'Recordset N0.1 to be processed
				Select Case CurrentRSFieldName
    				Case "ФИО" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "FIO"
    				Case "Организация" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "ORG"
    				Case "Номер автомашины" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "CarID"
    				Case "Номер строки" ' - field to be processed
			    			MSWordInsertBookmarkText CurrentRSFieldValue, "N"
			    			If CurrentRSRecordNumber < CurrentRSRecordCount Then
			    				MSWordInsertRowInTable 1
			    				MSWordAddBookmarksToTable Array("FIO","ORG","CarID", "N"), 1, 0, 1
			    			End If
				End Select
		End Select
End Select

End Sub
%>