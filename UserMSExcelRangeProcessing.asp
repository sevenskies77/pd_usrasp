<%
Sub UserMSExcelRangeProcessing (oCurrentRecordset)
'
'Place here ASP code for your MS Excel range(cell) processing
'This subroutine is called for every document recordset, for every recordset record and for every record field 
'Use CASE statements to process all the values to be output
'Use the following variables in your code:
'oCurrentRecordset		- current recordset object (for example, use oCurrentRecordset.EOF expression to check end of recordset)
'CurrentClassDoc 			- current document type (ClassDoc field of the document record)
'CurrentDocID 			- current document ID (DocID field of the document record)
'CurrentRSNumber 			- current recordset number (starting at 1)
'CurrentRSRecordNumber	- current recordset record number (starting at 1)
'CurrentRSRecordCount	- number of records in the current recordset 
'CurrentRSFieldName		- current recordset field name
'CurrentRSFieldValue		- current recordset field value
'nDataSources				- number of recordsets 
'
'Use the following subroutines in your code:
'
'InsertRangeText Text, sRangeName - inserts text into the range(cell) in the current MS Excel document 
'Parameters: sRangeName - range(cell) name
'			   Text - text to be inserted
'
'Use AddLogD sub to output some debug values to debug log
'AddLogD "CurrentClassDoc: "+CurrentClassDoc
'AddLogD "CurrentRSNumber: "+CStr(CurrentRSNumber)
'AddLogD "CurrentRSRecordCount: "+CStr(CurrentRSRecordCount)
'AddLogD "CurrentRSRecordNumber: "+CStr(CurrentRSRecordNumber)
'AddLogD "CurrentRSFieldName: "+CurrentRSFieldName
'
'Example:

'Select Case CurrentClassDoc
    'Case "Invoices" ' - document type to be processed
		'Select Case CurrentRSNumber 
    		'Case 1 'Recordset N0.1 to be processed
				Select Case CurrentRSFieldName
    				Case "MemoField" ' - field to be processed
    					If CurrentRSRecordNumber=1 Then
			    			InsertRangeText "«Sample text to output together with the field value» "+CurrentRSFieldValue, "SampleRangeName"
			    		End If
    				'Case "???" ' -  other field to be processed
			    		'....
				End Select
    		'Case 2  'Recordset N0.2 -  other recordset to be processed
	    		'....
		'End Select
    'Case "???" ' - other document type to be processed
    		'....
'End Select

End Sub
%>