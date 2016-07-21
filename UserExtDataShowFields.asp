<%
Sub UserExtDataShowFields
'Show directory fields in the specified order
'Variables:
'FieldToEdit - field name from which this directory (external data) were called
'ExtGUID - editable external data source GUID
'ExtDirGUID - directory data source GUID
'CurrentDocID - document ID from which directory were called
'CurrentClassDoc - document category from which directory were called
'
'CurrentDirectoryFields - array of directory fields to be shown in the specified order
'CurrentDirectoryFieldsCanBeInserted - array of True/False values corresponding to array of directory fields to be shown. 
'Indicates directory fields that can be inserted into the document field.
'Array CurrentDirectoryFieldsCanBeInserted must contain the same number of valus as array CurrentDirectoryFields
'
'Select Case CurrentClassDoc
    'Case "Invoices" ' - document category to be processed
		Select Case ExtDirGUID ' external data recordset to be processed
    		Case "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}" 'external data GUID to be processed
				Select Case FieldToEdit 
   		 			Case "_IntegerField" 'Field "_MoneyField" to be processed

							CurrentDirectoryFields = 			  Array("MemoField", "IntegerField")
							CurrentDirectoryFieldsCanBeInserted=Array( True,        False) 'MemoField can be inserted from the directory, IntegerField can NOT be inserted

    				'Case "YYY" 'Field "YYY" to be processed
	    				'....
				End Select
		    Case "{7DFDF05F-21E5-4037-A313-3EADCE73B0CB}" ' Ph - 20081117 - Справочник проектов
				CurrentDirectoryFields = Array("ProjectID", "ProjectCode", "ProjectName", "Manager", "CostCenter")
				CurrentDirectoryFieldsCanBeInserted = Array(True, False, False, False, False) 'Выбирать можно только первое поле
'				CurrentDirectoryFields = Array("ProjectID", "ProjectCode", "ProjectName")
'				CurrentDirectoryFieldsCanBeInserted = Array(True, False, False) 'Выбирать можно только первое поле
		    Case "{F2DB7E4F-4A57-4D62-BFE4-2CC7B6BB2E55}" ' Ph - 20090317 - Справочник списков пользователей RU
				CurrentDirectoryFields = Array("Description", "UsersList")
				CurrentDirectoryFieldsCanBeInserted = Array(False, True) 'Выбирать можно только второе поле
		    Case "{050FEFFC-1E8B-44EA-A276-9AB49C08F1A4}" ' Ph - 20090317 - Справочник списков пользователей EN
				CurrentDirectoryFields = Array("Description", "UsersList")
				CurrentDirectoryFieldsCanBeInserted = Array(False, True)
		    Case "{3BA70F3A-09A3-46CE-98D2-B1767B9C747C}" ' Ph - 20090317 - Справочник списков пользователей CZ
				CurrentDirectoryFields = Array("Description", "UsersList")
				CurrentDirectoryFieldsCanBeInserted = Array(False, True)
		    'Case "???" ' - other directory GUID to be processed
    			'....
		End Select
    'Case "???" ' - other document category to be processed
    		'....
'End Select
End Sub
%>
