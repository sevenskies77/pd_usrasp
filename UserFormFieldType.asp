<%
Function UserFormFieldType ()
'
'Place here ASP code to define your form field type
'Possible form field types: check box, radio button, drop-down menu 
'This function is called in ChangeExtData.asp page and lets you to define your form field type for editable form fields 
'
'Use CASE statements to process all the values to be output
'Variables:
'ExtGUID	- editable data source GUID
'FieldName	- editable form field name
'
'Use the following functions in your code:
'
'PutFormFieldSelect(ArrayOfChoices, ArrayOfValues) - outputs DROP-DOWN MENU form field
'Parameters: ArrayOfChoices - array of choice options
'			   ArrayOfValues - array of choice values
'
'PutFormFieldRadio(ArrayOfChoices, ArrayOfValues) - outputs RADIO BUTTON form field
'Parameters: ArrayOfChoices - array of choice options
'			   ArrayOfValues - array of choice values
'
'PutFormFieldCheckbox (Value, Title) - outputs CHECK BOX form field
'Parameters: Value - choice value
'				Title - title text
'
AddLogD "ExtGUID:"+ExtGUID
UserFormFieldType=False
Select Case ExtGUID
	    Case "{60586E06-E3F3-4214-9877-8422A001C165}" ' - data source GUID to be processed
	    'Case "{60586E06-E3F3-4214-9877-8422A001C167}" ' - data source GUID to be processed
			Select Case FieldName 
    			Case "Sel" 'FieldName "Sel" to be processed
    				'Outputs field "Sel" as drop-down menu
					UserFormFieldType=PutFormFieldSelect(Array("Choice1", "Choice2", "Choice3"), Array("Value1", "Value2", "Value3"))
    			Case "Rad" 'FieldName "Rad" to be processed
    				'Outputs field "Rad" as radio button
					UserFormFieldType=PutFormFieldRadio(Array("Choice1", "Choice2", "Choice3"), Array("Value1", "Value2", "Value3"))
    			Case "Checkbox" 'FieldName "CheckBox" to be processed
    				'Outputs field "Check" as check box
					UserFormFieldType=PutFormFieldCheckbox ("ON", "User defined checkbox sample")
			End Select
		Case "???" ' - other data source GUID to be processed
    			'....

	   Case "{5EB5017F-334A-6303-BC09-BADCECCF8657}" ' - Invoice - shipping/billing details
   			Select Case FieldName 
    			Case "BillingSameAsShipping" 'FieldName "BillingSameAsShipping" to be processed
    				'Outputs field "BillingSameAsShipping" as drop-down menu
					UserFormFieldType=PutFormFieldSelect(Array("Yes", "No"), Array("YES", "NO"))
			End Select
End Select
End Function


%>