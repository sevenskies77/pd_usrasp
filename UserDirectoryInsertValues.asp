<%
'Update more than one field by one click from directory
'ExtGUID - editable data source GUID (optional)
'ExtDirGUID - directory data source GUID
'FieldToEdit - field name from which this directory were called 
'FieldArrayToEdit - array of field names that have to be updated simultaniously
'FieldArrayToEditFrom - array of directory field names that will be used for update (Field1, ..., Field6 for built-in user defined directory)
'	Use the expression forma._FieldName.value to provide the value of FieldName field
'bShowValues - True or False - Show inserted values
'
'Use the following functions:
'
'Inserts group of values from fields FieldArrayToEditFrom to the group of fields FieldArrayToEdit when you call the directory ExtDirGUID from the field FieldToEdit of the editable data source ExtGUID
'UserDirectoryInsertValues ExtGUID, ExtDirGUID, FieldToEdit, FieldArrayToEdit, FieldArrayToEditFrom, bShowValues
'Example:
'UserDirectoryInsertValues "{60586E06-E3F3-4214-9877-8422A001C165}", "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}", "_MoneyField", Array("_MemoField", "_DateField", "_IntegerField"), Array("_MemoField", "_DateField", "_IntegerField"), True
'
'Inserts group of values from fields FieldArrayToEditFrom to the group of fields FieldArrayToEdit when you call the directory ExtDirGUID from the editable data source ExtGUID
'UserDirectoryInsertValuesGroup ExtGUID, ExtDirGUID, FieldArrayToEdit, FieldArrayToEditFrom, bShowValues
'Example (external data directory):
'UserDirectoryInsertValuesGroup "{60586E06-E3F3-4214-9877-8422A001C165}", "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}", Array("_MemoField", "_DateField", "_IntegerField"), Array("_MemoField", "_DateField", "_IntegerField"), True
'
'Example 1 (using single field name to insert):
'UserDirectoryInsertValuesGroup "", "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}", Array("_MemoField", "_DateField"), Array("_MemoField", "_DateField"), True
'
'Example 2 (using multiple field name expression to insert):
'UserDirectoryInsertValuesGroup "", "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}", Array("_MemoField", "_DateField"), Array("'Content: '+forma._MemoField.value+', Date:'+forma._DateField.value", "_DateField"), True
'
'                               Счет-фактура - реквизиты						Контрагенты *
UserDirectoryInsertValuesGroup "{9FF4E1AF-E9FA-B62B-4098-B4184014F549}", "{8F8891BF-5944-A0BB-58DB-56BEA11D24E5}", Array("_Prodavec", "_Adres", "_INN_prodavca"), Array("forma._Name.value", "forma._PostAddress.value", "forma._TaxID.value"), True
UserDirectoryInsertValuesGroup "{9FF4E1AF-E9FA-B62B-4098-B4184014F549}", "{8F8891BF-5944-A0BB-58DB-56BEA11D24E5}", Array("_Gruzootpravitel", "_Adres_gruzootpravitelja"), Array("forma._Name.value", "forma._PostAddress.value"), True
UserDirectoryInsertValuesGroup "{9FF4E1AF-E9FA-B62B-4098-B4184014F549}", "{8F8891BF-5944-A0BB-58DB-56BEA11D24E5}", Array("_Gruzopoluchatel", "_Adres_gruzopoluchatelja"), Array("forma._Name.value", "forma._PostAddress.value"), True
UserDirectoryInsertValuesGroup "{9FF4E1AF-E9FA-B62B-4098-B4184014F549}", "{8F8891BF-5944-A0BB-58DB-56BEA11D24E5}", Array("_Pokupatel", "_Adres_pokupatelja", "_INN_pokupatelja"), Array("forma._Name.value", "forma._PostAddress.value", "forma._TaxID.value"), True

'Пример одновременной вставки реквизитов в поля карточки д-та
'UserDirectoryInsertValuesGroup "", "{8F8891BF-5944-A0BB-58DB-56BEA11D24E5}", Array("DocPartnerName", "UserFieldText1", "UserFieldText2", "UserFieldText3"), Array("forma._Name.value","forma._ManagerPosition.value", "forma._ManagerName.value", "forma._Fax.value"), True

'Для пропусков
'UserDirectoryInsertValuesGroup "", "C", Array("UserFieldText1", "UserFieldText2"), Array("forma._Name.value","forma._ManagerPosition.value", "forma._ManagerName.value", "forma._Fax.value"), True
UserDirectoryInsertValuesGroup "", "C", Array("UserFieldText1", "UserFieldText2"), Array("forma._Name.value","forma._PartnerName.value"), True

'Для SAP R/3 HelpDesk
UserDirectoryInsertValuesGroup "", "{172166E7-8AAE-05DD-3E0E-80EA7228ECD3}", Array("UserFieldText1", "UserFieldText2", "UserFieldText3", "UserFieldText4", "DocRank"), Array("forma._Processes.value","forma._Task.value","forma._Transaction.value","forma._Modules.value","forma._Rank.value"), False
%>
