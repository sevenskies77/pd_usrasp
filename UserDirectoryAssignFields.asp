<%
Sub UserDirectoryAssignFields
'Assign directory fields to the document fields
'Variables:
'ExtGUID - editable data source GUID
'FieldToEdit - field name from which directory were called
'CurrentDocID - document ID from which directory were called
'CurrentClassDoc - document category from which directory were called
'
'CurrentDirectory - assigned directory index in the format:
' U - user list directory
' C - contact names directory
' P - partner list directory
' D - department list directory
' I - inventory unit directory
' A - activity directory
' T - document category directory
' Z - position directory
' R - report types directory
' E - currency rates directory
' K - context marks directory
' M - measure units directory
' ! - directory call prohibited
'
'CurrentDirectoryGUID - user defined or external data source Directory GUID, for example, "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}"
'
'Out "FieldToEdit: "+FieldToEdit 'debug FieldToEdit value
'Select Case CurrentClassDoc
    'Case "Invoices" ' - document category to be processed
		Select Case FieldToEdit 
    		Case "_IntegerField" 'Field "_IntegerField" to be processed
					'CurrentDirectory="U" 'U - user list directory
					'CurrentDirectory="!" 'supress any directory for this field
					'CurrentDirectoryGUID = "{EAB2C1BF-2676-E606-B671-7D7B051A5DC4}"
					'CurrentDirectoryGUID = "{C0EB84E0-1941-6428-55D1-EEFC12E4EA1F}"
    		'Case "YYY" 'Field "YYY" to be processed
	    		'....
		End Select
    'Case "???" ' - other document category to be processed
    		'....
'End Select

'If FieldToEdit="_Edinica_izmerenija" Then
'CurrentDirectoryGUID = "{A4BB9E88-A9B3-209A-F4BE-8CF11FF5CB81}"
'End If
'Out "CurrentDirectory:"+CurrentDirectory
'Out "CurrentDirectoryGUID:"+CurrentDirectoryGUID
'Out FieldToEdit

'If InStr(UCase(CurrentClassDoc), UCase("Пропуска"))>0 Or InStr(UCase(Request("ClassDoc")), UCase("Пропуска"))>0 Then 
'	If FieldToEdit = "UserFieldText1" Or FieldToEdit = "UserFieldText2" Then
'		CurrentDirectory="C"
'	End If
'End If
'If InStr(UCase(CurrentClassDoc), UCase("Договора"))>0 Or InStr(UCase(Request("ClassDoc")), UCase("Договора"))>0 Then 
'	If FieldToEdit = "UserFieldText1" Then
'		CurrentDirectoryGUID = "{D2D11C5C-F9D8-9DE2-AFA1-A431C7E4DFFD}"
'	End If
'End If
'CurrentDirectory="U"

'If FieldToEdit = "DocListToReconcile" Then
'  CurrentDirectoryGUID = "{53D3E531-8DCB-4413-603A-8268443FCBFF}"
'End If

'CurrentDirectoryGUID = "{FBAE2C89-AEB3-411E-6411-E87700C3EF4F}"


' Настройки для Sitronics/STS
' SAY 2008-07-21
' AM 240808

'Определяем бизнес направление
' DmGorsky_5: меняем местами проверку: вначале ищем SIT_SITRU ("СИТРОНИКС ИТ"), потом, если не найдено, - SIT_SITRONICS ("СИТРОНИКС")
If InStr(UCase(Session("Department")), UCase(SIT_SITRU)) = 1 Then ' DmGorsky_3
  sDepartmentRoot = SIT_SITRU ' DmGorsky_3
ElseIf InStr(UCase(Session("Department")), UCase(SIT_STS)) = 1 Then
  sDepartmentRoot = SIT_STS
ElseIf InStr(UCase(Session("Department")), UCase(SIT_SITRONICS)) = 1 Then
  sDepartmentRoot = SIT_SITRONICS
'Запрос №1 - СИБ - start
ElseIf InStr(UCase(Session("Department")), UCase(SIT_SIB_ROOT_DEPARTMENT)) = 1 Then
  sDepartmentRoot = SIT_SIB
ElseIf InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then
  sDepartmentRoot = SIT_RTI
ElseIf InStr(UCase(Session("Department")), UCase(SIT_MINC)) = 1 Then
  sDepartmentRoot = SIT_MINC
'Запрос №1 - СИБ - end
'amw
ElseIf InStr(UCase(Session("Department")), UCase(SIT_MIKRON)) = 1 Then
  sDepartmentRoot = SIT_MIKRON
'amw

Else
  sDepartmentRoot = ""
End If


If InStr(UCase(CurrentClassDoc), UCase(SIT_VHODYASCHIE))>0 Then 
  Select Case FieldToEdit 
    Case "UserFieldText3" 
      CurrentDirectory = "C"
    Case "UserFieldText4" 
      CurrentDirectory = "C"
    Case "UserFieldText5"
      CurrentDirectory="!"
    Case "UserFieldText6"
      CurrentDirectory="!"
    Case "UserFieldText7" ' DmGorsky_3
      If sDepartmentRoot = SIT_SITRU Then ' DmGorsky_3
        CurrentDirectoryGUID = "{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7}" ' DmGorsky_3
      End If ' DmGorsky_3
  End Select
End If

'SAY 2008-10-27
If InStr(UCase(CurrentClassDoc), UCase(SIT_VHODYASCHIE_ACC))>0 Then 
  Select Case FieldToEdit 
    Case "UserFieldText2" 
      Select Case UCase(Request("l"))
        Case "RU" 'RU
          CurrentDirectoryGUID = "{6D42F0FB-1389-4CA0-8A00-9E0CD3F09CF7}"
        Case "" 'EN
          CurrentDirectoryGUID = "{6D57662F-7DD0-41E1-806B-3562412FDFAF}"
        Case "3" 'CZ
          CurrentDirectoryGUID = "{0D620DAB-1B89-4E7B-BB6A-29EB77F9AEE9}"
      End Select
    Case "UserFieldText3"
      CurrentDirectory = "C"
    Case "UserFieldText4" 
      CurrentDirectory = "C"
    Case "UserFieldText5"
      CurrentDirectory="!"
    Case "UserFieldText6"
      CurrentDirectory="!"
'Ph - 20090311 - Start
      Case "UserFieldText7"
  Select Case UCase(Request("l"))
    Case "RU" 'RU
      If sDepartmentRoot = SIT_SITRONICS then
        CurrentDirectoryGUID = "{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7}"
      ElseIf sDepartmentRoot = SIT_STS then
        CurrentDirectoryGUID = "{DA5960BE-A65D-4D21-BF89-73233FFEAEE8}"
      Else 'Другие
        CurrentDirectoryGUID = "{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7}"
      End If
    Case "" 'EN
      If sDepartmentRoot = SIT_SITRONICS then
        CurrentDirectoryGUID = "{FC8A0260-6B28-4F9F-BF2D-6F95DDE21E1C}"
      ElseIf sDepartmentRoot = SIT_STS then
        CurrentDirectoryGUID = "{6A200BD7-1A53-40FC-9DBB-44499F65B74C}" '--- новый справочник
      Else 'Другие
        CurrentDirectoryGUID = "{FC8A0260-6B28-4F9F-BF2D-6F95DDE21E1C}"
      End If
    Case "3" 'CZ
      If sDepartmentRoot = SIT_SITRONICS then
        CurrentDirectoryGUID = "{521C56BD-EC92-4AF5-BE8C-229391C37673}"
      ElseIf sDepartmentRoot = SIT_STS then
        CurrentDirectoryGUID = "{2F4D0C04-FD15-4321-A5E3-5AA2FCB0D70E}" ' --- новый справочник
      Else 'Другие
        CurrentDirectoryGUID = "{521C56BD-EC92-4AF5-BE8C-229391C37673}"
      End If
  End Select
'Ph - 20090311 - End
  End Select
End If

If InStr(UCase(CurrentClassDoc), UCase(SIT_ISHODYASCHIE))>0 Then 
  Select Case FieldToEdit 
    Case "UserFieldText5"
      'CurrentDirectory="P"
      CurrentDirectory="C"
  End Select
End If

If InStr(UCase(CurrentClassDoc), UCase(SIT_NORM_DOCS))>0 Then 
  Select Case FieldToEdit 
	Case "UserFieldText4" 
		CurrentDirectory = "U"
  End Select
End If

'Запрос №11 - СТС - start
If InStr(UCase(CurrentClassDoc), UCase(SIT_DOGOVORI_OLD)) = 1 or InStr(UCase(CurrentClassDoc), UCase(SIT_DOGOVORI_NEW)) = 1 Then 
'If InStr(UCase(CurrentClassDoc), UCase(SIT_DOGOVORI))>0 Then 
'Запрос №11 - СТС - end
  Select Case FieldToEdit 
	Case "UserFieldText8" 
	  CurrentDirectoryGUID = "{2CC714EB-5836-49E9-B873-3A34EDB85098}" 'Список проектов для заявок
'      Select Case UCase(Request("l"))
'        Case "RU" 'RU
'          CurrentDirectoryGUID = "{959D450F-9E5A-4358-B445-1D082041987A}"
'        Case "" 'EN
'          CurrentDirectoryGUID = "{F9A6AADA-7DDD-4776-836A-A3EE4032D957}"
'        Case "3" 'CZ
'          CurrentDirectoryGUID = "{7E9A8B94-3C6E-4597-9B09-FCABD40BB155}"
'      End Select
	Case "UserFieldText1" 
      CurrentDirectory = "P"
  End Select
End If

'vnik_payment_order
If InStr(UCase(CurrentClassDoc), UCase(SIT_PAYMENT_ORDER)) > 0 Then 
'out FieldToEdit
  Select Case FieldToEdit 
	Case "_CFC" 
	  CurrentDirectoryGUID = "{E2D63CEF-6DC4-44F3-B4BF-BD4B0D9438D0}"
	Case "_Percent" 
      CurrentDirectory = ""
  End Select
End If
'vnik_payment_order

'Для согласующих и подписантов всех категорий по умолчанию справочник ролей
If sDepartmentRoot = SIT_SITRONICS Then 'СИТРОНИКС
  If (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        CurrentDirectoryGUID = "{78961D78-DB41-4483-99AF-C36BD0A98701}"
      Case "" 'EN
        CurrentDirectoryGUID = "{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A}"
      Case "3" 'CZ нет, используется английский
        CurrentDirectoryGUID = "{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A}"
    End Select
  End If

ElseIf sDepartmentRoot = SIT_RTI Then     ' Справочник "Роли РТИ"
  If (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
    CurrentDirectoryGUID = "{26A64828-4E82-4398-84E1-11F9F092FDD8}"
    VAR_NotToShowUserDirectories = "Y"
    CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ",{71B3A81A-0A45-41BB-8DAE-A0EB1DA292A6},{3012D11D-199C-4D46-8B58-6704EEF4A3EF},{F70782E3-B1A3-4AFE-B800-905763A24E70},{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A},{78961D78-DB41-4483-99AF-C36BD0A98701},{3482B53C-018A-457A-BA91-BCDBD1EB106A},{95F98947-C5E4-452B-B8B1-238D887CC168},{154E1D8C-B044-4E32-B40B-B6857DD00B5A},{E4BF3B18-ACAE-4E77-8BDF-34CF75481C34},{8FF3157E-1099-4256-A801-51DE178950AF},{9EB619A3-61E8-4C1A-9573-27C87DABEF76},{4AEC67F9-80C1-427F-904E-489920B71940},{9850A686-F991-4F36-8EF2-C0F043103276},{08481F91-3506-40EE-852F-FAB1568E935E},{ECEAF686-3552-44BD-A49B-941376AE4109},{6A8607D5-88A1-4706-87D4-B37D633B2671},{84E1A1BB-0CBB-4258-9017-9D92EEEE2522},{4B95FC42-1E31-4574-93B2-FAC28C5D7C3A}"
  End If

  If (FieldToEdit = "DocListToView") Then
    VAR_NotToShowUserDirectories = "Y"
    CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ",{71B3A81A-0A45-41BB-8DAE-A0EB1DA292A6},{3012D11D-199C-4D46-8B58-6704EEF4A3EF},{F70782E3-B1A3-4AFE-B800-905763A24E70},{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A},{78961D78-DB41-4483-99AF-C36BD0A98701},{3482B53C-018A-457A-BA91-BCDBD1EB106A},{95F98947-C5E4-452B-B8B1-238D887CC168},{154E1D8C-B044-4E32-B40B-B6857DD00B5A},{E4BF3B18-ACAE-4E77-8BDF-34CF75481C34},{8FF3157E-1099-4256-A801-51DE178950AF},{9EB619A3-61E8-4C1A-9573-27C87DABEF76},{4AEC67F9-80C1-427F-904E-489920B71940},{9850A686-F991-4F36-8EF2-C0F043103276},{08481F91-3506-40EE-852F-FAB1568E935E},{ECEAF686-3552-44BD-A49B-941376AE4109},{6A8607D5-88A1-4706-87D4-B37D633B2671},{84E1A1BB-0CBB-4258-9017-9D92EEEE2522},{4B95FC42-1E31-4574-93B2-FAC28C5D7C3A}"
  End If

ElseIf sDepartmentRoot = SIT_MIKRON Then     ' Справочник "Роли МИКРОН"
   If (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
      CurrentDirectoryGUID = MIKRON_CATALOG_ROLES
   End If
   
ElseIf sDepartmentRoot = SIT_MINC Then     ' Справочник пользователи  
   If InStr("#_ListToReconcile#_NameAproval#_NameResponsible#_Correspondent#_ListToView#_Registrar#_NameControl#_ListToEdit#", "#" & FieldToEdit & "#") Then
      CurrentDirectory = "U"
      CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ",{71B3A81A-0A45-41BB-8DAE-A0EB1DA292A6},{3012D11D-199C-4D46-8B58-6704EEF4A3EF},{F70782E3-B1A3-4AFE-B800-905763A24E70},{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A},{78961D78-DB41-4483-99AF-C36BD0A98701},{3482B53C-018A-457A-BA91-BCDBD1EB106A},{95F98947-C5E4-452B-B8B1-238D887CC168},{154E1D8C-B044-4E32-B40B-B6857DD00B5A},{E4BF3B18-ACAE-4E77-8BDF-34CF75481C34},{8FF3157E-1099-4256-A801-51DE178950AF},{9EB619A3-61E8-4C1A-9573-27C87DABEF76},{4AEC67F9-80C1-427F-904E-489920B71940},{9850A686-F991-4F36-8EF2-C0F043103276},{08481F91-3506-40EE-852F-FAB1568E935E},{ECEAF686-3552-44BD-A49B-941376AE4109},{6A8607D5-88A1-4706-87D4-B37D633B2671},{84E1A1BB-0CBB-4258-9017-9D92EEEE2522},{4B95FC42-1E31-4574-93B2-FAC28C5D7C3A}"
   End If


ElseIf sDepartmentRoot = SIT_STS Then 'СТР
  If (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
    If InStr(UCase(CurrentClassDoc), UCase(SIT_RASP_DOCS)) = 1 and FieldToEdit = "DocListToReconcile" Then
        CurrentDirectory = "U"
    Else
      Select Case UCase(Request("l"))
        Case "RU" 'RU
          CurrentDirectoryGUID = "{71B3A81A-0A45-41BB-8DAE-A0EB1DA292A6}"
        Case "" 'EN
          CurrentDirectoryGUID = "{F70782E3-B1A3-4AFE-B800-905763A24E70}"
        Case "3" 'CZ
          CurrentDirectoryGUID = "{3012D11D-199C-4D46-8B58-6704EEF4A3EF}"
      End Select
    End If
  End If
'Запрос №1 - СИБ - start
ElseIf sDepartmentRoot = SIT_SIB Then 'СИБ
  If (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
    If InStr(UCase(CurrentClassDoc), UCase(SIT_RASP_DOCS)) = 1 and FieldToEdit = "DocListToReconcile" Then
        CurrentDirectory = "U"
    Else
	  'Пока справочник только один - на русском
      CurrentDirectoryGUID = "{95F98947-C5E4-452B-B8B1-238D887CC168}"
    End If
  End If
'Запрос №1 - СИБ - end
ElseIf sDepartmentRoot = SIT_SITRU Then 'DmGorsky
	 ' Справочники ролей не требуются по умолчанию
Else 'Другие (сейчас как в СИТРОНИКС)
  If (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
    Select Case UCase(Request("l"))
      Case "RU" 'RU
        CurrentDirectoryGUID = "{78961D78-DB41-4483-99AF-C36BD0A98701}"
      Case "" 'EN
        CurrentDirectoryGUID = "{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A}"
      Case "3" 'CZ нет, используется английский
        CurrentDirectoryGUID = "{7E4C27B9-CF2A-4620-8B98-7CBF965CB93A}"
    End Select
  End If
End If

'vnik_payment_order
If InStr(UCase(CurrentClassDoc), UCase(SIT_PAYMENT_ORDER))>0 Then
    If (FieldToEdit = "UserFieldText1") Then
        'CurrentDirectoryGUID = "{B6A71449-1D6F-42CB-8F3E-B29D03D5D19F}"
	CurrentDirectoryGUID = "{08481F91-3506-40EE-852F-FAB1568E935E}"
    ElseIf (FieldToEdit = "UserFieldText2") Then
        CurrentDirectoryGUID = "{6D6236CD-DA05-4B52-87F1-6C657F2544EE}"
    ElseIf (FieldToEdit = "UserFieldText3") Then
        CurrentDirectoryGUID = "{15EB5243-22D8-425D-B31A-9CBA4396FCFC}"
    ElseIf (FieldToEdit = "UserFieldText5") Then
        CurrentDirectoryGUID = "{507B2058-B7C4-40B2-8EC4-D75A8E4CE28D}"
    End If   
End If
'vnik_payment_order

'vnik_purchase_order
If InStr(UCase(CurrentClassDoc), UCase(SIT_PURCHASE_ORDER))>0 Then
    If (FieldToEdit = "UserFieldText1") Then
        'CurrentDirectoryGUID = "{B6A71449-1D6F-42CB-8F3E-B29D03D5D19F}"
	CurrentDirectoryGUID = "{08481F91-3506-40EE-852F-FAB1568E935E}"
    ElseIf (FieldToEdit = "UserFieldText5") Then
        CurrentDirectoryGUID = "{507B2058-B7C4-40B2-8EC4-D75A8E4CE28D}"
    End If   
End If
'vnik_purchase_order

'rti_payment_order
If InStr(UCase(CurrentClassDoc), UCase(RTI_PAYMENT_ORDER))>0 Then
    CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ",{88DF8A0F-A36D-4CF4-81D0-6AC6D9D33643} "
    VAR_NotToShowUserDirectories = "Y"
    If (FieldToEdit = "UserFieldText2") Then
        CurrentDirectoryGUID = "{6A8607D5-88A1-4706-87D4-B37D633B2671}"
    ElseIf (FieldToEdit = "UserFieldText4") Then
        CurrentDirectoryGUID = "{365C2A1C-D404-47AF-AC76-9421A36E8E6A}"
    End If   
End If
'rti_payment_order

'rti_purchase_order
If InStr(UCase(CurrentClassDoc), UCase(RTI_PURCHASE_ORDER))>0 Then
    CurrentProhibitedDirectoryGUIDs = CurrentProhibitedDirectoryGUIDs & ",{88DF8A0F-A36D-4CF4-81D0-6AC6D9D33643} "
    VAR_NotToShowUserDirectories = "Y"
    If (FieldToEdit = "UserFieldText2") Then
	    CurrentDirectoryGUID = "{6A8607D5-88A1-4706-87D4-B37D633B2671}"
    ElseIf (FieldToEdit = "UserFieldText4") Then
	    CurrentDirectoryGUID = "{365C2A1C-D404-47AF-AC76-9421A36E8E6A}"
    End If   
End If
'rti_purchase_order

'mikron_purchase_order
If InStr(UCase(CurrentClassDoc),UCase(MIKRON_RL_MEMO)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_S_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_PURCHASE_ORDER)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_EXPORT_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_EXPADD_CONTRACT)) > 0 Then
   If (FieldToEdit = "UserFieldText4") Then 'amw привязываем к полю справочник статей затрат "Бюджетный классификатор МИКРОН"       
      CurrentDirectoryGUID = MIKRON_CATALOG_EXPENDITURE
   End If
End If

''mikron_purchase_order
''mikron_BSAP | mikron Protocol ZK
If InStr(UCase(CurrentClassDoc),UCase(MIKRON_BSAP)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_RL_MEMO)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_RL_PROTOCOL)) > 0 Then
   If (FieldToEdit = "UserFieldText1") Then  'amw привязываем к полю Справочник Контрагенты
      CurrentDirectory = "P"
   ElseIf (FieldToEdit = "UserFieldText2") Then  'amw привязываем к полю Справочник Контрагенты
      CurrentDirectory = "P"
   ElseIf (FieldToEdit = "UserFieldMoney1") Then 'amw привязываем к полю Справочник Валюта
      CurrentDirectory = "!"
   End If
End If
''mikron_BSAP | mikron Protocol ZK
''mikron_contract
If InStr(UCase(CurrentClassDoc),UCase(MIKRON_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_S_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_ADD_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_EXPORT_CONTRACT)) > 0 or _
   InStr(UCase(CurrentClassDoc),UCase(MIKRON_EXPADD_CONTRACT)) > 0 Then
   If (FieldToEdit = "UserFieldText1") Then  'amw привязываем к полю Справочник Контрагенты
      CurrentDirectory = "P"
   ElseIf (FieldToEdit = "DocListToReconcile") or (FieldToEdit = "DocNameAproval") Then
      CurrentDirectoryGUID = MIKRON_CATALOG_ROLES
   End If
End If
''mikron_contract

''mikron_payment_order
''mikron_payment_order

' *** ЗАЯВКИ ДЛЯ СТС
If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PaymentOrder)) > 0 or InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) = 1 Then
  Select Case FieldToEdit
    Case "UserFieldText1"
      CurrentDirectoryGUID = "{33F9C053-E51D-4738-91CD-45ABB82C1D8A}"
'    Case "UserFieldText2"
'      CurrentDirectoryGUID = "{8E24E3EF-F350-4D29-8BA0-430E425F54E0}"
    Case "UserFieldText3"
      CurrentDirectoryGUID = "{2CC714EB-5836-49E9-B873-3A34EDB85098}"
    Case "UserFieldText5"
      CurrentDirectoryGUID = "{2FC22FA3-CCC7-41F9-8137-F907DC9C1F24}"
    Case "UserFieldText8"
      CurrentDirectoryGUID = "{3A4F4557-A6E8-4382-A69F-59CF8895645F}"
  End Select
  If InStr(UCase(Session("CurrentClassDoc")),UCase(STS_PurchaseOrder)) > 0 Then
    If FieldToEdit = "DocNameResponsible" Then
	  CurrentDirectory = "D"
	End If
'Запрос №43 - СТС - start
'    If FieldToEdit = "DocPartnerName" Then
'	  CurrentDirectoryGUID = "{F5E696BF-135A-4C81-9920-7F444D51CE14}"
'	End If
'Запрос №43 - СТС - end
  End If
End If

'rmanyushin 136964 08.11.2010 Start
'Запрос №46 - СТС - start
' If InStr(UCase(CurrentClassDoc), UCase(STS_SLUZH_ZAPISKA_OVERTIME2))>0 Then
 If InStr(UCase(CurrentClassDoc), UCase(STS_SLUZH_ZAPISKA_OVERTIME2)) = 1 or InStr(UCase(CurrentClassDoc), UCase(STS_SLUZH_ZAPISKA_OVERTIME_PLAN)) = 1 or InStr(UCase(CurrentClassDoc), UCase(STS_SLUZH_ZAPISKA_OVERTIME_FACT)) = 1 Then
'Запрос №46 - СТС - end
   Select Case FieldToEdit
     Case "UserFieldText3"
       CurrentDirectoryGUID = "{2CC714EB-5836-49E9-B873-3A34EDB85098}"
     Case "DocNameResponsible"
	   CurrentDirectory = "U"	
    End Select
 End If
'rmanyushin 136964 08.11.2010 End

'Назначение справочников для полей других справочников
'Запрос №1 - СИБ - start
If FieldToEdit = "Leader" or FieldToEdit = "_Users" or FieldToEdit = "_UsersList" or FieldToEdit = "DirField2" Then
'If FieldToEdit = "Leader" or FieldToEdit = "_Users" or FieldToEdit = "_UsersList" Then
'Запрос №1 - СИБ - end
  CurrentDirectory = "U"
End If
If FieldToEdit = "BusinessUnit" or FieldToEdit = "_BusinessUnit" or FieldToEdit = "BusinessUnits" Then
  CurrentDirectoryGUID = "{8E24E3EF-F350-4D29-8BA0-430E425F54E0}"
End If

'Запрос №31 - СТС - start
'Справочник Согласующих/Регистраторов по БЕ (ДЗК)
'If InStr("#_RegistrarIn#_RegistrarOut#_RegistrarOrder#_ListToReconcileOut#_ListToReconcileOrder#", "#" & FieldToEdit & "#") Then
'Поля справочника правил
If InStr("#_ListToReconcile#_NameAproval#_NameResponsible#_Correspondent#_ListToView#_Registrar#_NameControl#_ListToEdit#", "#" & FieldToEdit & "#") Then
  CurrentDirectory = "U"
End If
If FieldToEdit = "_ChartOfAccounts" Then
  CurrentDirectoryGUID = "{3A4F4557-A6E8-4382-A69F-59CF8895645F}"
End If
If FieldToEdit = "_CostCenter" Then
  CurrentDirectoryGUID = "{33F9C053-E51D-4738-91CD-45ABB82C1D8A}"
End If
If FieldToEdit = "_ProjectCode" Then
  CurrentDirectoryGUID = "{2CC714EB-5836-49E9-B873-3A34EDB85098}"
End If
If FieldToEdit = "_ClassDoc" Then
  CurrentDirectory = "T"
End If
If FieldToEdit = "_PartnerName" Then
  CurrentDirectory = "P"
End If
'Справочник не получается поставить, он многоязычный, а заявки при сохранении заменяют значения на англ., здесь должно быть только англ. значение
'If FieldToEdit = "_KindOfPayment" Then
'  CurrentDirectoryGUID = "{2FC22FA3-CCC7-41F9-8137-F907DC9C1F24}"
'End If
'Запрос №31 - СТС - end

End Sub
%>