<%
Sub UserProhibitedDirectories
'Set the prohibited directory list
'Variables:
'FieldToEdit - field name from which directory were called
'ExtGUID - editable data source GUID
'ExtDirGUID - directory data source GUID
'CurrentDocID - document ID from which directory were called
'CurrentClassDoc - document category from which directory were called
'Session("WriteSecurityLevel"), Session("ReadSecurityLevel") - write/read Security Levels of current user
'VAR_AdminSecLevel - Administrative Level constant
'
'CurrentProhibitedDirectories - text string containing indexes of prohibited directories
' U - user list directory
' C - contact names directory
' P - partner list directory
' D - department list directory
' I - inventory unit directory
' L - external data directory
' A - activity directory
' B - document registry directory
' T - document category directory
' Z - position directory
' R - report types directory
' E - currency rates directory
' K - context marks directory
' M - measure units directory
' F - user folders directory
' 7 - company list directory
' 8 - operation type directory
' 9 - report type directory

' G - user access groups

'
' Y - reserverd for paper file registry
'

'Select Case CurrentClassDoc
'    Case "Invoices" ' - document category to be processed
			'CurrentProhibitedDirectories = "MZ" 'measure units directory and position directory are prohibited for "Invoices" document category
    'Case "???" ' - other directory GUID to be processed
    			'....
'End Select

CurrentProhibitedDirectories="MI" 'measure units directory and inventory directory  are usually not used
If RUS()<>"RUS" Then
	CurrentProhibitedDirectories=CurrentProhibitedDirectories+"B" 'document registry directory is used only in Russia
End If

If Var_ApplicationType="Пропуска" Then
	CurrentProhibitedDirectories=CurrentProhibitedDirectories+"IABEKMF8"
End If

If IsHelpDesk() Then
	CurrentProhibitedDirectories=CurrentProhibitedDirectories+"IABEKMF8"
End If

If Var_ApplicationType=DOCS_Chancery Then
	CurrentProhibitedDirectories="ILABZRKM789"
	CurrentProhibitedDirectoryGUIDs="{460C85E8-9602-9C36-1FA1-EFA8F3C59127}, {A4BB9E88-A9B3-209A-F4BE-8CF11FF5CB81}, {EAB2C1BF-2676-E606-B671-7D7B051A5DC4}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}, {49463843-832E-E2AE-5673-D72BD1997598}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {3BC76743-EF9D-5221-F1A9-BC585E2CC162}, {B4F7DFE6-E326-E8A7-559C-5288A790354C}, {680412FE-B15E-281B-E23E-CB30851BD31E}, {B21FCB1B-76F5-D2FF-993E-0BCB4B2FC37D}, {07231FC3-91EC-5F5B-A1E4-83DCB5387939}, {FBAE2C89-AEB3-411E-6411-E87700C3EF4F}, {A67904A1-B9CE-CD83-A8E0-CB1CDEBAF62A}, {53D3E531-8DCB-4413-603A-8268443FCBFF}, {D2D11C5C-F9D8-9DE2-AFA1-A431C7E4DFFD}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {49463843-832E-E2AE-5673-D72BD1997598}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}"
End If
If IsPublucUser() Then
	CurrentProhibitedDirectories="ILABTZRKMF789"
	CurrentProhibitedDirectoryGUIDs="{460C85E8-9602-9C36-1FA1-EFA8F3C59127}, {A4BB9E88-A9B3-209A-F4BE-8CF11FF5CB81}, {EAB2C1BF-2676-E606-B671-7D7B051A5DC4}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}, {49463843-832E-E2AE-5673-D72BD1997598}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {3BC76743-EF9D-5221-F1A9-BC585E2CC162}, {B4F7DFE6-E326-E8A7-559C-5288A790354C}, {680412FE-B15E-281B-E23E-CB30851BD31E}, {B21FCB1B-76F5-D2FF-993E-0BCB4B2FC37D}, {07231FC3-91EC-5F5B-A1E4-83DCB5387939}, {FBAE2C89-AEB3-411E-6411-E87700C3EF4F}, {A67904A1-B9CE-CD83-A8E0-CB1CDEBAF62A}, {53D3E531-8DCB-4413-603A-8268443FCBFF}, {D2D11C5C-F9D8-9DE2-AFA1-A431C7E4DFFD}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {49463843-832E-E2AE-5673-D72BD1997598}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}"
End If

'CurrentProhibitedDirectoryGUIDs="{FBAE2C89-AEB3-411E-6411-E87700C3EF4F}, {A67904A1-B9CE-CD83-A8E0-CB1CDEBAF62A}"


' SAY 2008-08-22
' прячем лишние справочники от простых пользователей
If Not IsAdmin() Then

  'стандартные справочники
  CurrentProhibitedDirectories = "ILABTZREKM89GY"

  'пользовательские справочники
' AM 24082008  CurrentProhibitedDirectoryGUIDs="{CAAA819C-DBBA-4B38-9001-58CD15FDC678}, {F0103F47-DA1C-47BC-ACAD-DE69AAF0F852},{16D4E61F-FCCB-44AE-AB53-7AC07285C6A5},{D68C37BE-9EBA-4CA8-B0D2-C6369123E7C7},{C632B46B-3AAF-4607-BBC5-AC51C0A4971B},{3685D3AA-FB15-4ECF-993F-8AC5AB87F4D6},{E0E79CEC-5DDE-4184-92BE-85556566BD14},{BFC71550-2605-4679-8A3F-C04211891D7E},{3ECADCD6-0985-4659-8774-C8C9D77EE381},{6D42F0FB-1389-4CA0-8A00-9E0CD3F09CF7},{0575A1AB-617A-4B3F-A80F-25AFC7E83ABF},{C028387B-3E99-438B-85D2-038397B73181},{37E16CD5-BC8F-4D0C-9569-D14DAA895440}"
  'out CurrentProhibitedDirectoryGUIDs
End If

End Sub

%>