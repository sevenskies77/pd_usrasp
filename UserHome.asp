<%
'Place here some ASP code and variables for Home.asp page
'
'sFrontPic - front-end picture 216x480 pixels
'sFrontPic="UserImages/Front.jpg"
'
'sLogoSmall - to show smallier PayDox logo on the home page
sLogoSmall="Y"
'
'sNoPaymentButtons - do not show payment buttons 
If IsPublucUser() Then
	sNoPaymentButtons="Y"
	sNoLeftMenu=sNoLeftMenu+"ABSLMF"
End If

'sNoCentralButtons - do not show cental buttons 
'C - control button/all controlled documents
'Y - control button/you are the controlling person
'N - completion button
'R - responsible button
'V - documents to agree button
'O - documents to aprove by you button
'B - documents to aprove button
'U - urgent button
'W - viewed button
'L - calendar button
'P - public documents
'M - outstanding documents created by me
'A - all documents for me
'G - list of aproved documents required to be register
'D - documents having your resolutions 
'H - documents having all resolutions resolutions for you
'T - not completed yet documents
'S - incoming payments
' 
sNoCentralButtons=""
If IsPublucUser() Then
	sNoCentralButtons="GANRVOMYWOEPSDH"
End If
'
'sNoRightPaneButtons - do not show right pane buttons 
'F - add favorite button and add shortcut button
'L - check license button
'D - get description button
'M - monitor button
'O - in/out office button
'S - e-mail support button
'A - application type button
' 
'sNoRightPaneButtons="LDS"
sNoRightPaneButtons="FLDMOSA"
'
'sNoLeftMenu - do not show some left menu items
'A - activities
'D - documents
'B - business processes
'S - scorecard dashboard
'R - reports
'L - registration logs
'C - directories
'M - administration
'T - document templates
'F - user folders
' 
'sNoLeftMenu="ADBSRLCMT"

If RUS()<>"RUS" Then
	sNoLeftMenu=sNoLeftMenu+"L" 'registration logs are used only in Russia
End If
If Not bDEMO Then 'switching OFF scorecard dashboard by default in the Not DEMO mode
	'sNoLeftMenu=sNoLeftMenu+"S"
End If

'sTitlePicture=title text for large picture
'
'sTitlePicture="My customized text"
'
'sNoFooterLinks - do not show footer links 
'H - home page
'M - my account
'C - contacts
' 
'sNoFooterLinks="HMC"
'
'sNoTopScreenSearchPane - do not show top screen search pane
'
'sNoTopScreenSearchPane="Y"
'
'bSystemHomeScreenText - show system or show customized HTML-code for the system front page
'
bSystemHomeScreenText=True
'bSystemHomeScreenText=False
Sub UserHomeScreenText 'place here your customized HTML-code for the system front page
%>
Your customized HTML-code for the system front page
<%
End Sub
Sub UserHomeScreenTextBottom 'place here your customized HTML-code for the bottom of the system front page
%><!--Insert here your customized HTML-code for the bottom of the system front page--><%
End Sub

'HomeRightUserMenu cParRef, cParHead, cParText, cSQL - add your own menu in the right side of the system front page
'
'Example:
'HomeRightUserMenu "http://www.mywebaddress.com", "My title 1", "My text 1"
'HomeRightUserMenu "http://www.mywebaddress.com", "My title 2", "My text 2"

'
'AddCentralButton cLink, cTitle, cName, cSQL- add your own large blue button on the system front page
'
'Example:
'AddCentralButton "http://www.mywebaddress.com", "My title 1", "MY BUTTON 1", ""
'AddCentralButton "ListDoc.asp", "My invoices", "Invoices I have created", "select * from Docs where ClassDoc='Invoices' and NameCreation= '"+Session("UserID")+"'"
'AddCentralButton "ListDoc.asp?l=ru", "Счета, созданные мной", " Мои счета ", "select * from Docs where ClassDoc= 'Счета-фактуры' and NameCreation= '"""+Session("Name")+""" <"+Session("UserID")+">;'"
'AddCentralButton "ListDoc.asp?l=ru", "Счета, созданные мной", " Мои счета 1", "select * from Docs where ClassDoc= 'Счета-фактуры' and NameCreation= '"""+Session("Name")+""" <"+Session("UserID")+">;'"
'AddCentralButton "GetReport.asp?NameRequest=Состояние согласования документов&l=ru&DocReportDetailed=ON", "Состояние согласования документов 1", "СОГЛАСОВАНИЕ", ""
'
If Session("UserID") = "donkovtsev_oaorti" or Session("UserID") = "mzabolotneva_oaorti" Then
AddCentralButton "ListDoc.asp?l=ru", "Документы, по которым я отказал(-а) в согласовании", "СОГЛАСОВАНИЕ: ОТКАЗЫ","select * from Docs where ListReconciled like N'%-<"+Session("UserID")+">;%'"
end if

If Session("UserID") = "boev_oaorti" Then
AddCentralButton "ListDoc.asp?l=ru", "Исполненные поручения, в которых я поручитель/контролер", "ИСПОЛНЕННЫЕ ПОРУЧЕНИЯ","select * from docs where ClassDoc = N'Поручения*Tasks*Úkoly' and (Author like N'%<boev_oaorti>%' or NameControl like N'%<boev_oaorti>%' ) and DateCompleted is not null and StatusCompletion = 1 order by DateCompleted desc"
end if

'If Session("UserID") = "usr_savchenko" or Session("UserID") =  "usr_Admin" Then
'AddCentralButton "ListDoc.asp?l=ru", "Поручения, по которым было запрошено исполнение, но мне НЕ БЫЛО отправлено уведомление", "ЗАПРОШЕНО ИСПОЛНЕНО","select * from docs where Department like N'MINC%' and ClassDoc = N'Поручения*Tasks*Úkoly' and StatusCompletion = N'+' and StatusDevelopment = 1 and author like N'%<usr_savchenko>%' and docid not in (select distinct docid from Comments where docid in (select docid from docs where Department like N'MINC%' and ClassDoc = N'Поручения*Tasks*Úkoly' and StatusCompletion = N'+' and StatusDevelopment = 1 and author like N'%<usr_savchenko>%') and CommentType = N'NOTIFICATION' and Comment like N'%запрошено%')"
'end if

If Session("UserID") = "vivanon_oaorti" Then
AddCentralButton "ListDoc.asp?l=ru", "Мои документы, имеющие отказ в согласовании", "СОГЛАСОВАНИЕ: ОТКАЗЫ","select * from docs where author like N'%<vivanon_oaorti>%' and ListReconciled like N'%-<%' and not (statuscompletion = N'0' and isnull(DateCompleted,'')<>'') order by docid"
end if

If IsHelpDesk() Then
'Out "HelpDesk"
	sNoLeftMenu="ABLFST"
	'sNoCentralButtons="CYNVUWMOL"
	sNoCentralButtons="AGCNRVTUWOMYOBES"
	sNoPaymentButtons="Y"
	AddCentralButton "ChangeDoc.asp?reg=&create=y&ClassDoc="+Doc_ClassDocHelpDesk+"&DocID=&ActDoc="+LPar()+"&empty=y", Doc_HelpDesk_NewIncident, But_HelpDesk_NewIncident, ""
	
	If sVersion = "MSSQL" Then
   		sTemp = "select * from Docs where ClassDoc= '"+Doc_ClassDocHelpDesk+"' and CHARINDEX('<" + Session("UserID") + ">', NameCreation)>0 "
	ElseIf sVersion = "MSACCESS" Then
   		sTemp = "select * from Docs where ClassDoc= '"+Doc_ClassDocHelpDesk+"' and InStr(NameCreation, '<" + Session("UserID") + ">')>0 "
	End If
	'If IsHelpDeskUser() Then
	AddCentralButton "ListDoc.asp"+LPar1()+"&ClassDoc="+Doc_ClassDocHelpDesk, Doc_HelpDesk_MyIncidents, But_HelpDesk_MyIncidents, sTemp
	'Else
	If IsAdmin() or IsHelpDeskAdminOrConsultant() Then
		AddCentralButton "ListDoc.asp?ClassDoc=HelpDesk"+LPar(), Doc_HelpDesk_NotDistibutedYet, But_HelpDesk_NotDistibutedYet, "select * from Docs where ClassDoc= '"+Doc_ClassDocHelpDesk+"' and (NameResponsible= '' or NameResponsible='"+DOCS_ToBeDefined+"') and (Correspondent= '' or Correspondent is Null)"
       If sVersion = "MSSQL" Then
       	sTemp = " and CHARINDEX('<" + Session("UserID") + ">', ISNULL(Correspondent, ''))>0 "
       ElseIf sVersion = "MSACCESS" Then
			sTemp = " and InStr(IIF(IsNull(Correspondent),'',Correspondent), '<" + Session("UserID") + ">')>0 "
		End If
		If IsHelpDeskAdmin() Then
			sTemp = " and Correspondent<>'' and Not Correspondent Is Null  "
		End If
		AddCentralButton "ListDoc.asp?ClassDoc=HelpDesk"+LPar(), IIF(IsHelpDeskAdmin(),Doc_HelpDesk_NotAssignetYet, Doc_HelpDesk_NotAssignetYet0), But_HelpDesk_NotAssignetYet, "select * from Docs where ClassDoc= '"+Doc_ClassDocHelpDesk+"' and (NameResponsible= '' or NameResponsible='"+DOCS_ToBeDefined+"')"+sTemp
		AddCentralButton "ListDoc.asp?ClassDoc=HelpDesk"+LPar()+"&ShowComments=comment&RespDocs=y&ClassDoc="+Doc_ClassDocHelpDesk, DOCS_YouAreResponsible, BUT_RESPONSIBLE, ""
		AddCentralButton "ListDoc.asp?ClassDoc=HelpDesk"+LPar()+"&ShowComments=&searchstatus="+HTMLEncode(DOCS_OUTSTANDING), DOCS_OUTSTANDING, BUT_OUTSTANDING, ""
		'AddCentralButton "", "", "", ""
		'AddCentralButton "", "", "", ""
	End If
End If

'AddLogD "Var_ApplicationType:"+Var_ApplicationType
If Var_ApplicationType="Пропуска" Then
AddLogD "*Пропуска!*"

	sNoCentralButtons="AGCNVTUMYLWOBES"
	sNoPaymentButtons="Y"
	sNoLeftMenu="ABSL"
	AddCentralButton "ListDoc.asp?l=ru&ClassDoc=Пропуска%20/%20Разовые%20пропуска", "Разовые пропуска", "РАЗОВЫЕ", ""
	AddCentralButton "ListDoc.asp?l=ru&ClassDoc=Пропуска%20/%20Разовые%20пропуска%20для%20выходных", "Разовые пропуска для выходных", "ДЛЯ ВЫХОДНЫХ", ""
	AddCentralButton "ListDoc.asp?l=ru&ClassDoc=Пропуска%20/%20Разовые%20пропуска%20для%20иностранцев", "Разовые пропуска на иностранцев", "НА ИНОСТРАНЦЕВ", ""
	AddCentralButton "ListDoc.asp?l=ru&ClassDoc=Пропуска%20/%20Разовые%20пропуска%20для%20списка", "Разовые пропуска для списка", "ДЛЯ СПИСКА", ""
	AddCentralButton "ListDoc.asp?l=ru&ClassDoc=Пропуска%20/%20Разовые%20пропуска%20на%20ам", "Разовые пропуска на ам", "НА А/М", ""
	sFrontPic="UserImages/Propusk.gif"
End If

If Var_ApplicationType=DOCS_Controller Then
	sNoCentralButtons="ARNVWGMEOBS"
	sNoPaymentButtons="Y"
	sNoLeftMenu="ABSLTMRC"
	sNoRightPaneButtons="FLDMOS"
	If RUS()="RUS" Then
		AddCentralButton "GetReport.asp?l=ru&DocReportDetailed=ON&NameRequest="+HTMLEncode("Состояние согласования документов"), "Состояние согласования документов", UCase("Согласование"), ""
	End If
	sFrontPic="Images/Controlling"+RUS()+".jpg"
End If

If Var_ApplicationType=DOCS_Viewing Then
	sNoCentralButtons="GACNRVTUOMYLOBES"
	sNoPaymentButtons="Y"
	sNoLeftMenu="ABSLTMRCD"
	sNoRightPaneButtons="FLDMOS"
	sFrontPic="Images/Viewing"+RUS()+".jpg"
End If

If Var_ApplicationType=DOCS_Chancery Then
	sNoCentralButtons="ARNVTMYWOBPSED"
	sNoPaymentButtons="Y"
	sNoLeftMenu="ABSLTMR"
	sNoRightPaneButtons="FLDMOS"
	AddCentralButton "ListDoc.asp?l="+Request("l")+"&Incomings=y", DOCS_Incomings+" - "+LCase(Title_Actual), UCase(DOCS_Incomings), ""
	AddCentralButton "ListDoc.asp?l="+Request("l")+"&Outgoings=y", DOCS_Outgoings+" - "+LCase(Title_Actual), UCase(DOCS_Outgoings), ""
	sFrontPic="Images/Chancery"+RUS()+".jpg"
ElseIf Not IsPublucUser() Then
	sNoCentralButtons=sNoCentralButtons+"BS"
End If

If Var_ApplicationType=DOCS_Chancery Then
	CurrentProhibitedDirectories="ILABZRKM789"
	CurrentProhibitedDirectoryGUIDs="{460C85E8-9602-9C36-1FA1-EFA8F3C59127}, {A4BB9E88-A9B3-209A-F4BE-8CF11FF5CB81}, {EAB2C1BF-2676-E606-B671-7D7B051A5DC4}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}, {49463843-832E-E2AE-5673-D72BD1997598}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {3BC76743-EF9D-5221-F1A9-BC585E2CC162}, {B4F7DFE6-E326-E8A7-559C-5288A790354C}, {680412FE-B15E-281B-E23E-CB30851BD31E}, {B21FCB1B-76F5-D2FF-993E-0BCB4B2FC37D}, {07231FC3-91EC-5F5B-A1E4-83DCB5387939}, {FBAE2C89-AEB3-411E-6411-E87700C3EF4F}, {A67904A1-B9CE-CD83-A8E0-CB1CDEBAF62A}, {53D3E531-8DCB-4413-603A-8268443FCBFF}, {D2D11C5C-F9D8-9DE2-AFA1-A431C7E4DFFD}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {49463843-832E-E2AE-5673-D72BD1997598}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}"
End If
If IsPublucUser() Then
	CurrentProhibitedDirectories="ILABTZRKMF789"
	CurrentProhibitedDirectoryGUIDs="{460C85E8-9602-9C36-1FA1-EFA8F3C59127}, {A4BB9E88-A9B3-209A-F4BE-8CF11FF5CB81}, {EAB2C1BF-2676-E606-B671-7D7B051A5DC4}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}, {49463843-832E-E2AE-5673-D72BD1997598}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {3BC76743-EF9D-5221-F1A9-BC585E2CC162}, {B4F7DFE6-E326-E8A7-559C-5288A790354C}, {680412FE-B15E-281B-E23E-CB30851BD31E}, {B21FCB1B-76F5-D2FF-993E-0BCB4B2FC37D}, {07231FC3-91EC-5F5B-A1E4-83DCB5387939}, {FBAE2C89-AEB3-411E-6411-E87700C3EF4F}, {A67904A1-B9CE-CD83-A8E0-CB1CDEBAF62A}, {53D3E531-8DCB-4413-603A-8268443FCBFF}, {D2D11C5C-F9D8-9DE2-AFA1-A431C7E4DFFD}, {98E64AF6-7C8B-C4E5-CE9A-1C7C9102B3F4}, {49463843-832E-E2AE-5673-D72BD1997598}, {16F64E7E-D3C9-7923-4FD1-49F6111BD56E}"
End If
If InStr(sNoCentralButtons, "O")>0 Then
	sNoCentralButtons=Replace(sNoCentralButtons,"B", "")
End If
'If Var_ApplicationType<>DOCS_Chancery Then
'	sNoCentralButtons=sNoCentralButtons+"H"
'End If

sNoLeftMenu=sNoLeftMenu+"L" 'NO registration logs
'AddCentralButton "ListDoc.asp?l=ru&delegate=y", "Документы, по которым Вы делегировали согласование", "ДЕЛЕГИРОВАНИЕ", "Select * from Docs Left Outer Join Comments ON (Docs.DocID = Comments.DocID And Comments.CommentType='VISA' and Comments.SpecialInfo='DELEGATE' and Comments.UserID = '" + Session("UserID") + "') where StatusCompletion<>'" + VAR_StatusCompletion + "' and StatusCompletion<>'" + VAR_StatusCancelled +" order by Comments.DateCreation desc" 


' SAY 2008-08-21
sNoCentralButtons=sNoCentralButtons+"CT"
' SAY 2008-08-28 дополнительно
sNoCentralButtons=sNoCentralButtons+"DHL"
' SAY 2008-11-10
sNoCentralButtons=sNoCentralButtons+"Y"

sNoPaymentButtons="Y"
'sNoRightPaneButtons="LSOAM"
sNoRightPaneButtons="LOAM"


' кнопка создания ярлыка только на стартовой
If UCase(Request.ServerVariables("URL"))=UCase("/home.asp") Then
  sNoRightPaneButtons=sNoRightPaneButtons+"F"  
End If

If Not IsAdmin() Then
  'Кнопку Регистрация прячем для не регистраторов
  If Not CheckPermit(Session("Permitions"), "REGISTRAR") Then
    sNoCentralButtons = sNoCentralButtons + "G"
  End If
  'прячем лишнее от простых пользователей
  'vnik_archive
  If InStr(UCase(ReplaceRoleFromDir(SIT_AccessToArchive,Session("Department"))),UCase(Session("UserID"))) > 0 Then
    sNoLeftMenu = "ABSLT"
  Else
  'vnik_archive
  sNoLeftMenu = "ABSLMT"
End If
End If

sNoLeftMenu = sNoLeftMenu +"B"
sNoLeftMenu = sNoLeftMenu +"Z" 'сокрытие смены дизайна

'инструкция пользователям
HomeRightUserMenu SIT_LinkInstructionFile, SIT_LinkInstruction, SIT_LinkInstructionHint

'SAY 2008-10-22 кнопка для Ушаковой О. для поручений АФК
'If InStr(Session("UserID"),"ushakova") > 0 Then

'регистраторам показываем кнопку с созданными, но неактивированными входящими
  If CheckPermit(Session("Permitions"), "REGISTRAR") Then
    AddCentralButton "ListDoc.asp?l=ru", "Входящие документы, которые не были активированы после создания", "ВХОДЯЩИЕ: НЕАКТИВНЫЕ","select * from docs d " + _
 "where ClassDoc = N'Входящие документы*Incoming correspondence*Příchozí dokumenty' and Department like N'РТИ%'" + _
 "and isnull(isactive,'N') = N'N'" + _
 "and not (statuscompletion = N'0' and isnull(DateCompleted,'')<>'') order by d.docid"
  End If

If Session("UserID") = "ushakova" Then
'  AddCentralButton "ListDoc.asp?l="+Request("l")+"&T_AFK=Y", SIT_BUT_AFKTasksHint, SIT_BUT_AFKTasks, "select * from Docs where StatusArchiv<>'1' and ExtInt=' ' and ClassDoc IN (N'Поручения*Tasks*objednávek') and userfieldtext1 like N'Поручения АФК' and isActive='Y' order by DateCompletion" 
  AddCentralButton "ListDoc.asp?l="+Request("l")+"&T_AFK=Y", SIT_BUT_AFKTasksHint, SIT_BUT_AFKTasks, "select * from Docs where StatusArchiv<>'1' and ExtInt=' ' and ClassDoc IN ("+sUnicodeSymbol+"'"+DOCS_Notices+"') and userfieldtext1 in ("+SIT_AFK_Tasks+") and isActive='Y' order by DateCompletion" 
End If

'SAY 2008-11-18 прячем helpdesk, убираем быстрый старт
If UCase(Request.ServerVariables("SERVER_NAME")) <> UCase("gl-paydox-01.global.sitronics.com") and Request.ServerVariables("SERVER_NAME") <> "172.26.0.180" and UCase(Request.ServerVariables("SERVER_NAME")) <> UCase("paydox.oaorti.ru") Then
  sNoRightPaneButtons="FDLOM"
Else
  sNoRightPaneButtons="DLOAM"
End If

'Кнопка Соисполнитель
AddCentralButton "ListDoc.asp?UserBtn=CoResponsible"+LPar(), SIT_CentralBut_CoResponsibleHint, SIT_CentralBut_CoResponsible, "select * from Docs where (PATINDEX("+sUnicodeSymbol+"'%<"+Session("UserID")+">%', Correspondent)>0 and (StatusCompletion<>'1' and StatusCompletion<>'0' Or StatusCompletion is Null) and DateCompleted is Null and ClassDoc = "+sUnicodeSymbol+"'"+DOCS_Notices+"') order by Docs.DateCreation desc, Docs.DocID desc"
'Кнопка Заявки
If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
  AddCentralButton "ListDoc.asp?UserBtn=PurchaseOrders"+LPar(), SIT_CentralBut_PurchaseOrdersHint, SIT_CentralBut_PurchaseOrders, "select * from Docs where (CHARINDEX("+sUnicodeSymbol+"'"+STS_PurchaseOrder+"', ClassDoc)=1 and (StatusCompletion<>'0' Or StatusCompletion is Null) and (DateCompleted is Null or DateCompleted > GetDate()-30)) order by StatusCompletion, DateCompleted desc, Docs.DateActivation desc, Docs.DocID desc"
End If

'ph - 20101113 - start
'Кнопка Рецензии для СТС
If InStr(Session("Department"), SIT_STS_ROOT_DEPARTMENT) = 1 Then
'{ph - 20120829 - СТС
'  AddCentralButton "ListDoc.asp?UserBtn=Reviews&ShowComments=y"+LPar(), DOCS_Reviews, UCase(DOCS_Reviews), "select Comments.*, Docs.*, BoardOrder =  CASE SUBSTRING(ISNULL(Comments.SpecialInfo, '    '), 1, 4) WHEN 'RESP' THEN SUBSTRING(Comments.SpecialInfo, 5, LEN(Comments.SpecialInfo)-4) ELSE CONVERT(varchar(24), ISNULL(Comments.DateCreation, GetDate()), 120)+' '+(CONVERT(varchar(12), ISNULL(Comments.KeyField, '            '))) END , Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Comments.FileName as CommentsFileName  from Docs  Left Outer Join Comments ON (Docs.DocID = Comments.DocID  And CommentType='REVIEW' And Address='REQUEST') where (CHARINDEX(N' <" & Session("UserID") & ">', Comment) > 0 or UserID = N'" & Session("UserID") & "') and (StatusCompletion is NULL or (StatusCompletion <> '1' and StatusCompletion <> '0')) and (IsActive <> 'N' or IsActive is Null) order by Docs.DateCreation desc, Docs.DocID desc, BoardOrder, Comments.DateCreation"
  AddCentralButton "ListDoc.asp?UserBtn=Reviews&ShowComments=y"+LPar(), DOCS_Reviews, UCase(DOCS_Reviews), "select Comments.*, Docs.*, BoardOrder =  CASE SUBSTRING(ISNULL(Comments.SpecialInfo, '    '), 1, 4) WHEN 'RESP' THEN SUBSTRING(Comments.SpecialInfo, 5, LEN(Comments.SpecialInfo)-4) ELSE CONVERT(varchar(24), ISNULL(Comments.DateCreation, GetDate()), 120)+' '+(CONVERT(varchar(12), ISNULL(Comments.KeyField, '            '))) END , Comments.DateCreation as CommentsDateCreation, Docs.DateCreation as DocsDateCreation, Comments.FileName as CommentsFileName  from Docs  Left Outer Join Comments ON (Docs.DocID = Comments.DocID  And CommentType='REVIEW' And Address='REQUEST') where (CHARINDEX(N' <" & Session("UserID") & ">', Comment) > 0 or UserID = N'" & Session("UserID") & "') and (StatusCompletion is NULL or (StatusCompletion <> '1' and StatusCompletion <> '0')) and (IsActive <> 'N' or IsActive is Null) and (IsNull(NameApproved, '') = '') order by Docs.DateCreation desc, Docs.DocID desc, BoardOrder, Comments.DateCreation"
'ph - 20120829 - СТС}
End If
'ph - 20101113 - end
%>