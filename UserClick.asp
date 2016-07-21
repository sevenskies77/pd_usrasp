<%
'Sep "Title", "Message" 											'Format: Sep <Title>, <Message>
'Click GetURL("YourURL.asp","?docid=",S_DocID), "Message", "Title" 	'Format: Click <URL>, <Message>, <Title>

'Select Case UCase(Request.ServerVariables("URL"))
	'Case UCase("/ListDoc.asp") 
		'UserSep "TitleListDoc", "MessageListDoc"
		'Click GetURL("YourURL.asp","?docid=",S_DocID), "Message", "Title" 	'Format: Click <URL>, <Message>, <Title>
	'Case UCase("/ShowDoc.asp") 
		'If sButtonJustShown=But_MSWord Then 'User button output after the But_MSWord group of buttons title
		'	Click GetURL("YourURL.asp","?docid=",S_DocID), "Message", "Title" 	'Format: Click <URL>, <Message>, <Title>
		'End If
		'If sButtonJustShown=But_Reconciliation Then 'User button output after the But_Reconciliation group of buttons title
		'	Click GetURL("YourURL.asp","?docid=",S_DocID), "Message1", "Title2" 	'Format: Click <URL>, <Message>, <Title>
		'End If
'End Select

'AddLogD "UserClick.asp"
'AddLogD "sButtonJustShown:"+sButtonJustShown
Select Case sButtonJustShown
'  Case But_MSWord 'сейчас будет вставлена группа кнопок «MS Word»?
'	 Sep "Моя группа кнопок", "Подсказка для моей группы кнопок" 'ставим группу кнопок «Моя группа кнопок» перед группой кнопок «MS Word»
'	 Click "MyPage1.asp?l=ru", "Подсказка для моей кнопки1", "Моя кнопка 1" 'ставим кнопку "Моя кнопка 1" в группе кнопок «Моя группа кнопок»
'	 Click "MyPage2.asp?l=ru", "Подсказка для моей кнопки2", "Моя кнопка 2" 'ставим кнопку "Моя кнопка 2" в группе кнопок «Моя группа кнопок»
'  Case "ClickMSOfficeReconciliationList" 'появилась ли только что кнопка «Лист согласования»?
'	 Click "MyPage2.asp?l=ru", "Подсказка для моей кнопки2", "Моя кнопка 3" 'ставим кнопку "Моя кнопка 3" перед кнопкой «Лист согласования»
'amw 25-10-2013 (start)
'Добавление кнопок прямого создания подчиненных документов, т.к. на Микроне
'крайне низкий уровень квалификации конечных пользователей. Множество ошибок.
'  Case "ClickCreateDocDependant" 'появилась ли только что кнопка «Подчиненный»?
  Case "ClickESign" 'появилась ли только что кнопка «Подчиненный»? вставляем свою кнопку вместо нее
'Для Опросного листа ПЗК подчиненный документ Протокол ЗК. Создавать его может только пользователь
'с ролью "Секретарь ЗК". Данный пользователь может формировать ПЗК для ЗАОЧНОЙ комиссии.
     If InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_PROTOCOL)) > 0 Then
        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
           If InStr(UCase(ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON)),UCase(Session("UserID"))) > 0 Then
              'Проверяем чтобы не создавать больше одного подчиненного
              Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
              sSQL = "SELECT * FROM Docs where DocIDParent = " + _
                     sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocID"))+"'" + _
                     " and ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_PROTOCOL)+"'"
              vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
              If vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
                 Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_PROTOCOL), _
                               "&docid=",S_DocID,"&ClassDoc=", HTMLEncode(S_ClassDoc)), _
                               "Создать карточку документа Протокол ЗК", "Протокол ЗК"
              End If
              vnikdsTemp1.Close
           End If
        End If
'Для Заявки на закупку подчиненный документ либо БСАП если сумма от 50.000 до 500.000, либо Опросный лист.
'Если пользователь "Секретарь ЗК", то он может формировать ПЗК для очной комиссии.
     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
           If dsDoc("UserFieldMoney1") >= 50000 and dsDoc("UserFieldMoney1") < 500000 Then
              'Проверяем чтобы не создавать больше одного подчиненного
              Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
              sSQL = "SELECT * FROM Docs where DocIDParent = " + _
                     sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocID"))+"'" + _
                     " and ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_BSAP)+"'"
              vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
              If vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
                 Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_BSAP), _
                               "&docid=",S_DocID,"&ClassDoc=", HTMLEncode(S_ClassDoc)), _
                               "Создать карточку документа БСАП", "БСАП"
              End If
              vnikdsTemp1.Close
           ElseIf dsDoc("UserFieldMoney1") >= 500000 or dsDoc("UserFieldMoney1") = 0 Then
              If InStr(UCase(ReplaceRoleFromDir(MIKRON_SecretaryOfPC,SIT_MIKRON)),UCase(Session("UserID"))) = 0 Then
                 'Проверяем чтобы не создавать больше одного подчиненного
                 Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                 sSQL = "SELECT * FROM Docs where DocIDParent = " + _
                        sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocID"))+"'" + _
                        " and ( ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_RL_PROTOCOL)+"'" + _
                        " or ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_PROTOCOL)+"' )"
                 vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
                 If vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
                    Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_RL_PROTOCOL), _
                                  "&docid=",S_DocID,"&ClassDoc=",HTMLEncode(S_ClassDoc)), _
                                  "Создать карточку документа опросного листа к ПЗК", "Опросный лист к ПЗК"
                 End If
                 vnikdsTemp1.Close
              Else
                 'Проверяем чтобы не создавать больше одного подчиненного
                 Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                 sSQL = "SELECT * FROM Docs where DocIDParent = " + _
                        sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocID"))+"'" + _
                        " and ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_PROTOCOL)+"'"
                 vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
                 If vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
                    Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_PROTOCOL),"&docid=",S_DocID,"&ClassDoc=", HTMLEncode(S_ClassDoc)), _
                                  "Создать карточку документа Протокол ЗК", "Протокол ЗК"
                 End If
                 vnikdsTemp1.Close
              End If
           End If
        End If
'Для справки к листу согласования подчиненный документ доп.соглашение
     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_RL_MEMO)) > 0 Then
        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
           SIT_bRedirect = False
           'Проверяем, чтобы у справки не было больше одного последующего доп.соглашения
           Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
           sSQL = "SELECT * FROM Docs where DocIDPrevious = " + _
                  sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocID"))+"'" + " and ClassDoc = " + _
                  sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_ADD_CONTRACT)+"'"
           vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
           If vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
              SIT_bRedirect = True
           End If
           vnikdsTemp1.Close          

           If SIT_bRedirect Then
           'Вычисляем категорию родительского документа. Может быть либо договор, либо старый договор.
              Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
              sSQL = "SELECT * FROM Docs where DocID="+sUnicodeSymbol+"'"+MakeSQLSafeSimple(S_DocIDPrevious)+"'"
              vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
              If not vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
                 Click GetURL3("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_ADD_CONTRACT), _
                       "&docid=",S_DocIDPrevious,"&ClassDoc=",HTMLEncode(vnikdsTemp1("ClassDoc")),"&DocIDPrevious=",S_DocID), _
                       "Создать карточку документа Доп.соглашение", "Доп.соглашение"
              End If
              SIT_bRedirect = False
              vnikdsTemp1.Close
           End If
        End If
'для БСАП и для ПЗК подчиненный документ Договор
     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_BSAP)) > 0 or _
            InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PROTOCOL)) > 0 Then
        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
                 'Проверяем чтобы не создавать больше одного подчиненного
                 Set vnikdsTemp1 = Server.CreateObject("ADODB.Recordset")
                 sSQL = "SELECT * FROM Docs where DocIDParent = " + _
                        sUnicodeSymbol+"'"+MakeSQLSafeSimple(Request("DocID"))+"'" + _
                        " and ClassDoc = "+sUnicodeSymbol+"'"+MakeSQLSafeSimple(MIKRON_CONTRACT)+"'"
                 vnikdsTemp1.Open sSQL, Conn, 3, 1, &H1
                 If vnikdsTemp1.EOF Then 'Такого документа нет, можем создавать
                    Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_CONTRACT),"&docid=",S_DocID,"&ClassDoc=", HTMLEncode(S_ClassDoc)), _
                                  "Создать карточку документа Договор", "Договор"
                 End If
                 vnikdsTemp1.Close
        End If
'для Договора подчиненный документ Дополнительное соглашение. Можно создавать несколько.
     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_CONTRACT)) > 0 or _
            InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_S_CONTRACT)) > 0 or _
            InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_OLD_CONTRACT)) > 0 Then
        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
           Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_ADD_CONTRACT), _
                         "&docid=",S_DocID,"&ClassDoc=", HTMLEncode(S_ClassDoc)), _
                         "Создать карточку документа Доп.соглашение", "Доп.соглашение"
        End If
'для Экспортного контракта подчиненный документ - Дополния разного типа(поэтому несколько кнопок). Можно создавать несколько.
     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_EXPORT_CONTRACT)) = 1 Then
        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
           Click GetURL3("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_EXPADD_CONTRACT), _
                         "&docid=",S_DocID,"&ClassDoc=",HTMLEncode(S_ClassDoc),"&DocName=",MIK_EA_1), _
                         "Создать дополнение на разовую отгрузку", MIK_EA_1
           Click GetURL3("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_EXPADD_CONTRACT), _
                         "&docid=",S_DocID,"&ClassDoc=",HTMLEncode(S_ClassDoc),"&DocName=",MIK_EA_2), _
                         "Создать дополнение на расширение номенклатуры", MIK_EA_2
           Click GetURL3("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_EXPADD_CONTRACT), _
                         "&docid=",S_DocID,"&ClassDoc=",HTMLEncode(S_ClassDoc),"&DocName=",MIK_EA_3), _
                         "Создать дополнение на добавление спецификации", MIK_EA_3
           Click GetURL3("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode(MIKRON_EXPADD_CONTRACT), _
                         "&docid=",S_DocID,"&ClassDoc=",HTMLEncode(S_ClassDoc),"&DocName=",MIK_EA_4), _
                         "Создать дополнение на изменение условий", MIK_EA_4
        End If
'     ElseIf InStr(UCase(Session("CurrentClassDoc")),UCase(MIKRON_PURCHASE_ORDER)) > 0 Then
'        If UCase(Request("l"))="RU" and ShowStatusDevelopment(dsDoc("StatusDevelopment"))=DOCS_Approved Then
'           Click GetURL("CreateComment.asp?create=y&review=y","&DocID=",S_DocID)+"&Companies="+"kjhkjhk", "Создать карточку документа БСАП", "БСАП"
'        End If
     End If
'amw 25-10-2013 (end)
End Select

      If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") and _
                InStr(UCase(Session("CurrentClassDoc")), UCase(SIT_ZADACHI)) = 1 and _
                InStr(UCase(Session("Department")), UCase(SIT_RTI)) = 1 Then 'Удостоверяемся, что находимся на странице просмотра поручения ОАО "РТИ"
            If sButtonJustShown = "ClickCreateNotice" Then 'Системная кнопка перед которой вставляем свою кнопку
                Click GetURL2("ChangeDoc.asp?create=y&ClassDocDependant="+HTMLEncode("Служебные записки*Office memo*Interní sdělení/Общая форма*Universal form*Universal form"),"&docid=",S_DocID,"&ClassDoc=", HTMLEncode(S_ClassDoc)), "Создать подчиненную служебную записку", "Служебная записка"
        End If
    End If

'amw DEBUG just for testing
'out sButtonJustShown
'amw DEBUG just for testing

'Кнопка загрузки проектов из файла (при показе редактируемого справочника проектов)
If CanLoadSTSProjectList(Session("UserID")) and UCase(Request.ServerVariables("URL")) = UCase("/ListUserDirVExtData.asp") and InStr(Request("DirGUID"), "2AE2C457-96FE-4379-BC33-BA048E4C06B8") > 0 Then
  If sButtonJustShown = "" Then
    Click "javascript:LoadProjectList()", SIT_ButtonLoadProjectsHint, SIT_ButtonLoadProjects
  End If
End If
%>
<script language="JavaScript"><!--
function LoadProjectList() 
{
  if(confirm('<%=SIT_ButtonLoadProjectsConfirm%>'))
  {
	window.open('<%=GetURL("Agent/AgentSynch.asp","?ManualStart=","ProjectList")%>');
  }
  return;
}
// --></script>
<%

'Кнопка показа таблицы руководителей Ситроникс
If (Session("UserID") = "Admin" or Session("UserID") = "makeev") and (UCase(Request.ServerVariables("URL")) = UCase("/ListUsers.asp")) Then
  If sButtonJustShown = "" Then
    Click "javascript:GetUserLeaders()", "Show table of Leaders", "LeadersTable"
  End If
End If
%>

<script language="JavaScript"><!--
function GetUserLeaders() 
{
  if(confirm('Show table of Leaders?'))
  {
	window.open('<%=GetURL("UserASP/GetUserLeaders.asp","?docid=",S_DocID)%>');
  }
  return;
}
// --></script>

<%
'Запрос №30 - СТС - start
'Показать в заявках на оплату кнопку Переназначить в обход системных правил
If sButtonJustShown = But_Completion Then
  'Показ раздела Исполнение, следом нужно показать кнопку Переназначить, если необходимо
  If bSTS_ShowClickReSetResponsible = "Y" Then
    bSTS_ShowClickReSetResponsible = "YY"
  End If
Else
  If bSTS_ShowClickReSetResponsible = "YY" Then
%>
<%=TableButton%>
  <TBODY><TR><form name="SetRespForm" method="POST" action="ListUsers.asp?addusers=y&reqtype=setresp<%=LPar()%>&DocID=<%=URLEncode(Request("DocID"))%>&UserID=<%=URLEncode(Request("UserID"))%>&NOUserID=<%=URLEncode(Request("UserID"))%>">
<TD <%=MenuRightbgColor%> align="left">
<%If IsHelpDesk() Then%>
<input type="hidden" name="context" value="y">
<%
sContext=S_Correspondent
sContext=Replace(sContext, "<"+Session("UserID")+">", "")
sContext=Replace(sContext, "<"+GetLogin(S_NameResponsible)+">", "")
%>
<input type="hidden" name="Correspondent" value="<%=HTMLEncode(sContext)%>">
<%End If%>
<input type="hidden" name="ClassDoc" value="<%=HTMLEncode(S_ClassDoc)%>">
<A href="javascript:SetRespForm.submit();" title="<%=DOCS_SetResponsible1%>" <%=StyleMenuRight%>><%=But_SetResponsible1%></A>
</TD></TR></TBODY></form></TABLE>
<%
    'Сбрасываем флаг, чтобы больше не выполнять этот код
    bSTS_ShowClickReSetResponsible = ""
  End If
End If
'Запрос №30 - СТС - end


'Запрос №38 - СТС - start
If sButtonJustShown = But_Reconciliation Then
  'Кнопка Добавить в разделе согласования (показывается когда добавление невозможно)
  If bSTS_ShowClickVisaAddFake = "Y" Then
    bSTS_ShowClickVisaAddFake = "YY"
  End If
Else
  If bSTS_ShowClickVisaAddFake = "YY" Then
    'С цветом могут быть проблемы, тогда убрать font
    Click "javascript:alert('" & SIT_CannotAddUsersToListToReconcile & "');", SIT_CannotAddUsersToListToReconcile, "<font color = #C0C0C0>" & But_AGREEADD & "</font>"
    'Сбрасываем флаг, чтобы больше не выполнять код
    bSTS_ShowClickVisaAddFake = ""
  End If
End If
'Запрос №38 - СТС - end
'{Запрос №50 - СТС
If UCase(Request.ServerVariables("URL")) = UCase("/ShowDoc.asp") Then 'Страница просмотра документа
	If IsAdmin() and InStr(S_Department, SIT_STS_ROOT_DEPARTMENT) = 1 Then
		If sButtonJustShown = "ClickAuditingDoc" Then
			Click "javascript:ShowReport_STSRulesLog();", "Получить отчет по логу правил", "Лог правил"
		End If
		%>
		<script language="JavaScript"><!--
		function ShowReport_STSRulesLog() 
		{
			window.open('<%=GetURL4("GetReport.asp?DocPrintableView=ON&DocReportDetailed=ON&R1=HTML&nUserPars=1", "&GUIDRequest=", "DE37E6E8-58E7-468E-A635-1F3BE63BBE9C", "&UserPar1=", Request("DocID"), "&SQLContext1=", "#DOCID#", "&UserParTitle1=", DOCS_DocID)%>');
			return;
		}
		// --></script>
		<%
	End If
End If
'Запрос №50 - СТС}
%>
