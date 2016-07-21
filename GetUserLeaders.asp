<!--#INCLUDE VIRTUAL="\TextANSI.txt" -->
<!--#INCLUDE VIRTUAL="\Common.asp" -->
<%

'***********************************
'* Вывод руководителей сотрудников *
'***********************************
Sub GetUserLeaders(pConnectString)
  Dim Conn, ds, i
  
  Set Conn = CreateObject("ADODB.Connection")
  Conn.Open pConnectString
  Set ds = CreateObject("ADODB.Recordset")
  ds.Open "select * from Users where WinLogin like 'global\%' and (StatusActive='"+VAR_StatusActiveUser+"' or StatusActive='"+VAR_StatusActiveUserEMail+"' or StatusActive='"+VAR_StatusNotActiveRole+"') And Not (StatusActive is Null)", Conn, 3, 1, &H1
 '  ds.Open "select * from Users", Conn, 3, 1, &H1

  Response.Write "<TABLE BORDER=1 COLS=5 BGCOLOR=white align=left><TR><TH>№</TH><TH>User</TH><TH>Position</TH><TH>Department</TH><TH>Leader</TH><TH>LeaderOfDepartment</TH><TH>Vice-President</TH></TR>"
  i = 1
  do while not ds.EOF
    sDepartmentToSee = ""

    Response.Write "<TR>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+CStr(i)+"</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+Replace(Replace(GetFullName(SurnameGN(ds("Name")), ds("UserID")), ">", "&gt;"), "<", "&lt;")+";"+"</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+ds("Position")+"</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+DelOtherLangFromFolder(ds("Department"))+"</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+Replace(Replace(GetChiefOfDepUpperByLevel(ds("Department"), 3, ds("UserID"), ""), ">", "&gt;"), "<", "&lt;")+".</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+Replace(Replace(GetChiefOfDepUpperByLevel(ds("Department"), 2, ds("UserID"), ""), ">", "&gt;"), "<", "&lt;")+".</FONT></TD>"
    Response.Write "<TD><FONT size=3 face=""courier new"">"+Replace(Replace(GetChiefOfDepUpperByLevel(ds("Department"), 1, ds("UserID"), ""), ">", "&gt;"), "<", "&lt;")+".</FONT></TD>"
    Response.Write "</TR>"
	i = i + 1
    ds.MoveNext
  loop

  Response.Write "</TABLE>"
  ds.Close
  Conn.Close
End Sub

On Error Resume Next
GetUserLeaders(Application("ConnectStringRUS"))

'GetChiefNameUpFromExtField(ds("Department"), ds("UserID"))
'GetChiefNameUpperWithDepName(ds("Department"), "департамент")
'GetChiefNameFromDepRoot(ds("Department"), 2, sDepartmentToSee)

%>
