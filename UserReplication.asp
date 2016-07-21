<%
'DO NOT LEAVE ANY SPACES OR VbCrLf OR ANY OTHER SYMBOLS OUTSIDE OF ASP CODE!
'Because theese symbols will be inserted into output file and output file will be corrupted!
'Provide the variable VAR_DownloadFileURL in this user setup file UserAsp\UserReplication.asp
'If Left(Request.ServerVariables ("REMOTE_ADDR"),7)="192.168" Then 	'All users from the subnet 192.168.*.*
	'VAR_DownloadFileURL="http://192.168.1.1"	    				'will receive files from their local PayDox server
'End If

'Provide the variable VAR_PathReplicationFiles in this user setup file UserAsp\UserReplication.asp
VAR_PathReplicationFiles="" 'set up this variable to some value to permit file replication from remote PayDox server to this directory
'If Request.ServerVariables ("REMOTE_ADDR"),7)="192.168.1.1" Then 	'permit file replication from remote PayDox server 192.168.1.1 to the subdirectory "ReplicationFiles\" in the PayDox root directory
	'VAR_PathReplicationFiles = Application("PayDoxHomeDir")+"ReplicationFiles\" 
'End If

'If Left(Request.ServerVariables("REMOTE_ADDR"), 8) = "172.26.1") Then
'  VAR_DownloadFileURL = "http://172.26.1.41/"
'End If 

'ph - 20090714 - start
'Разрешаем доступ к файлам документов некоторым пользователям
If Session("SIT_UserCanDownloadFiles") Then
  VAR_ReadAccess = "Y"
End If
'ph - 20090714 - end
%>