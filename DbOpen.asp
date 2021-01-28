<%
Dim CN
Dim RS
Dim COM

Dim Fs
Dim Connect
Connect = 0
Set Fs = Server.CreateObject("Scripting.FileSystemObject")
set CN = Server.CreateObject("ADODB.Connection")
CN.Mode = 3
'CN.Open "Driver={Microsoft ODBC for Oracle};CONNECTSTRING=SHIO;UID=shioid;PWD=;shioid"
'CN.Open "Driver={Oracle in xe};" & _
'             "CONNECTSTRING=XE; UID=shioid; PWD=shioid;"
CN.Open "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User ID=shioid;Password=shioid;"
If Err.number <> 0 then
		DBConnect = Err.number
		set dbCN = Nothing
	End If
	set COM = Server.CreateObject("ADODB.Command")
	set RS = Server.CreateObject("ADODB.RecordSet")
	set COM.ActiveConnection = CN
	If Err.number <> 0 then
		Connect = Err.number
		set CN = Nothing
	End If


'スキーマ名をsqlで修正する方法
SQL = "ALTER SESSION SET CURRENT_SCHEMA = SHIO"
CN.Execute SQL
If Err.Number <> 0 Then
	Response.Write "スキーマ名変更失敗。" & Err.number &  & Err.description &SQL
	Response.End
	CN.Close
	set CN = Nothing
End If
End Function

%>
