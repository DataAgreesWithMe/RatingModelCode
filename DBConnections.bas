Attribute VB_Name = "DBConnections"
'################################################################################################################################################################################################
'
'These modules contains the syntax to open up a sql connection and convert the table into an excel array
'
'######################################################################################################################################################################################################################


Option Base 1

'Execute Sql Script
Public Sub ExecuteSQL(SQLStr As String)
    If cn Is Nothing Then Call openDB(Live_Server_Name, Live_Database_Name)
    cn.CommandTimeout = 0
    cn.Execute SQLStr
End Sub
 
'Get recordset
Public Function getRecordSet(SQLStr As String) As ADODB.Recordset
    If cn Is Nothing Then Call openDB(Live_Server_Name, Live_Database_Name)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient 'adUseServer  'adUseClient 'adUseServer  'adUseClient ''adUseServer
    rs.Open SQLStr, cn, adOpenForwardOnly 'adOpenStatic 'adOpenForwardOnly
    Set getRecordSet = rs
End Function

'Convert Recordset to array
Public Function getArrayfromRst(rs)
    If rs.EOF = False Then
    getArrayfromRst = rs.GetRows
    Else
    getArrayfromRst = ""
    End If
End Function

'Open database Connection
Public Sub openDB(ServerName, DBName)
    If cn Is Nothing Then
        Set cn = New ADODB.Connection
        cn.Open "Driver={SQL Server};Server=" & ServerName & ";Database=" & DBName & ";"
    End If
End Sub

'Close database Connection
Public Sub closeDB()
    If Not (cn Is Nothing) Then
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
    End If
End Sub






