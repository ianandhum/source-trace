Attribute VB_Name = "ModuleData"
Public DbConnection As ADODB.Connection
Public querySuccess As Boolean

'table names in the database
Public Const dbTaskTable = "[SourceTrace].[dbo].[tb_task]"
Public Const dbProjectTable = "[SourceTrace].[dbo].[tb_project]"

'Connection String
Private Const ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=virtualbox\Code;Initial Catalog=SourceTrace;Data Source=(local)\SQLEXPRESS"

Public Function InitializeConnection() As Boolean

On Error GoTo ConnectionFailed
    'Change the connection String as obtained from datacontrol
    Set DbConnection = New ADODB.Connection
    
    DbConnection.ConnectionString = ConnectionString
    DbConnection.Open ConnectionString
    
    querySuccess = True
    InitializeConnection = True
    
    Exit Function

ConnectionFailed:
    Debug.Print (Err.Description)
    Debug.Print ("Line " & Err.Number)
    querySuccess = False

End Function

Public Sub CleanConnection()
    If (DbConnection.State = adStateOpen) Then DbConnection.Close
    Set DbConnection = Nothing
End Sub


'Generic function to run a sql query with current connection
Public Function runQuery(query As String) As Recordset
    If DbConnection.State = adStateOpen Then
        Dim result As New Recordset
        With result
            .ActiveConnection = DbConnection
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Source = query
            .Open
        End With
        
        Set runQuery = result
    End If
End Function
Public Function selectFromDB(tableName As String, condition As String) As Recordset
On Error GoTo commandFailed
        Set selectFromDB = runQuery("SELECT * FROM " & tableName & " WHERE " & condition)
        querySuccess = True
        Exit Function
commandFailed:
        Debug.Print ("SelectFromDB")
        Debug.Print (Err.Description)
        querySuccess = False
End Function



Public Function updateSingleColumn(Col As String, value As String, condition As String, dataType As Integer, tableName As String) As Integer
    If DbConnection.State = adStateOpen Then
        Dim result As New Recordset
        Dim paramValue As ADODB.Parameter
        Dim query As String
        Dim runCmd As ADODB.Command
        Dim affectedRecords As Integer
        
        Set runCmd = New ADODB.Command
        
        query = "UPDATE " & tableName & " SET " & Col & " = ? WHERE " & condition
        
        Set paramValue = runCmd.CreateParameter()
        paramValue.value = value
        paramValue.Size = Len(value)
        paramValue.Type = dataType
        
        
        With runCmd
            .ActiveConnection = DbConnection
            .CommandText = query
            .CommandType = adCmdText
            .Parameters.Append paramValue
            .Execute affectedRecords
        End With
        
        updateSingleColumn = affectedRecords
        
        
    End If
End Function
Public Function updateDB(cols() As String, values() As String, dataTypes() As Integer, tableName As String, Optional condition As String = " 1=1 ") As Integer
    If DbConnection.State = adStateOpen Then
        Dim result As New Recordset
        Dim paramValue As ADODB.Parameter
        Dim query As String
        Dim runCmd As ADODB.Command
        Dim affectedRecords As Integer
        Set runCmd = New ADODB.Command
        Dim colList As String
        Dim valCount As String
        colList = ""
        i = 0
        While (i <= UBound(cols) And cols(i) <> "")
            colList = colList & cols(i)
            colList = colList & " = ?  ,"
            i = i + 1
        Wend
        colList = Left(colList, Len(colList) - 1)
        
        query = "UPDATE  " & tableName & " SET " & colList & " WHERE " & condition
        
        Debug.Print query
        
        For j = 0 To i - 1
            Set paramValue = runCmd.CreateParameter()
            paramValue.value = values(j)
            paramValue.Size = Len(values(j))
            paramValue.Type = dataTypes(j)
            runCmd.Parameters.Append paramValue
            
            
        Next j
        
        With runCmd
            .ActiveConnection = DbConnection
            .CommandText = query
            .CommandType = adCmdText
            .Execute affectedRecords
        End With
        
        updateDB = affectedRecords
        
        
    End If
End Function



Public Function deleteFromDB(tableName As String, condition As String) As Integer
    If DbConnection.State = adStateOpen Then
        Dim result As New Recordset
        Dim paramValue As ADODB.Parameter
        Dim query As String
        Dim runCmd As ADODB.Command
        Dim affectedRecords As Integer
        
        Set runCmd = New ADODB.Command
        
        query = "DELETE FROM " & tableName & " WHERE " & condition
        
        
        With runCmd
            .ActiveConnection = DbConnection
            .CommandText = query
            .CommandType = adCmdText
            .Execute affectedRecords
        End With
        
        deleteFromDB = affectedRecords
    End If
End Function



Public Function insertToDB(cols() As String, values() As String, dataTypes() As Integer, tableName As String) As Integer
    If DbConnection.State = adStateOpen Then
        Dim result As New Recordset
        Dim paramValue As ADODB.Parameter
        Dim query As String
        Dim runCmd As ADODB.Command
        Dim affectedRecords As Integer
        Set runCmd = New ADODB.Command
        Dim colList As String
        Dim valCount As String
        colList = ""
        i = 0
        While (i <= UBound(cols) And cols(i) <> "")
            colList = colList & cols(i)
            colList = colList & " ,"
            i = i + 1
        Wend
        colList = Left(colList, Len(colList) - 1)
        
        For j = 0 To i - 1
            valCount = valCount & "? ,"
        Next j
        valCount = Left(valCount, Len(valCount) - 1)
        
        
        query = "INSERT INTO  " & tableName & " ( " & colList & " ) VALUES ( " & valCount & " ) "
        
        Debug.Print query
        
        For i = 0 To j - 1
            Set paramValue = runCmd.CreateParameter()
            paramValue.value = values(i)
            paramValue.Size = Len(values(i))
            paramValue.Type = dataTypes(i)
            runCmd.Parameters.Append paramValue
            
            
        Next i
        
        With runCmd
            .ActiveConnection = DbConnection
            .CommandText = query
            .CommandType = adCmdText
            .Execute affectedRecords
        End With
        
        insertToDB = affectedRecords
        
        
    End If
End Function


