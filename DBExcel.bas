Attribute VB_Name = "DBExcel"
Dim curFormat As String 'Current Cell Format
Dim addtoColsTemp As String 'fields to add
Dim addtoValsTemp As String 'values to add to fields
Dim SQLStr As String 'Container for SQL String

'Creates all ratings tables
Public Sub CreateAllTables()
    CreateLogTable
    CreatePolicyListTable
    CreateAllRatingTables
End Sub

'Loops through all of the tables in SQL and all of the corresponding tables in excel and adds tables into sql if there is a missing table
Public Function CreateExtraTables(StringofSchedules As String, WhereField As String, Schema As String) As Boolean
 Call CreateLogTable
Dim SQLListofTablesArr, j

SQLListofTablesArr = getArrayfromRst(getRecordSet("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '" & Schema & "'"))
ExcelListofTablesArr = Split(StringofSchedules, ",")

For j = 0 To UBound(ExcelListofTablesArr)
    If IsError(Application.Match(ExcelListofTablesArr(j), SQLListofTablesArr, 0)) Then
    ' table
    Call createTable((ExcelListofTablesArr(j)), (Schema & "." & ExcelListofTablesArr(j)), WhereField)
    End If
Next j

End Function

'###ADD SCHEMA to the DB###
Public Sub AddSchema()

    Dim SQLStr2 As String
    
    SQLStr = "CREATE SCHEMA " & RatingSchema
    SQLStr2 = "CREATE SCHEMA " & ParamSchema

    ExecuteSQL (SQLStr)
    ExecuteSQL (SQLStr2)

    ExecuteSQL (SQLStr)
    ExecuteSQL (SQLStr2)

End Sub

'###Creates policy list table###'
Public Sub CreatePolicyListTable()
    
    SQLStr = "CREATE TABLE " & RatingSchema & ".PolicyList" & _
            "(PolicyNo int" & _
            ",[PortfolioName] nvarchar(255)" & _
            ",[SectionRef] nvarchar(255)" & _
            ",[Underwriter] nvarchar(255)" & _
            ",[InceptionDate] datetime2(0)" & _
            ",[YoA] int" & _
            ",[RiskName] nvarchar(255)" & _
            ",[WorkflowStatus] nvarchar(255)" & _
            ",[RiskStatus] nvarchar(255)" & _
            ",[DateCreated] datetime2(0)" & _
            ",[UserCreated] nvarchar(255)" & _
            ",[DateLastSaved] datetime2(0)" & _
            ",[UserLastSaved] nvarchar(255)" & _
            ",[ParameterNo] int" & _
            ",[ExpiryingPolicyNo] int" & _
            ",[ModelVersion] Numeric(10,5)" & _
            ",[DeletePolicyNo] varchar(20))"

    ExecuteSQL (SQLStr)

End Sub
'###This updates the policy list table in the table in the database, it first checks as to whether the Policyno exisits in the policy list table
'if it exist it will prompt the user to see if they would like to overwrite the risk
''then it checks to see if the assured name has changed, if changed with prompt the user if they want to overwrite the risk
' the entries which the table are populated with are picked up with named ranges in the workbook.

Public Sub UpdatePolicyListTable()

    Dim msgResult As VbMsgBoxResult, msgResult1 As VbMsgBoxResult, msgResult2 As VbMsgBoxResult

'Select Policy No from PolicyList to check if it exists
    SQLStr = "SELECT * FROM " & RatingSchema & ".PolicyList WHERE PolicyNo = " & Policy_No
    
    Set rs = getRecordSet(SQLStr)
    
    'If Policy exists prompt user to select whether they want to overwrite
    If rs.RecordCount > 0 Then
        msgResult = MsgBox("PolicyNo " & Policy_No & " already exists.  Overwrite it?", vbOKCancel + vbExclamation, APP_NAME)
        
        If msgResult = vbCancel Then
            MsgBox "Case not saved", vbExclamation, APP_NAME
            closeDB
            EndMacro
            End
        Else
            SQLStr = "SELECT PortfolioName FROM " & RatingSchema & ".PolicyList WHERE PolicyNo = " & Policy_No
            CurrentPortfolio = getArrayfromRst(getRecordSet(SQLStr))
            
            
            If CurrentPortfolio(0, 0) <> Range("SelectionsC_NameOfPortfolio") Then 'check against assured name and sections reference
            
                msgResult1 = MsgBox("You will OVERWRITE " & vbNewLine + vbNewLine & CurrentPortfolio(0, 0) & vbNewLine + vbNewLine & _
                " with " & vbNewLine + vbNewLine & Range("SelectionsC_NameOfPortfolio") & vbNewLine + vbNewLine & "are you sure you want to continue?", vbOKCancel + vbCritical)
                
                If msgResult1 = vbCancel Then
                     msgResult2 = MsgBox("Do you want to save this as a new Policy Number?", vbOKCancel + vbInformation)
                    If msgResult2 = vbCancel Then
                            closeDB
                            MsgBox "Risk has not been saved!", vbOKOnly + vbExclamation
                            EndMacro
                            End
                    Else
                        Policy_No = GetNextPolicyNo
                        Range(RatingPrefix & "_Policy_No") = Policy_No
                        Set rs = getRecordSet("SELECT * FROM " & RatingSchema & ".PolicyList WHERE PolicyNo = " & Policy_No)
                    End If
                End If
            End If
        End If
    End If
   
   
    
    If rs.RecordCount = 0 Then
            SQLStr = "INSERT INTO " & RatingSchema & ".PolicyList (PolicyNo, DateCreated, UserCreated) VALUES (" & _
                 Policy_No & ", '" & Format(Now(), "dd/mmm/yyyy") & "', '" & Environ("USERNAME") & "')"
            ExecuteSQL (SQLStr)
    End If
            SQLStr = "UPDATE " & RatingSchema & ".PolicyList SET " & _
            "PolicyNo = '" & Policy_No & "'" & _
            ",[PortfolioName] ='" & Range("SelectionsC_NameOfPortfolio") & "'" & _
            ",[SectionRef] ='" & "NULL" & "'" & _
            ",[Underwriter] =" & "NULL" & "" & _
            ",[InceptionDate] =" & "NULL" & "" & _
            ",[YoA] =" & "NULL" & "" & _
            ",[RiskName] ='" & "NULL" & "'" & _
            ",[WorkflowStatus] ='" & "NULL" & "'" & _
            ",[RiskStatus] ='" & "NULL" & "'" & _
            ",[DateLastSaved] ='" & Format(Now(), "dd/mmm/yyyy hh:mm:ss") & "'" & _
            ",[UserLastSaved] ='" & Environ("username") & "'" & _
            ",[ParameterNo] =" & Range(ParameterPrefix & "_ParameterNo") & _
            ",[ExpiryingPolicyNo] =" & IIf(Range(RatingPrefix & "_Policy_No_" & ExpiringSuffix) = "", "Null", Range(RatingPrefix & "_Policy_No_" & ExpiringSuffix)) & _
            ",[ModelVersion] =" & Range("Database_ModelVersion") & _
            " WHERE PolicyNo = " & Policy_No
            
            ExecuteSQL (SQLStr)
    
End Sub

'this is a one off procedure which creates the parameters table initially in the sql database, this is called when the make parameter tables button is pressed
Public Sub CreateParameterTable()
  
  SQLStr = "CREATE TABLE " & ParamSchema & ".ParametersTable" & _
     "(ParameterNo INT," & _
     "ClassOfBusiness varchar(20)," & _
     "UnderwritingTeam varchar(20)," & _
     "VersionDate Date," & _
     "UserLastSaved varchar(20)," & _
     "DateLastSaved Date," & _
     "DateCreated Date," & _
     "UserCreated varchar(20))"

      ExecuteSQL (SQLStr)

End Sub

'updates the main paramters table which stores the date last saved.
Public Sub UpdateParameterTable()
  
    ' Check to see if the parameter exists already
    Param_No = Range(ParameterPrefix & "_ParameterNo")
    Set rs = getRecordSet("SELECT * FROM " & ParamSchema & ".ParametersTable WHERE ParameterNo = " & Param_No)
    
    ' If so, check to see if they mean to overwrite it
    If rs.RecordCount > 0 Then
        msgResult = MsgBox("Parameter No " & Param_No & " already exists.  Overwrite it?", vbOKCancel + vbExclamation, APP_NAME)
    
        If msgResult = vbCancel Then
            MsgBox "Case not saved", vbExclamation, APP_NAME
            closeDB
            Exit Sub
        End If
    End If
  
   If rs.RecordCount = 0 Then
    ' If not, create it
        SQLStr = "INSERT INTO " & ParamSchema & ".ParametersTable " & _
        "(ParameterNo, ClassOfBusiness, UnderwritingTeam, DateCreated, UserCreated) " & _
        "VALUES (" & Param_No & ",'" & _
         Division & "', '" & _
         UW_Team & "', '" & _
         Format(Now(), "dd/mmm/yyyy hh:mm:ss") & "', '" & _
         Environ("USERNAME") & "')"
    ExecuteSQL (SQLStr)
    End If
    
    SQLStr = "UPDATE " & ParamSchema & ".ParametersTable " & _
    "SET VersionDate ='" & VersionDate & "' ," & _
    "DateLastSaved =" & "'" & Format(Now(), "dd/mmm/yyyy hh:mm:ss") & "'," & _
    "UserLastSaved ='" & Environ("username") & "'" & _
    "WHERE ParameterNo= " & Param_No & ";"
    
    ExecuteSQL (SQLStr)

End Sub

'''###CREATE TABLE AND ADD ALL COLUMNS###'''
Public Sub createTable(Prefix As String, TableName As String, WhereField As String)
    
    ExecuteSQL ("CREATE TABLE " & TableName & "(" & WhereField & " INT)")
    strTemp = "CREATE TABLE " & TableName & "(" & WhereField & " INT)"
    ExecuteSQL ("INSERT INTO " & RatingSchema & ".Log " & " WITH (TABLOCK) (SQLScriptUsed, DateCreated, UserCreated,TableName,Command) " & _
                        "VALUES ('" & strTemp & "','" & Format(Now(), "dd/mmm/yyyy hh:mm:ss") & "','" & Environ("username") & "','" & TableName & "','Add Table');")
                        
    ExecuteSQL ("CREATE INDEX Index_" & WhereField & " ON " & TableName & " (" & WhereField & ")")
    strTemp = "CREATE INDEX Index_" & WhereField & " ON " & TableName & " (" & WhereField & ")"
    ExecuteSQL ("INSERT INTO " & RatingSchema & ".Log " & " WITH (TABLOCK) (SQLScriptUsed, DateCreated, UserCreated,TableName,FieldName,Command) " & _
                        "VALUES ('" & strTemp & "','" & Format(Now(), "dd/mmm/yyyy hh:mm:ss") & "','" & Environ("username") & "','" & TableName & "','" & Column & "','Create Index');")
                        
    UpdateAllColumns Prefix, TableName, WhereField
    
End Sub

'''###DELETE ENTRY TO SQL TABLE###'''
Public Sub deleteEntry(TableName As String, WhereField As String, WhereValue As Double)
   ExecuteSQL ("DELETE FROM " & TableName & " WHERE " & WhereField & " = " & WhereValue)
End Sub

'''###ADD FIELD TO SQL TABLE###'''
'checks the data type of the cell first and uses this to define the data type for the field in the sql table
Public Sub UpdateAllColumns(Prefix As String, TableName As String, WhereField As String)
    
    Dim curColumns As Collection, colFormats As Collection, actualColumns As Collection
    Dim colName As String, strTemp As String
    
    Dim curName As Name
    Dim Length As Double
    Dim Column As Variant
    Dim fld As ADODB.Field
    
    Length = Len(Prefix)
    
    Set curColumns = New Collection
    Set colFormats = New Collection
    Set actualColumns = New Collection
    
    
    Set rs = getRecordSet("SELECT top 1 * FROM " & TableName)
    
    For Each fld In rs.Fields
        If fld.Name <> WhereField Then actualColumns.Add fld.Name
    Next
    
    For Each curName In Names
       
    
        If Len(curName.Name) <> Length And InStr(1, curName.Name, Prefix & "_") = 1 Then
            If DoesItemExist(actualColumns, Right(curName.Name, Len(curName.Name) - Length - 1)) = False Then
                curFormat = curName.RefersToRange.NumberFormat
                
                If Right(curName.Name, Len(curName.Name) - Length - 1) = "OrderNumber" Then
                    curFormat = "decimal(30,15);"
                ElseIf curFormat = "0%" Then
                    curFormat = "varchar(250) null;"
                ElseIf curFormat = "General" And ((curName.RefersToRange.ColumnWidth * curName.RefersToRange.RowHeight) > 500 Or curName.RefersToRange.MergeArea.Cells.Count > 3) Then
                    curFormat = "varchar(max) null;"
                ElseIf curFormat = "General" Then
                    curFormat = "varchar(250) null;"
                ElseIf curFormat Like "*[m/d/yyyy]*" Then
                    curFormat = "Date;"
                ElseIf curFormat Like "*[.%/]*" Then
                    curFormat = "decimal(30,15);" 'need to expand out datatypes
                ElseIf curFormat = "###,###,###,###" Then
                    curFormat = "varchar(250) null;"
                ElseIf curFormat Like "*[#0]*" Then
                    curFormat = "decimal(30,15);" 'need to expand out datatypes
                Else
                    curFormat = "varchar(250) null;"
                End If
                colName = Right(curName.Name, Len(curName.Name) - Length - 1)
                curColumns.Add colName
                colFormats.Add curFormat, colName
            End If
        End If
    Next
    
    Call CreateLogTable
    
    For Each Column In curColumns
        On Error Resume Next
       ' strTemp = ""
       ' strTemp = actualColumns.Item(Column)
        
       ' If strTemp = "" Then
           ' strTemp = "ALTER TABLE " & TableName & " ALTER COLUMN [" & Column & "] " & colFormats.Item(Column) 'alters the data column type
            strTemp = "ALTER TABLE " & TableName & " ADD [" & Column & "] " & colFormats.item(Column)
            ExecuteSQL (strTemp)
            ExecuteSQL ("INSERT INTO " & RatingSchema & ".Log " & " WITH (TABLOCK) (SQLScriptUsed, DateCreated, UserCreated,TableName,FieldName,DataType,Command) " & _
                        "VALUES ('" & strTemp & "','" & Format(Now(), "dd/mmm/yyyy hh:mm:ss") & "','" & Environ("username") & "','" & TableName & "','" & Column & "','" & Left(colFormats.item(Column), Len(colFormats.item(Column)) - 1) & "', 'Add Column');")
                
       ' End If
    On Error GoTo 0
   Next
End Sub

Sub CreateLogTable()
          SQLStr = "SET NOCOUNT ON; " & _
                "IF NOT EXISTS (SELECT 'SQLScriptUsed' " & _
                   "FROM   INFORMATION_SCHEMA.TABLES " & _
                   "WHERE  TABLE_NAME = 'Log' " & _
                           "AND TABLE_SCHEMA = '" & RatingSchema & "') " & _
                "BEGIN " & _
                "CREATE TABLE " & RatingSchema & ".Log " & _
                    "([SQLScriptUsed] nvarchar(max)" & _
                    ",[Notes] nvarchar(max)" & _
                    ",[DataType] nvarchar(max)" & _
                    ",[TableName] nvarchar(255)" & _
                    ",[FieldName] nvarchar(255)" & _
                    ",[Command] nvarchar(255)" & _
                    ",[DateCreated] datetime2(0)" & _
                    ",[UserCreated] nvarchar(255))" & _
                " END"
           
           ExecuteSQL (SQLStr)
End Sub

Public Function DoesItemExist(mySet As Collection, myCheck As String) As Boolean
    DoesItemExist = False
    For Each elm In mySet
        If myCheck = elm Then
            DoesItemExist = True
            Exit Function
        End If
    Next
End Function


Private Function addToVals2(off_set_rows As Double, off_set_cols As Double, ResultsArr As Variant) As String
    On Error Resume Next
    If Not (IsEmpty(ResultsArr(1 + off_set_rows, 1)) Or ResultsArr(1 + off_set_rows, 1 + off_set_cols) = "") Then
        If curFormat Like "*[m/d/yyyy]*" Then
            'value is a date - need to format it carefully
            addToVals2 = "'" & Format(ResultsArr(1 + off_set_rows, 1 + off_set_cols), "dd mmm yyyy") & "'"
            '##### If formatted as Interest, then treat as a string
        ElseIf curFormat = "[>0]0.0%;@" Or curFormat = "[>0]0.00%;@" Or curFormat = "###,###,###,###" Or curFormat = "0%" Then
            addToVals2 = "'" & Replace(ResultsArr(1 + off_set_rows, 1 + off_set_cols), "'", "''") & "'"
           '#####
        ElseIf curFormat Like "*[.%/#0]*" Then
               'value is a number - don't need to do anything with it
            addToVals2 = ResultsArr(1 + off_set_rows, 1 + off_set_cols)
        Else
            'value is a string - need to put quotes around it
            ' and escape any quotes within it
            addToVals2 = "'" & Replace(ResultsArr(1 + off_set_rows, 1 + off_set_cols), "'", "''") & "'"
        End If
    Else
        addToVals2 = "NULL"
    End If
    On Error GoTo 0
End Function

Private Function addToVals4(off_set_rows As Double, off_set_cols As Double, ResultsArr As Variant) As String
    On Error Resume Next
    If Not (IsEmpty(ResultsArr(1 + off_set_rows, 1)) Or ResultsArr(1 + off_set_rows, 1 + off_set_cols) = "") Then
        addToVals4 = "'" & Replace(ResultsArr(1 + off_set_rows, 1 + off_set_cols), "'", "''") & "'"
    Else
        addToVals4 = "NULL"
    End If
    On Error GoTo 0
End Function


Private Function addToVals5(curName As Name, off_set As Double) As String
   
    On Error Resume Next
    If Not (IsEmpty(curName.RefersToRange.Offset(off_set, 0).Formula) Or curName.RefersToRange.Offset(off_set, 0).Formula = "") Then
    
    
    addToVals5 = "'" & Replace(curName.RefersToRange.Offset(off_set, 0).Formula, "'", "''") & "'"
    
            
    Else
        addToVals5 = ""
    End If
    On Error GoTo 0
End Function



Private Function addToVals(curName As Name, off_set As Double, TempStor) As String
   
    On Error Resume Next
    If Not (IsEmpty(TempStor) Or TempStor = "") Then
        curFormat = curName.RefersToRange.NumberFormat
        If curFormat Like "*[m/d/yyyy]*" Then
            'value is a date - need to format it carefully
            addToVals = "'" & Format(TempStor, "dd mmm yyyy") & "'"
            '##### If formatted as Interest, then treat as a string
        ElseIf curFormat = "[>0]0.0%;@" Or curFormat = "[>0]0.00%;@" Or curFormat = "###,###,###,###" Or curFormat = "0%" Then
            addToVals = "'" & Replace(TempStor, "'", "''") & "'"
           '#####
        ElseIf curFormat Like "*[.%/#0]*" Then
               'value is a number - don't need to do anything with it
            addToVals = TempStor
        Else
            'value is a string - need to put quotes around it
            ' and escape any quotes within it
            addToVals = "'" & Replace(TempStor, "'", "''") & "'"
        End If
    Else
        addToVals = ""
    End If
    On Error GoTo 0
End Function




Private Function addToForVals(curName As Name, off_set As Double, TempStor) As String
   
    On Error Resume Next
    If Not (IsEmpty(TempStor) Or TempStor = "") Then
            addToForVals = "'" & Replace(TempStor, "'", "''") & "'"
    Else
        addToForVals = ""
    End If
    On Error GoTo 0
End Function


Private Function addToVals3(curName As Name, off_set As Double) As String
   
    On Error Resume Next
    If Not (IsEmpty(curName.RefersToRange.Offset(off_set, 0).Value) Or curName.RefersToRange.Offset(off_set, 0).Value = "") Then
        curFormat = curName.RefersToRange.NumberFormat
        If curFormat Like "*[m/d/yyyy]*" Then
            'value is a date - need to format it carefully
            addToVals3 = "'" & Format(curName.RefersToRange.Offset(off_set, 0).Value, "dd mmm yyyy") & "'"
            '##### If formatted as Interest, then treat as a string
        ElseIf curFormat = "[>0]0.0%;@" Or curFormat = "[>0]0.00%;@" Or curFormat = "###,###,###,###" Or curFormat = "0%" Then
            addToVals3 = "'" & Replace(curName.RefersToRange.Offset(off_set, 0).Value, "'", "''") & "'"
           '#####
        ElseIf curFormat Like "*[.%/#0]*" Then
               'value is a number - don't need to do anything with it
            addToVals3 = curName.RefersToRange.Offset(off_set, 0).Value
        Else
            'value is a string - need to put quotes around it
            ' and escape any quotes within it
            addToVals3 = "'" & Replace(curName.RefersToRange.Offset(off_set, 0).Value, "'", "''") & "'"
        End If
    Else
        addToVals3 = "NULL"
    End If
    On Error GoTo 0
End Function

'''###ADD COLUMNS TO SQL TABLE###'''
Private Function addToCols(curName As Name, off_set As Double, Length As Double, TempStor) As String
    On Error Resume Next
    If Not (IsEmpty(TempStor) Or TempStor = "") Then
        addToCols = Right(curName.Name, Len(curName.Name) - Length - 1)
    Else
        addToCols = ""
    End If
    On Error GoTo 0
End Function

'''###SAVE TABLE###'''
Public Sub saveTable(Prefix As String, TableName As String, WhereField As String, WhereValue As Double)
    'this creates a new table with a column for each named cell in the spreadsheet with the prefix
    'each table also has the PolicyNo
    
    Dim thisCol As String, thisVal As String
    Dim curName As Name, Length As Double
      
    Static storedNames() As Collection
    Static storedPrefix As Collection
    Dim maxName As Double
    Dim curPrefix As Double
    
    If storedPrefix Is Nothing Then Set storedPrefix = New Collection
    
    curPrefix = -1
    On Error Resume Next
    maxName = UBound(storedNames, 1)
    curPrefix = storedPrefix.item(Prefix)
    On Error GoTo 0
    
    Length = Len(Prefix)
    
    If curPrefix = -1 Then
        
        curPrefix = maxName + 1
        storedPrefix.Add curPrefix, Prefix
        ReDim Preserve storedNames(curPrefix)
        Set storedNames(curPrefix) = New Collection
        
        For Each curName In Names
            If InStr(1, curName.Name, Prefix & "_") = 1 Then
                storedNames(curPrefix).Add curName
            End If
        Next
        
    End If
          
    ExecuteSQL ("INSERT INTO " & TableName & " (" & WhereField & ") VALUES (" & WhereValue & ");")
    
    For Each curName In storedNames(curPrefix)
    
      TempStor = curName.RefersToRange.Offset(0, 0).Value
        
        thisCol = addToCols(curName, 0, Length, TempStor)
        thisVal = addToVals(curName, 0, TempStor)
        
        If thisCol <> "" And thisVal <> "" Then
            ExecuteSQL ("UPDATE " & TableName & " SET " & thisCol & " = " & thisVal & " WHERE " & WhereField & " = " & WhereValue & ";")
            
        End If
    Next

End Sub

Function ValidateEntry(ExcelEntry, SQLDataType As String, SQLLength As Single)


End Function


'''###SAVE SCHEDULE###'''
Public Sub saveSchedule(Prefix As String, TableName As String, NoRows As Variant, WhereField As String, WhereValue As Double, RowsOrCols As String)
    'this creates a new table with a column for each named cell in the spreadsheet with the prefix
    'each table also has the policy no
    
    Dim cols As String, vals As String
    Dim curName As Name ' current named range
    Dim j As Double
    Dim maxName As Double, curPrefix As Double, Length As Double
    
    If NoRows = 0 Then Exit Sub
    
    Static storedNames() As Collection ' used to store all of the name ranges, related to the given prefix
    Static storedPrefix As Collection
   
    If storedPrefix Is Nothing Then Set storedPrefix = New Collection
    
    curPrefix = -1
    On Error Resume Next
    maxName = UBound(storedNames, 1)
    curPrefix = storedPrefix.item(Prefix)
    On Error GoTo 0
    
    Length = Len(Prefix)
    
    If curPrefix = -1 Then
        curPrefix = maxName + 1
        storedPrefix.Add curPrefix, Prefix
        ReDim Preserve storedNames(curPrefix)
        Set storedNames(curPrefix) = New Collection
        
        For Each curName In Names ' looping through name ranges within all the named ranges in the workbook and if matches against prefix names, adds to the collection
            If Len(curName.Name) <> Length And InStr(1, curName.Name, Prefix & "_") = 1 Then
                storedNames(curPrefix).Add curName ' adds name to the collection
            End If
        Next
    End If
            
        Dim TempArray, DataTypeArray
        TempArray = ""
        DataTypeArray = ""
        
        Dim Counter As Boolean, FieldName As String
        Counter = False
        counter2 = 1
        Dim SheetReultsArray
        ReDim TempArray(0 To (NoRows - 1))
     
    For Each curName In storedNames(curPrefix)
        FieldName = Right(curName.Name, Len(curName.Name) - Length - 1)
    
        nonblank = False
            If RowsOrCols = "Rows" Then '####ROWS###'
                
                SheetReultsArray = curName.RefersToRange.Resize(NoRows).Value2
              
                If NoRows > 1 Then
                     For N = 1 To NoRows
                        If SheetReultsArray(N, 1) <> "" Then
                            nonblank = True
                        End If
                    Next N
                ElseIf NoRows = 1 Then
                    nonblank = True
                End If
                
                If Not (nonblank) Then GoTo NextCurName:
                
                
            ElseIf RowsOrCols = "Cols" Then '####COLUMNS###'
                SheetReultsArray = curName.RefersToRange.Resize(1, NoRows).Value2
                
                If NoRows > 1 Then
                    For N = 1 To NoRows
                        If SheetReultsArray(1, N) <> "" Then
                            nonblank = True
                        End If
                    Next N
                ElseIf NoRows = 1 Then
                    nonblank = True
                End If
                
                If Not (nonblank) Then GoTo NextCurName:
            ElseIf RowsOrCols = "Cell" Then '####CEllS###'
                SheetReultsArray = curName.RefersToRange.Value2
                If SheetReultsArray = "" Or IsEmpty(SheetReultsArray) Then GoTo NextCurName:
            End If
                
           curFormat = curName.RefersToRange.NumberFormat
            

            For j = 0 To NoRows - 1
                If NoRows > 1 Then
                        If RowsOrCols = "Rows" Then
                            addtoValsTemp = addToVals2(j, 0, SheetReultsArray)
                        ElseIf RowsOrCols = "Cols" Then
                            addtoValsTemp = addToVals2(0, j, SheetReultsArray)
                        End If
                Else
                         
                     addtoValsTemp = addToVals3(curName, j)

                End If
                
                If Not (Counter) Then
                    TempArray(j) = addtoValsTemp
                Else
                    TempArray(j) = TempArray(j) & "," & addtoValsTemp
                End If
                
            Next j
              
            addtoColsTemp = ""
            addtoValsTemp = ""
            
            addtoColsTemp = FieldName & ","
            cols = cols & addtoColsTemp
            
        Counter = True
NextCurName:
        vals = ""
         
     Next
           If cols = "" Then Exit Sub
            cols = WhereField & "," & Left(cols, Len(cols) - 1)
     
           For j = 0 To NoRows - 1
                TempArray(j) = "(" & WhereValue & "," & TempArray(j) & ")"
           Next j
                       
        Dim TempArray2
           'Calculate how many loops of 25 '(max upload to sql and how many in residual loop
            x = ((NoRows) - (NoRows Mod 25)) / 25
        Dim FirstLoop
        FirstLoop = 1
           
           If (NoRows) <= 25 Then
                    vals = Join(TempArray, ",")
                    vals = Left(vals, Len(vals) - 1)
                    ExecuteSQL ("INSERT INTO " & TableName & " WITH (TABLOCK) (" & cols & ") VALUES " & vals & ");")
           Else
            '''add to value tables
            
                    For Y = 0 To (x - 1)
                            TempArray((25 * (Y + 1)) - 1) = TempArray((25 * (Y + 1)) - 1) & "¬#"
                    Next Y
                    
                    vals = Join(TempArray, ",")
                    If (NoRows Mod 25) = 0 Then vals = Left(vals, Len(vals) - 2)
                    TempArray2 = Split(vals, "¬#")
                    vals = "": SheetReultsArray = "": TempArray = ""
                    ExecuteSQL ("INSERT INTO " & TableName & " WITH (TABLOCK) (" & cols & ") VALUES " & TempArray2(0) & ";")
              
                For Y = 1 To UBound(TempArray2)
                    SQLStr = Right(TempArray2(Y), Len(TempArray2(Y)) - 1)
                    ExecuteSQL ("INSERT INTO " & TableName & " WITH (TABLOCK) (" & cols & ") VALUES " & SQLStr & ";")
                Next Y
                              
           End If
    
End Sub

'''###CLEAR CELLS###'''
Public Sub ClearCells(Prefix As String)
    
    Dim curName As Name
    
    On Error Resume Next
    For Each curName In Names
        If InStr(1, curName.Name, Prefix & "_") = 1 Then
            If curName.RefersToRange.HasFormula = False Then curName.RefersToRange.Value = ""
        End If
        
        
    Next
          
    On Error GoTo 0

End Sub

'''###CLEAR SCHEDULE###'''
Public Sub ClearSchedule(Prefix As String, NoRows As Double, noCols As Double)
    Dim curName As Name
       On Error Resume Next
    For Each curName In Names
        If InStr(1, curName.Name, Prefix & "_") = 1 Then
            If Not (curName.RefersToRange.MergeCells) Then
                If curName.RefersToRange.HasFormula = False Then curName.RefersToRange.Resize(NoRows, noCols).ClearContents
            Else
                If curName.RefersToRange.HasFormula = False Then curName.RefersToRange.MergeArea.Resize(NoRows).ClearContents
            End If
                
        End If
    Next
        On Error GoTo 0
End Sub

'''###Clear Formula###'''
Public Sub ClearFormula()

Dim FormulaNamRng(), Identifier(), MaxNoEntries(), TableName(), Cell, i
ReDim FormulaNamRng(1 To Range("FormulaR")), Identifier(1 To Range("FormulaR")), MaxNoEntries(1 To Range("FormulaR")), TableName(1 To Range("FormulaR"))

FormulaNamRng = Application.Transpose(Range("Formula_NamedRanges").Resize(Range("FormulaR")))
Identifier = Application.Transpose(Range("Formula_Identifier").Resize(Range("FormulaR")))
MaxNoEntries = Application.Transpose(Range("Formula_MaxNoEntries").Resize(Range("FormulaR")))
TableName = Application.Transpose(Range("Formula_TableName").Resize(Range("FormulaR")))

    For i = 1 To UBound(FormulaNamRng)
        
        If MaxNoEntries(i) > 0 Then
            If Not (Range(FormulaNamRng(i)).MergeCells) Then
                Range(FormulaNamRng(i)).Resize(IIf(Identifier(i) = "Cell" Or Identifier(i) = "Cols", 1, MaxNoEntries(i)), IIf(Identifier(i) = "Cell" Or Identifier(i) = "Rows", 1, MaxNoEntries(i))).ClearContents
            Else
                Range(FormulaNamRng(i)).MergeArea.Resize(IIf(Identifier(i) = "Cell" Or Identifier(i) = "Cols", 1, MaxNoEntries(i)), IIf(Identifier(i) = "Cell" Or Identifier(i) = "Rows", 1, MaxNoEntries(i))).ClearContents
            End If
        End If
    Next i

End Sub




'''###LOAD TABLE###'''
Public Sub loadTable(Prefix As String, TableName As String, WhereField As String, WhereValue As Double, Optional FieldList As String)
       
    Dim ResultsArray
    Dim columnNames() As String
    Dim numfields As Double, i As Double, j As Double
       
    SQLStr = "SELECT " & IIf(FieldList = "", "*", FieldList) & " FROM " & TableName & " WHERE " & WhereField & " = '" & WhereValue & "';"
    Set rs = getRecordSet(SQLStr)
    
    If rs.RecordCount = 0 Then Exit Sub
    numfields = rs.Fields.Count
    ResultsArray = getArrayfromRst(rs)
    
    ReDim columnNames(numfields)
    rs.MoveFirst
    
    For i = 0 To numfields - 1
        columnNames(i) = rs.Fields(i).Name
    Next
    
    On Error Resume Next

    For i = 0 To numfields - 1
        If columnNames(i) <> WhereField Then
            If ResultsArray(i, 0) <> "" And Range(Prefix & "_" & columnNames(i)).HasFormula = False Then
                Range(Prefix & "_" & columnNames(i)).Value = ResultsArray(i, 0)
            End If
        End If
    Next
  
    On Error GoTo 0
    
End Sub

'''###LOAD SCHEDULE###'''
Public Sub loadSchedule(Prefix As String, TableName As String, WhereField As String, WhereValue As Double, RowsOrCols As String, Optional FieldList As String)
       
    Dim ResultsArray
    Dim columnNames() As String, TempArr() As String
    Dim numRecords As Double, numfields As Double, i As Double

    SQLStr = "SELECT " & IIf(FieldList = "", " * ", FieldList) & " FROM " & TableName & " WHERE " & WhereField & " = " & WhereValue & " ORDER BY ORDERNUMBER;"

    Set rs = getRecordSet(SQLStr)
    
    If rs.RecordCount = 0 Then Exit Sub
    numRecords = rs.RecordCount
    numfields = rs.Fields.Count
    ReDim columnNames(numfields)
    ResultsArray = getArrayfromRst(rs)
    
    rs.MoveFirst
    
    For i = 0 To numfields - 1
        columnNames(i) = rs.Fields(i).Name
    Next
    On Error Resume Next
    For i = 0 To numfields - 1
        If columnNames(i) <> WhereField Then
            If IIf(FieldList = "", Range(Prefix & "_" & columnNames(i)).Offset(0, 0).HasFormula = False, True) Then
       
        Dim TempArray5
        ReDim TempArray5(numRecords - 1)
        
        For j = 0 To numRecords - 1
             If IsNull(ResultsArray(i, j)) Then
                TempArray5(j) = ""
            Else
                TempArray5(j) = ResultsArray(i, j)
            End If
        Next j
    
                If RowsOrCols = "Rows" Then
                    TempArray6 = TransposeArray(TempArray5)
                    Range(Prefix & "_" & columnNames(i)).Resize(numRecords, 1) = TempArray6
                Else
                    Range(Prefix & "_" & columnNames(i)).Resize(1, numRecords) = TempArray5
                End If
            End If
        End If
    Next
    On Error GoTo 0
End Sub

'###Tranposed the array###'
Public Function TransposeArray(MyArray As Variant) As Variant
Dim x As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim TempArray As Variant
    Xupper = 0
    Yupper = UBound(MyArray, 1)
    ReDim TempArray(Yupper, Xupper)
        For Y = 0 To Yupper
            TempArray(Y, x) = MyArray(Y)
        Next Y
    TransposeArray = TempArray
End Function

'''###BUTTON CLEAR ALL CELLS###'''
Public Sub ClearAllCells()

   ClearAllRatingTables
   GetNextPolicyButton

End Sub

'''###BUTTON CREATE NEW RISK###'''
Public Sub CreateNewRisk()
    
    StartMacro
        ClearAllRatingTables
        NewRiskProcedure
        Sheets(RatingSheetName).Activate
    EndMacro

End Sub

Sub NewRiskProcedure()
    
End Sub

'''###BUTTON GET NEXT POLICY NUMBER###'''
Public Sub GetNextPolicyButton()
    Range(RatingPrefix & "_Policy_No").Value = GetNextPolicyNo
End Sub

'''###GET NEXT POLICY PROCEDURE###'''
Public Function GetNextPolicyNo() As Double

    Dim No As Double
    Set rs = getRecordSet("SELECT PolicyNo FROM " & RatingSchema & ".PolicyList ORDER BY PolicyNo DESC;")
    
    If rs.RecordCount < 1 Then
        'new database
         No = 1
    Else
        rs.MoveFirst
        No = rs.Fields(0).Value + 1
    End If
    
    closeDB
    
    GetNextPolicyNo = No
End Function

'''###USERFORM BUTTON LOAD RISK###'''
Public Sub loadRisk(PolicyNo As Double)
StartMacro
     Application.ScreenUpdating = False
    Policy_No = PolicyNo
    'Clear Fomula Cells
    ModelVersionCheck
    
    ClearAllRatingTables
    Range(RatingPrefix & "_Policy_No").Value = Policy_No
    
    Param_No = GetOriginalParameterNo(RatingSchema & ".PolicyList", Policy_No)
    
    LoadAllRatingTables
 
    If SelectRiskAction = "Renew" Then
        loadParameters getLatestParameterNo
'      LoadAllExpiringScheduleTables
        Range(RatingPrefix & "_Policy_No") = GetNextPolicyNo
        RenewProcedure
    Else
        loadParameters Param_No
    End If
    
Unload Control_Form
Sheets("Selections").Activate
End Sub


Public Sub ModelVersionCheck()
    
    Dim No As Double
    Set rs = getRecordSet("SELECT ModelVersion FROM " & RatingSchema & ".PolicyList WHERE PolicyNo = " & Policy_No)
    
        rs.MoveFirst
        No = rs.Fields(0).Value
    If Range("Database_ModelVersion") <> No Then
       Result = MsgBox("This risk has been rated using Model Version " & No & vbNewLine & vbNewLine & "You are loading the risk in Model Version " & _
        Range("Database_ModelVersion") & vbNewLine & vbNewLine & "Are you sure you want to continue?", vbOKCancel + vbExclamation)
        
        If Result = vbCancel Then
            closeDB
            EndMacro
            End
        End If
        
    End If
    closeDB

    
End Sub

Public Sub RenewProcedure()

End Sub

'''###BUTTON SAVE RISK###'''
Public Sub saveRiskButton()
    StartMacro
    If Range(RatingPrefix & "_Policy_No") = "" Then GetNextPolicyButton
    Policy_No = Range(RatingPrefix & "_Policy_No")
    SaveRisk
    EndMacro
End Sub

'''###SAVE RISK PROCEDURE###'''
Public Sub SaveRisk()
    
    If Range("Update_Tables") = True Or Range("Update_Columns") = True Then
        Call UpdatePolicyListTable
        If Range("Update_Tables") = True Then Call CreateExtraTables(MultipleRatingPrefix & "," & RatingSchedules, "PolicyNo", RatingSchema)
        If Range("Update_Columns") = True Then Call UpdateAllRatingColumns
        closeDB
    
    Else
        'Validate all entries
        ValidateAllRatingTables
        closeDB
        Call UpdatePolicyListTable
        If Range("Update_Tables") = True Then Call CreateExtraTables(MultipleRatingPrefix & "," & RatingSchedules, "PolicyNo", RatingSchema)
        If Range("Update_Columns") = True Then Call UpdateAllRatingColumns
        closeDB
   End If

    ' Delete current entries
    DeleteAllRatingEntries
    closeDB
    
    ' And save new entries
    SaveAllRatingTables
    closeDB
    
    curPrefix = ""
    
    MsgBox "Case saved", vbInformation, APP_NAME
    
End Sub

Public Function GetOriginalParameterNo(TableName As String, PolicyNo As Double) As Double

    Dim numfields As Double
    Dim i As Double
       
    SQLStr = "SELECT * FROM " & TableName & " WHERE PolicyNo = " & Policy_No & ";"
    Set rs = getRecordSet(SQLStr)
    
    If rs.RecordCount = 0 Then
        GetOriginalParameterNo = getLatestParameterNo
        Exit Function
    End If
    numfields = rs.Fields.Count
    rs.MoveFirst
    
    For i = 0 To numfields - 1
        If rs.Fields(i).Name = "ParameterNo" Then GetOriginalParameterNo = rs.Fields(i).Value
    
    Next
    
End Function
      
Public Function getLatestParameterNo() As Double

    Dim No As Double
    
    Set rs = getRecordSet("SELECT ParameterNo FROM " & ParamSchema & ".ParametersTable ORDER BY ParameterNo DESC;")
    
    If rs.RecordCount < 1 Then
        No = 1
    Else
        rs.MoveFirst
        No = rs.Fields(0).Value
    End If
    
    closeDB
    
    getLatestParameterNo = No
End Function

Public Function GetNextParameterNo() As Double
    
    Set rs = getRecordSet("SELECT ParameterNo FROM " & ParamSchema & ".ParametersTable ORDER BY ParameterNo DESC;")
    
    If rs.RecordCount < 1 Then
        'new database
        GetNextParameterNo = 1
    Else
        rs.MoveFirst
        GetNextParameterNo = rs.Fields(0).Value + 1
    End If
    
    closeDB
    
End Function
      
Public Sub loadParameters(ParameterNo As Double)
    
    ClearAllParamTables
    Range(ParameterPrefix & "_ParameterNo").Value = ParameterNo
    Param_No = ParameterNo
    LoadAllParamTables
    
End Sub

Public Sub saveParameterButton()
    saveParameter
End Sub

Public Sub saveParameter()

    Call UpdateParameterTable
    If Range("Update_Param_Tables") = True Then Call CreateExtraTables(ParameterPrefix & "," & ParameterSchedules, "ParameterNo", ParamSchema)
    If Range("Update_Param_Columns") = True Then Call UpdateAllParamColumns
    
    ' Delete current entries
    DeleteAllParamEntries
    
    ' And save new entries
    SaveAllParamTables
  
   MsgBox "Parameters Successfully Saved", vbInformation, APP_NAME
    
End Sub

Public Function GetLatestParametersBtn()
StartMacro
    loadParameters getLatestParameterNo
EndMacro
End Function

Public Function GetLatestParameters()
        loadParameters getLatestParameterNo
End Function

Public Function GetOriginalParametersBtn()
    StartMacro
    GetOriginalParameters
    EndMacro
End Function

Public Function GetOriginalParameters()
        Policy_No = Range(RatingPrefix & "_Policy_No").Value
        loadParameters GetOriginalParameterNo(RatingSchema & ".PolicyList", Policy_No)
End Function

Public Sub getNextParameters()
    Range(ParameterPrefix & "_ParameterNo").Value = GetNextParameterNo
End Sub

Public Function GetCurrentParameters()
    loadParameters Range(ParameterPrefix & "_ParameterNo").Value
End Function

'creates all param tables
Public Sub CreateAllParam()
    CreateParameterTable
    CreateAllParamTables
End Sub

'Loads all of this years entries into expiring entries upon renewal, for example if there is 'FIELD1' it will load into 'FIELD1_Expiry'
'Slightly different syntax depending on if loading 'Rows' or 'Cols'
'table means one-to-one tables
Public Sub loadExpiringTable(Prefix As String, TableName As String)
    
    Dim i
    
    SQLStr = "SELECT left(COLUMN_NAME,len(COLUMN_NAME)-7) FROM INFORMATION_SCHEMA.Columns WHERE TABLE_NAME= '" & Prefix & "' and TABLE_SCHEMA = '" & RatingSchema & "' and Column_Name <> 'PolicyNo' and Right(Column_Name,7) ='_" & ExpiringSuffix & "'"
    ColNameExpiry = Application.Transpose(getArrayfromRst(getRecordSet(SQLStr)))
    If rs.RecordCount = 0 Then Exit Sub
    
    If UBound(ColNameExpiry, 1) = 1 Then
        FieldsColNameExpiry = ColNameExpiry(1)
    Else
        FieldsColNameExpiry = Join(Application.Transpose(ColNameExpiry), "],[")
    End If
    
    SQLStr = "SELECT [" & FieldsColNameExpiry & "] from " & TableName & " WHERE PolicyNo = " & Policy_No
    ColNameExpiryValues = getArrayfromRst(getRecordSet(SQLStr))
    On Error Resume Next
    If Not (UBound(ColNameExpiryValues) >= 1) Then Exit Sub
    On Error GoTo 0
    For i = 0 To UBound(ColNameExpiry, 1) - 1
            
            If UBound(ColNameExpiry, 1) = 1 Then
                If Range(Prefix & "_" & ColNameExpiry(i + 1) & "_" & ExpiringSuffix).HasFormula = False Then
                    Range(Prefix & "_" & ColNameExpiry(i + 1) & "_" & ExpiringSuffix).Value = ColNameExpiryValues(i, 0)
                End If
            Else
                If Range(Prefix & "_" & ColNameExpiry(i + 1, 1) & "_" & ExpiringSuffix).HasFormula = False Then
                    Range(Prefix & "_" & ColNameExpiry(i + 1, 1) & "_" & ExpiringSuffix).Value = ColNameExpiryValues(i, 0)
                End If
            End If
    Next
       
End Sub

'Loads all of this years entries into expiring entries upon renewal, for example if there is 'FIELD1' it will load into 'FIELD1_Expiry'
'Slightly different syntax depending on if loading 'Rows' or 'Cols'
'Schedules means one-to-many tables
Public Sub loadExpiringSchedule(Prefix As String, TableName As String, RowsofCols As String)
    
    Dim i
    Dim NumberOfRecords, NumberOfFields
    NumberOfRecords = 0
    NumberOfFields = 0
    
    SQLStr = "SELECT left(COLUMN_NAME,len(COLUMN_NAME)-7) FROM INFORMATION_SCHEMA.Columns WHERE TABLE_NAME= '" & Prefix & "' and TABLE_SCHEMA = '" & RatingSchema & "' and Column_Name <> 'PolicyNo' and Right(Column_Name,7) ='_" & ExpiringSuffix & "'"
    ColNameExpiry = Application.Transpose(getArrayfromRst(getRecordSet(SQLStr)))
    
    If rs.RecordCount = 0 Then Exit Sub
    NumberOfFields = rs.RecordCount
  
    If NumberOfFields = 1 Then
        FieldsColNameExpiry = ColNameExpiry(1)
    Else
        FieldsColNameExpiry = Join(Application.Transpose(ColNameExpiry), "],[")
    End If
    
    SQLStr = "SELECT [" & FieldsColNameExpiry & "] from " & TableName & " WHERE PolicyNo = " & Policy_No
    ColNameExpiryValues = getArrayfromRst(getRecordSet(SQLStr))
    NumberOfRecords = rs.RecordCount
    
    If NumberOfRecords = 0 Then Exit Sub
    
    Dim TempArr
    ReDim TempArr(NumberOfRecords - 1)
    
    For i = 0 To NumberOfFields - 1
        If Range(Prefix & "_" & ColNameExpiry(i + 1, 1) & "_" & ExpiringSuffix).HasFormula = False Then
                For j = 0 To NumberOfRecords - 1
                     If IsNull(ColNameExpiryValues(i, j)) Then
                        TempArr(j) = ""
                    Else
                        TempArr(j) = ColNameExpiryValues(i, j)
                    End If
                Next j
            
            If RowsofCols = "Rows" Then
                TempArray10 = TransposeArray(TempArr)
                Range(Prefix & "_" & ColNameExpiry(i + 1, 1) & "_" & ExpiringSuffix).Resize(NumberOfRecords, 1) = TempArray10
            Else
                Range(Prefix & "_" & ColNameExpiry(i + 1, 1) & "_" & ExpiringSuffix).Resize(1, NumberOfRecords) = TempArr
            End If
                
        End If
    Next i
       
End Sub

'''###DATA VALIDATION CHECKER###'''
Sub DataValidationChecker(TableName As String, NoRows, TableType As String)
   
    Dim Position As Single, FieldName As String, DataTypeArray, Length, DataValue
    
    DataTypeArray = getArrayfromRst(getRecordSet("SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & TableName & "' and TABLE_SCHEMA = '" & RatingSchema & "'"))
    Length = Len(TableName)
       
        For Each curName In Names ' looping through name ranges within all the named ranges in the workbook and if matches against prefix names, adds to the collection
            If Len(curName.Name) <> Length And InStr(1, curName.Name, TableName & "_") = 1 Then
                FieldName = Right(curName.Name, Len(curName.Name) - Length - 1)
                    Position = FindLoop(DataTypeArray, FieldName)
                    DataType = DataTypeArray(1, Position)
                    DataLength = IIf(DataTypeArray(2, Position) = -1, 8000, DataTypeArray(2, Position))
                     If TableType = "Cell" Then
                         DataValue = curName.RefersToRange.Value
                         ErrorCheck = CheckDataEntries(DataValue, DataType, DataLength)
                         If ErrorCheck <> "Passed" Then GoTo ErrorNotSaveMsgBox:
                     ElseIf TableType = "Rows" And NoRows > 0 Then
                         
                         
                         DataValueArray = curName.RefersToRange.Resize(NoRows)
                         If NoRows = 1 Then
                            DataValue = DataValueArray
                            ErrorCheck = CheckDataEntries(DataValue, DataType, DataLength)
                            If ErrorCheck <> "Passed" Then GoTo ErrorNotSaveMsgBox:
                        Else
                         For Each DataValue In DataValueArray
                             ErrorCheck = CheckDataEntries(DataValue, DataType, DataLength)
                             If ErrorCheck <> "Passed" Then GoTo ErrorNotSaveMsgBox:
                         Next DataValue
                         End If
                         
                         
                     ElseIf TableType = "Cols" And NoRows > 0 Then
                         
                         DataValueArray = curName.RefersToRange.Resize(1, NoRows)
                        If NoRows = 1 Then
                            DataValue = DataValueArray
                            ErrorCheck = CheckDataEntries(DataValue, DataType, DataLength)
                            If ErrorCheck <> "Passed" Then GoTo ErrorNotSaveMsgBox:
                        Else
                           
                            For Each DataValue In DataValueArray
                                
                                ErrorCheck = CheckDataEntries(DataValue, DataType, DataLength)
                                If ErrorCheck <> "Passed" Then GoTo ErrorNotSaveMsgBox:
                            Next DataValue
                        End If
                        
                        
                     End If
            End If
        Next
      
Exit Sub

ErrorNotSaveMsgBox:
MsgBox "Error:" & vbNewLine & vbNewLine & "Sheet Name: " & curName.RefersToRange.Parent.Name & vbNewLine & "Cell: " & Range(curName.Name).Offset(j).Address & _
vbNewLine & vbNewLine & vbNewLine & vbNewLine & ErrorCheck & vbNewLine & vbNewLine & " Risk Not Saved!", vbOKOnly + vbCritical

Sheets(curName.RefersToRange.Parent.Name).Activate
curName.RefersToRange.Activate
closeDB
EndMacro

End

End Sub
