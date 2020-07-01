Attribute VB_Name = "DBLoops"
'##########################################################################################################################################################################################
'Module lists out the Several loops procedures to loop through all of the rating tables and paramter tables to carry out (scehdules mean one-to-many procedures)
'For each of the procedures, the list of tables in the Global Constant tab is 'Split' and then looped through

'Creating of the tables

'Clear of the tables in the Spreadsheet
'Clear of the tables in the DB

'Saving into the DB (referred to as update)

'Loading the information from DB to Excel

'##########################################################################################################################################################################################
Dim xPrefix, xPrefix_1 As Variant
Dim i As Double
Dim PRows, RRows As Variant
Dim xParameterSchedules As Variant

'''###CREATE PARAM###'''
Public Function CreateAllParamTables() As Boolean
    
    createTable ParameterPrefix, ParamSchema & "." & ParameterPrefix, "ParameterNo"
    
    xPrefix = Split(ParameterSchedules, ",")
   
    For i = LBound(xPrefix) To UBound(xPrefix)
        createTable (xPrefix(i)), ParamSchema & "." & (xPrefix(i)), "ParameterNo"
    Next
End Function

'''###CLEAR RATING###'''
Public Function ClearAllParamTables() As Boolean
    
    ClearCells (ParameterPrefix)
    
    xPrefix = Split(ParameterSchedules, ",")
    RRows = Split(ParameterSchedulesRows, ",")
    
    For i = LBound(xPrefix) To UBound(xPrefix)
        ClearSchedule (xPrefix(i)), Range((RRows(i))).Value, 1
    Next
    
End Function

'''###LOAD PARAM###'''
Public Function LoadAllParamTables() As Boolean

    loadTable ParameterPrefix, ParamSchema & "." & ParameterPrefix, "ParameterNo", Param_No
    xPrefix = Split(ParameterSchedules, ",")
   
    For i = LBound(xPrefix) To UBound(xPrefix)
        loadSchedule (xPrefix(i)), ParamSchema & "." & (xPrefix(i)), "ParameterNo", Param_No, "Rows"
    Next
End Function

'''###UPDATE PARAM###'''
Public Function UpdateAllParamColumns() As Boolean
    
    UpdateAllColumns ParameterPrefix, ParamSchema & "." & ParameterPrefix, "ParameterNo"
    
    xPrefix = Split(ParameterSchedules, ",")
    For i = LBound(xPrefix) To UBound(xPrefix)
        UpdateAllColumns (xPrefix(i)), ParamSchema & "." & (xPrefix(i)), "ParameterNo"
    Next
End Function

'''###DELETE PARAM###'''
Public Function DeleteAllParamEntries() As Boolean
  
    deleteEntry ParamSchema & "." & ParameterPrefix, "ParameterNo", Param_No
  
    xPrefix = Split(ParameterSchedules, ",")
   
    For i = LBound(xPrefix) To UBound(xPrefix)
        deleteEntry ParamSchema & "." & (xPrefix(i)), "ParameterNo", Param_No
    Next
End Function

'''###SAVE PARAM###'''
Public Function SaveAllParamTables() As Boolean
    
  '  saveSchedule ParameterPrefix, ParamSchema & "." & ParameterPrefix, 1, "ParameterNo", Param_No, "Cell"
    
    
    ' saveSchedule (xPrefix_1(i)), (RatingSchema & "." & xPrefix_1(i)), 1, "PolicyNo", Policy_No, "Cell"
    
    saveTable ParameterPrefix, ParamSchema & "." & ParameterPrefix, "ParameterNo", Param_No
    
    
    xPrefix = Split(ParameterSchedules, ",")
    PRows = Split(ParameterSchedulesRows, ",")
    
    For i = LBound(xPrefix) To UBound(xPrefix)
        saveSchedule (xPrefix(i)), ParamSchema & "." & (xPrefix(i)), Range((PRows(i))).Value, "ParameterNo", Param_No, "Rows"
    Next
    
End Function

'''###CREATE RATING###'''
Public Function CreateAllRatingTables() As Boolean

    xPrefix = Split(RatingSchedules, ",")
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
   
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        createTable (xPrefix_1(i)), RatingSchema & "." & (xPrefix_1(i)), "PolicyNo"
    Next
    
    For i = LBound(xPrefix) To UBound(xPrefix)
        createTable (xPrefix(i)), RatingSchema & "." & (xPrefix(i)), "PolicyNo"
    Next
End Function

'''###LOAD RATING###'''
Public Function LoadAllRatingTables()
    
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    xPrefix = Split(RatingSchedules, ",")
    xPrefix_2 = Split(SchRowOrCol, ",")
    Dim RowOrCol As String
   
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        loadTable (xPrefix_1(i)), RatingSchema & "." & (xPrefix_1(i)), "PolicyNo", Policy_No
    Next
    
    For i = LBound(xPrefix) To UBound(xPrefix)
        loadSchedule (xPrefix(i)), RatingSchema & "." & (xPrefix(i)), "PolicyNo", Policy_No, Range((xPrefix_2(i)))
    Next
    
End Function

'''###UPDATE RATING###'''
Public Function UpdateAllRatingColumns() As Boolean
 
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    xPrefix = Split(RatingSchedules, ",")
    
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        UpdateAllColumns (xPrefix_1(i)), RatingSchema & "." & (xPrefix_1(i)), "PolicyNo"
    Next
    
    For i = LBound(xPrefix) To UBound(xPrefix)
        UpdateAllColumns (xPrefix(i)), RatingSchema & "." & (xPrefix(i)), "PolicyNo"
    Next
    
End Function

'''###DELETE RATING###'''
Public Function DeleteAllRatingEntries() As Boolean
    
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    xPrefix = Split(RatingSchedules, ",")
   
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        deleteEntry (RatingSchema & "." & xPrefix_1(i)), "PolicyNo", Policy_No
    Next
  
    For i = LBound(xPrefix) To UBound(xPrefix)
        deleteEntry (RatingSchema & "." & xPrefix(i)), "PolicyNo", Policy_No
    Next
End Function

'''###SAVE RATING###'''
Public Function SaveAllRatingTables() As Boolean
    Dim RowOrCol As String
    
    xPrefix = Split(RatingSchedules, ",")
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    RRows = Split(RatingSchedulesRows, ",")
    xPrefix_2 = Split(SchRowOrCol, ",")
    xPrefix_3 = Split(TotalSchCols, ",")
    
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        saveSchedule (xPrefix_1(i)), (RatingSchema & "." & xPrefix_1(i)), 1, "PolicyNo", Policy_No, "Cell"
    Next
   
    For i = LBound(xPrefix) To UBound(xPrefix)
        saveSchedule (xPrefix(i)), (RatingSchema & "." & xPrefix(i)), IIf(Range((xPrefix_2(i))) = "Rows", (Range((RRows(i))).Value), Range((xPrefix_3(i)))), "PolicyNo", Policy_No, Range((xPrefix_2(i)))
    Next
    
End Function

'''###VALIDATE RATING###'''
Public Function ValidateAllRatingTables() As Boolean
    Dim RowOrCol As String
    
    xPrefix = Split(RatingSchedules, ",")
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    RRows = Split(RatingSchedulesRows, ",")
    xPrefix_2 = Split(SchRowOrCol, ",")
    xPrefix_3 = Split(TotalSchCols, ",")
    
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        DataValidationChecker (xPrefix_1(i)), 1, "Cell"
    Next
   
    For i = LBound(xPrefix) To UBound(xPrefix)
        DataValidationChecker (xPrefix(i)), IIf(Range((xPrefix_2(i))) = "Rows", (Range((RRows(i))).Value), Range((xPrefix_3(i)))), Range((xPrefix_2(i)))
    Next
    
End Function


'''###CLEAR RATING###'''
Public Function ClearAllRatingTables() As Boolean
    
    xPrefix = Split(RatingSchedules, ",")
    RRows = Split(TotalSchRows, ",")
    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    xPrefix_3 = Split(SchRowOrCol, ",")
   
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        ClearCells (xPrefix_1(i))
    Next
    
    For i = LBound(xPrefix) To UBound(xPrefix)
    
        If Range((xPrefix_3(i))).Value = "Rows" Then
            ClearSchedule (xPrefix(i)), Range((RRows(i))).Value, 1
        Else
            ClearSchedule (xPrefix(i)), 1, Range((RRows(i))).Value
        End If

    Next
    
End Function

'''###LOAD EXPIRING RATING###'''
Public Function LoadAllExpiringScheduleTables() As Boolean

    xPrefix_1 = Split(MultipleRatingPrefix, ",")
    xPrefix = Split(RatingSchedules, ",")
    xPrefix_2 = Split(SchRowOrCol, ",")
    
    For i = LBound(xPrefix_1) To UBound(xPrefix_1)
        loadExpiringTable (xPrefix_1(i)), RatingSchema & "." & (xPrefix_1(i))
    Next
    
    For i = LBound(xPrefix) To UBound(xPrefix)
        loadExpiringSchedule (xPrefix(i)), RatingSchema & "." & (xPrefix(i)), Range((xPrefix_2(i))).Value
    Next

End Function

'''###Opens the database from
Sub Open_Database()
    Dim Command
    Command = Shell("ssms.exe -S " & Live_Server_Name & " -d " & Live_Database_Name & " -E ")
End Sub




