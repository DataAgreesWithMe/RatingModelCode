Attribute VB_Name = "GlobalFunctions"

'StartMacro turns off calculations to speed up macros
Public Sub StartMacro()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

'EndMacro turns on calculations after macro has finished
Public Sub EndMacro()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub




Public Function RemoveDupesColl(MyArray As Variant)
'DESCRIPTION: Removes duplicates from your array using the collection method.
'NOTES: (1) This function returns unique elements in your array, but
'it converts your array elements to strings.
'-----------------------------------------------------------------------
    Dim i As Long
    Dim arrColl As New Collection
    Dim arrDummy() As Variant
    Dim arrDummy1() As Variant
    Dim item As Variant
    ReDim arrDummy1(LBound(MyArray) To UBound(MyArray))

    For i = LBound(MyArray) To UBound(MyArray) 'convert to string
        arrDummy1(i) = CStr(MyArray(i))
    Next i
    On Error Resume Next
    For Each item In arrDummy1
       arrColl.Add item, item
    Next item
    Err.Clear
    ReDim arrDummy(LBound(MyArray) To arrColl.Count + LBound(MyArray) - 1)
    i = LBound(MyArray)
    For Each item In arrColl
       arrDummy(i) = item
       i = i + 1
    Next item
    RemoveDupesColl = arrDummy
End Function

Public Function interpolate(x1 As Double, x2 As Double, y1 As Double, y2 As Double, z As Double)
 
' this function linearly interpolates the function of a value z which lies between
' between x1 and x2. y1 and y2 are the functions of the values x1 and x2
 
interpolate = y1 + ((z - x1) / (x2 - x1)) * (y2 - y1)
 
End Function

Function CheckDataEntries(DataValue, DataType, DataLength)

'Check #ref error
If IsError(DataValue) Then
    CheckDataEntries = "A #Error entry has been found, please amend!": Exit Function
End If

'Dependant on datatype of what is held in sql
'First check if data type is decimal
If DataType = "decimal" Then
    If Not (Application.WorksheetFunction.IsNumber(DataValue)) And (DataValue <> "") Then CheckDataEntries = "A non numeric value has been found in numeric only cell": Exit Function
'Next check if data type is date
ElseIf Left(DataType, 4) = "date" Then
On Error Resume Next
    If (Not IsDate(DataValue) Or Not (Year(DataValue) > 1901)) And (DataValue <> "") Then CheckDataEntries = "The entry is not a valid date, please amend.": Exit Function
On Error GoTo 0
'Next check if varchar data type to check length
ElseIf DataType = "varchar" And DataValue <> "" Then
    If Not (Len(DataValue) <= DataLength) Then CheckDataEntries = "The entry is too long, please consult the actuaries": Exit Function
End If

CheckDataEntries = "Passed"

End Function

Function FindLoop(arr, val) As Single
    Dim r As Long, c As Long
    For r = 0 To UBound(arr, 1)
    For c = 0 To UBound(arr, 2)
        If arr(r, c) = val Then
            FindLoop = c
            Exit Function
        End If
    Next c
    Next r
End Function



