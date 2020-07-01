Attribute VB_Name = "Userform_Commands"
Sub LoadDataToListBox(OrderbyField As String, DescAsc As String, GlobalSearchItem As String)
    'SQL
    Dim SQLStr As String, WhereStr As String
    SQLStr = "SELECT [PolicyNo], [PortfolioName], [RiskName], [SectionRef], [Underwriter], format([InceptionDate], 'dd/MM/yyyy', 'en-gb' ), [WorkflowStatus],[RiskStatus] FROM " & RatingSchema & ".PolicyList WHERE"
    
    
    If Range("Search_1") <> "" Then WhereStr = WhereStr & " PolicyNo = '" & Range("Search_1") & "' AND "
    If Range("Search_2") <> "" Then WhereStr = WhereStr & " PortfolioName LIKE '%" & Range("Search_2") & "%' AND "
    If Range("Search_3") <> "" Then WhereStr = WhereStr & " SectionRef LIKE '" & Range("Search_3") & "%' AND "
    If Range("Search_4") <> "" Then WhereStr = WhereStr & " YOA >= '" & Range("Search_4") & "' AND "
    If Range("Search_5") <> "" Then WhereStr = WhereStr & " YOA <=  '" & Range("Search_5") & "' AND "
    
    GlobalSearch = "([PolicyNo] LIKE '%%' OR [PortfolioName] LIKE '%%' OR [RiskName] LIKE '%%' OR [SectionRef] LIKE '%%' OR [Underwriter] LIKE '%%' OR [InceptionDate] LIKE '%%' OR [WorkflowStatus] LIKE '%%' OR [RiskStatus] LIKE '%%')"
    GlobalSearch = Replace(GlobalSearch, "%%", "%" & GlobalSearchItem & "%")
    
    If WhereStr = "" Then
        SQLStr = SQLStr & " DeletePolicyNo is null AND " & GlobalSearch & " ORDER BY " & OrderbyField & " " & DescAsc & ";"
    Else
        SQLStr = SQLStr & WhereStr & " DeletePolicyNo is null AND " & GlobalSearch & " ORDER BY " & OrderbyField & " " & DescAsc & ";"
    End If
    
    Call openDB(Live_Server_Name, Live_Database_Name)
    
    Set rs = getRecordSet(SQLStr)
    If rs.EOF = False Then
        Control_Form.ListBox1.Column = getArrayfromRst(rs)
    Else
        Control_Form.ListBox1.Clear
    End If
    Call closeDB
    
End Sub

Sub ShowForm()

Control_Form.Show

End Sub

Sub Open_Database()
    Dim Command
    Command = Shell("ssms.exe -S " & Live_Server_Name & " -d " & Live_Database_Name & " -E ")
End Sub



