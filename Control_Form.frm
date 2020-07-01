VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Control_Form 
   Caption         =   "APP_NAME"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18945
   OleObjectBlob   =   "Control_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Control_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ResultsArry
Public AscDesc As String

Private Sub Btn_Delete_Click()
    
    Dim SQLStr, i, DelMsgBox

    For i = 0 To ListBox1.ListCount - 1
          
          If ListBox1.Selected(i) Then
                'checks to see if the user would like to delete the risk first before deleting
              DelMsgBox = MsgBox("Do you want to delete" & vbNewLine & vbNewLine & "PolicyNo: " & Me.ListBox1.List(i, 0) & vbNewLine & vbNewLine & "Assured Name: " & Me.ListBox1.List(i, 1), vbExclamation + vbOKCancel)
              
              If DelMsgBox = vbCancel Then
                  MsgBox "Risk not deleted", vbOKOnly + vbInformation
                  Exit Sub
              Else
                   SQLStr = "UPDATE " & RatingSchema & ".PolicyList SET " & _
                  "[DeletePolicyNo] =" & "'Yes'" & _
                  " WHERE PolicyNo = " & Me.ListBox1.List(i, 0)
                  ExecuteSQL (SQLStr)
              End If
            
          End If
          
      Next i
      
      Call LoadDataToListBox("InceptionDate", AscDesc, Me.tb_Search.Text)
      
      MsgBox "Risk Successfully Deleted!", vbExclamation + vbOKOnly

End Sub

Private Sub Btn_LoadRisk_Click()
SelectRiskAction = "Load"

If ListBox1.ListCount = 1 Then
    loadRisk (Me.ListBox1.List(i, 0))
Else
    For i = 0 To ListBox1.ListCount - 1
              If ListBox1.Selected(i) Then
                 loadRisk (Me.ListBox1.List(i, 0))
                 'loadRisk (639)
              End If
    Next i
End If

Unload Me
MsgBox "Risk Successfully Loaded", vbOKOnly + vbInformation
EndMacro
End Sub

Private Sub Btn_RenewRisk_Click()

SelectRiskAction = "Renew"
For i = 0 To ListBox1.ListCount - 1
          If ListBox1.Selected(i) Then
             loadRisk (Me.ListBox1.List(i, 0))
          End If
      Next i

Unload Me
MsgBox "Renewed Risk Successfully Loaded", vbOKOnly + vbInformation
EndMacro
End Sub


Private Sub CommandButton1_Click()
SelectRiskAction = "Load"
For i = 0 To ListBox1.ListCount - 1
          If ListBox1.Selected(i) Then
             loadRisk (Me.ListBox1.List(i, 0))
          End If
      Next i

Unload Me
MsgBox "Risk Successfully Loaded", vbOKOnly + vbInformation
EndMacro
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SelectRiskAction = "Load"
For i = 0 To ListBox1.ListCount - 1
          If ListBox1.Selected(i) Then
             loadRisk (Me.ListBox1.List(i, 0))
          End If
      Next i

Unload Me
MsgBox "Risk Successfully Loaded", vbOKOnly + vbInformation
EndMacro
End Sub


Sub SwitchAscDesc(AscDesc)
    If AscDesc = "Asc" Then
        AscDesc = "Desc"
    Else
       AscDesc = "Asc"
    End If
End Sub

Private Sub Btn_Policy_Reference_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("SectionRef", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub Btn_RiskName_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("RiskName", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub Btn_Risk_Status_Click()
  Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("RiskStatus", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub Btn_Underwriter_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("Underwriter", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub Btn_WorkflowStatus_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("WorkflowStatus", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub Btn_Inception_Date_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("InceptionDate", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub Btn_Insured_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("PortfolioName", AscDesc, Me.tb_Search.Text)
End Sub

Private Sub ListBox1_Click()

End Sub



Private Sub tb_Search_Change()

Call LoadDataToListBox("PolicyNo", AscDesc, Me.tb_Search.Text)

End Sub

Private Sub UserForm_Initialize()
    AscDesc = "Asc"
    Call LoadDataToListBox("InceptionDate", "Desc", Me.tb_Search.Text)
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
  Me.Caption = MODEL_NAME
    
   '' ListBox1.ColumnWidths = Btn_Policy_Number.Width & ";" & 50 & ";" & 300 & ";" & Btn_Policy_Reference.Width & ";" & Btn_Underwriter.Width & ";" & Btn_Inception_Date.Width & ";" & Btn_WorkflowStatus.Width & ";" & Btn_WorkflowStatus.Width & ";" & Btn_Risk_Status.Width & ";80;80;80;80;80"
    
    ListBox1.ColumnWidths = Btn_Policy_Number.Width & ";" & Btn_Insured.Width & ";" & Btn_RiskName.Width & ";" & Btn_Policy_Reference.Width & ";" & Btn_Underwriter.Width & ";" & Btn_Inception_Date.Width & ";" & Btn_WorkflowStatus.Width & ";" & Btn_WorkflowStatus.Width & ";" & Btn_Risk_Status.Width & ";80;80;80;80;80"
End Sub
Private Sub Btn_Policy_Number_Click()
    Call SwitchAscDesc(AscDesc)
    Call LoadDataToListBox("PolicyNo", AscDesc, Me.tb_Search.Text)
End Sub


Private Sub UsersListBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MsgBox "Enter key pressed"
    End If
End Sub

Private Sub TextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MsgBox "hello"
    End If
End Sub


Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub
