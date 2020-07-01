Attribute VB_Name = "GlobalsConst"
'Database information
Public Const Live_Server_Name As String = "DESKTOP-8GLJM1J\SQLEXPRESS"
Public Const Live_Database_Name As String = "Cargo"

Public cn As ADODB.Connection
Public rs As ADODB.Recordset

'spreadsheet controls
Public Const APP_NAME As String = "CPR" 'Used for Message Box Header
Public Const MODEL_NAME As String = "" 'Used for Userform Header to Load Risks
Public Const Password As String = "dreamteam" 'Not used currently
Public Const VersionDate = "01 Nov 2019" 'Used when saving the parameters - would ideally update this date

'Not Used
'Public Const Actuaries As String = "" 'Not used currently
'Public Const Underwriters As String = "" 'Not used currently
'Public Const AdminButtons As String = "Btn_Set_to_Live,Btn_Set_to_Test,Btn_Open_Live_DB,Btn_Open_Test_DB,Btn_Copy_Live_To_Test,Btn_Create_Rating_Tables,Btn_Create_Parameter_Tables,Btn_Create_Schema"

Public Const Division As String = "" 'Used in Parameter table
Public Const UW_Team As String = "" 'Used in Parameter table
'Public Const Business_Type As String = "ALL" 'Not used currently
Public Const ExpiringSuffix As String = "Expiry" 'Used to identify the Named Range suffix for expiring risks - needs "_" in Named Range

'---RATING SHEET INFORMATION---
Public Const RatingSheetName = "Selections" 'In this model, only used to select sheet once macros have finished
Public Const RatingPrefix As String = "SelectionsC" 'In this model, ensure that PolicyNo has this prefix in its Named Range
Public Const MultipleRatingPrefix As String = "SelectionsC,Summary" 'Used to define all of the One-to-One tables in the Schema

Public Const RatingSchedules As String = "Selections,ItemImportance"
Public Const RatingSchedulesRows As String = "SelectionsR,ItemImportanceR"
Public Const TotalSchRows As String = "SelectionsT,ItemImportanceT"
Public Const TotalSchCols As String = "SelectionsC,ItemImportanceC"
Public Const SchRowOrCol As String = "SelectionsI,ItemImportanceI" 'Identifier for rows or columns

'---PARAMETER SHEET INFORMATION---
Public Const ParameterPrefix As String = "Param"
Public Const ParameterSchedules As String = "CreditMatrix,TransactionType,Country,Security,Industry"
Public Const ParameterSchedulesRows As String = "CreditMatrixR,TransactionTypeR,CountryR,SecurityR,IndustryR"

'---DATABASE INFORMATION---
Public Const RatingSchema As String = "cpr"
Public Const ParamSchema As String = "cpp"

Public Policy_No As Double
Public Param_No As Double
Public SelectRiskAction As String 'Action is either to Load or Renew

'Unused Currently
'Public Function GetDBTestFileName()
'    GetDBTestFileName = ""
'End Function
'
'Public Function GetDBLiveFileName()
'    GetDBLiveFileName = ""
'End Function
'
'Public Function GetTableType()
'    GetDBPFileName = ""
'End Function

