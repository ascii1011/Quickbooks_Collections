VERSION 5.00
Begin VB.Form frmAlertSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert Settings"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5865
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   315
         Left            =   4440
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmAlertSettings.frx":0000
         Left            =   2280
         List            =   "frmAlertSettings.frx":0031
         TabIndex        =   5
         Text            =   "1"
         Top             =   720
         Width           =   795
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   1140
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmAlertSettings.frx":0068
         Left            =   2280
         List            =   "frmAlertSettings.frx":0078
         TabIndex        =   3
         Text            =   "1000"
         Top             =   1560
         Width           =   915
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmAlertSettings.frx":0094
         Left            =   3420
         List            =   "frmAlertSettings.frx":00B6
         TabIndex        =   2
         Text            =   "00"
         Top             =   1560
         Width           =   675
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   315
         Left            =   4440
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Active Invoices:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Importance Level:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Total Open Balance:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label5 
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   2100
         Width           =   5235
      End
      Begin VB.Label Label13 
         Caption         =   "Anything at or above these settings will be marked with an alert status."
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   4995
      End
   End
End
Attribute VB_Name = "frmAlertSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    prcUpdateAlertSettings
    prcGrabAllSettings
End Sub

Private Sub Form_Load()
    prcGrabAllSettings
End Sub

Sub prcInitSettingVars()
    prcClearForm
    prcFillImportanceDisplay
End Sub

Sub prcClearForm()
    Combo1.Text = ""
    Combo2.Clear
    Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
End Sub


Sub prcFillImportanceDisplay()
    Dim i As Integer
    
    Combo2.Clear
    
    For i = 0 To UBound(aryGImportLvl) - 1
        'fill alert management
        If aryGImportLvl(i, 0) = 1 Then
            Combo2.AddItem aryGImportLvl(i, 1) & ":" & aryGImportLvl(i, 2)
        End If
    Next i
End Sub

Sub prcGrabAllSettings()
    prcInitSettingVars
    prcGrabbingAlertSettings
End Sub

'''''''''''''''''''''''''''''Alert Settings'''''''''''''''''''''''''''''''
Sub prcGrabbingAlertSettings()
    
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsAlertSettings As New ADODB.Recordset
    Dim sMoney As String
    Dim i As Integer
    Dim lvl As Integer
    Dim sSpltMoney

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_alert_settings "
            
    Set rsAlertSettings = cmdCommand.Execute
    
    If Not rsAlertSettings.EOF Then
        rsAlertSettings.MoveFirst
        Combo1.Text = Trim(rsAlertSettings!alert_setting_max_invoices) & ""
        For i = 0 To UBound(aryGImportLvl) - 1
            lvl = Trim(rsAlertSettings!alert_setting_start_at_level)
            If aryGImportLvl(i, 1) = lvl Then
                Combo2.Text = aryGImportLvl(i, 1) & ":" & aryGImportLvl(i, 2)
            End If
        Next i
        
        sMoney = Trim(rsAlertSettings!alert_setting_max_dollar) & ""
        If sMoney <> "" Then
            sSpltMoney = Split(sMoney, ".")
            If UBound(sSpltMoney) > 0 Then
                Combo3.Text = sSpltMoney(0)
                Combo4.Text = sSpltMoney(1)
            Else
                Combo3.Text = sMoney
                Combo4.Text = "00"
            End If
        End If
    End If
    
    Set rsAlertSettings = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsAlertSettings = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcUpdateAlertSettings()
    
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sMoney As String
    Dim sSpltLvl
    Dim iLvl As Integer
    Dim iInvoices As Integer
    Dim iCheckGood As Integer

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    iCheckGood = 0
    'checking dollar values
    If IsNumeric(Trim(Combo3.Text)) = True And IsNumeric(Trim(Combo4.Text)) Then
        sMoney = Trim(Combo3.Text) & "." & Trim(Combo4.Text)
        iCheckGood = 1
    End If
    'checking level number
    sSpltLvl = Split(Trim(Combo2.Text), ":")
    If UBound(sSpltLvl) > 0 Then
        If IsNumeric(Trim(sSpltLvl(0))) = True Then
            If prcCheckNewImportanceNameAndIdBeforeUpdating(sSpltLvl) = True Then
                iLvl = Trim(sSpltLvl(0))
                iCheckGood = iCheckGood + 2
            Else
                'id and name are not correct!!!!!!!error note
            End If
        End If
    End If
    'checking max invoices
    If IsNumeric(Trim(Combo1.Text)) = True Then
        iInvoices = Trim(Combo1.Text)
        iCheckGood = iCheckGood + 4
    End If
    
    If iCheckGood <> 7 Then
        If iCheckGood = 1 Then
            Label6.Caption = "Error with Importance level and Invoices"
        ElseIf iCheckGood = 2 Then
            Label6.Caption = "Error with Total Balance and Invoices"
        ElseIf iCheckGood = 3 Then
            Label6.Caption = "Error with Invoices"
        ElseIf iCheckGood = 4 Then
            Label6.Caption = "Error with Total Balance and Importance level"
        ElseIf iCheckGood = 5 Then
            Label6.Caption = "Error with Importance level"
        ElseIf iCheckGood = 6 Then
            Label6.Caption = "Error with Total Balance"
        ElseIf iCheckGood = 0 Then
            Label6.Caption = "Error with all settings"
        Else
            Label6.Caption = "Error unknown"
        End If
        Exit Sub
    End If
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qbx_alert_settings " & _
                        " set alert_setting_max_dollar = CONVERT(MONEY, '" & sMoney & "'), " & _
                        " alert_setting_start_at_level = '" & iLvl & "', " & _
                        " alert_setting_max_invoices = '" & iInvoices & "' "
            
    cmdCommand.Execute
    
    Label6.Caption = "Alert Settings updated."
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Function prcCheckNewImportanceNameAndIdBeforeUpdating(aryImport) As Boolean
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsAlertSettings As New ADODB.Recordset

    prcCheckNewImportanceNameAndIdBeforeUpdating = False
On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_importance_levels where import_name = '" & aryImport(1) & "' and import_id = '" & aryImport(0) & "' "
            
    Set rsAlertSettings = cmdCommand.Execute
    
    If Not rsAlertSettings.EOF Then
        prcCheckNewImportanceNameAndIdBeforeUpdating = True
    End If
    
    Set rsAlertSettings = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsAlertSettings = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function
