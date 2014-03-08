VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Manager"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   Icon            =   "frmProfiles.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   11130
   Begin VB.CommandButton Command5 
      Caption         =   "-->"
      Height          =   315
      Left            =   8280
      TabIndex        =   27
      Top             =   4080
      Width           =   435
   End
   Begin VB.Frame Frame2 
      Height          =   4515
      Left            =   8880
      TabIndex        =   25
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Emergency Reset All Dates"
         Height          =   495
         Left            =   180
         TabIndex        =   26
         ToolTipText     =   "This will reset all of the years, months, days values for everyone to 4 years back."
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edit Attribute"
      Height          =   2295
      Left            =   5700
      TabIndex        =   14
      Top             =   1620
      Width           =   3075
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1620
         TabIndex        =   15
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "1 = yes / 0 = no"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Height          =   1395
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1635
      Left            =   3120
      TabIndex        =   8
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmProfiles.frx":0442
         Left            =   780
         List            =   "frmProfiles.frx":044F
         TabIndex        =   21
         Top             =   780
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   780
         TabIndex        =   20
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         Height          =   315
         Left            =   4860
         TabIndex        =   19
         Top             =   1140
         Width           =   675
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmProfiles.frx":0468
         Left            =   3300
         List            =   "frmProfiles.frx":046A
         TabIndex        =   13
         Top             =   1140
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmProfiles.frx":046C
         Left            =   3300
         List            =   "frmProfiles.frx":0479
         TabIndex        =   9
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "# type:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "When viewing Invoices, how far back should this user see?"
         Height          =   435
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "When viewing Payments, how far back should this user see?"
         Height          =   435
         Left            =   2640
         TabIndex        =   12
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "# type:"
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Type:"
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Profile"
      Height          =   1635
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   315
         Left            =   1860
         TabIndex        =   18
         Top             =   1200
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable This Profile"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         Top             =   660
         Width           =   1995
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6780
      TabIndex        =   1
      Top             =   4080
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   2835
      Left            =   60
      TabIndex        =   0
      Top             =   1680
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   5001
      _Version        =   393216
      Cols            =   3
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sUserAttrAry(iProfileCount) As String
Dim sManageThisProfile(1, iProfileCount) As String
Dim sUpdateIndex As String
Dim sSaveAttrSetting As String

Dim KeepThisValuePayments As String
Dim KeepThisValueInvoices As String

Sub prcGetSpecificedProfile()
    Dim i As Integer
    Dim iRowNumber As Integer
    
    prcClearForm
    frmMain.prcUpdateFormSec 0
    
    If Trim(Combo1.Text) = sUser Then
        frmMain.funGrabProfile
        Text1.Text = sFullName
        iRowNumber = 1
        prcInitGrid
        For i = 1 To iProfileCount
            Check1.Value = 1
            If sProfileAttrDtlsAry(4, i) = 1 Then
                    
                If (sSecLvl > 2 And i <> 5) Or (sSecLvl < 3) Then
                    If i = 12 Then
                        Combo3.Text = sProfileAttrDtlsAry(1, i) & ""
                    ElseIf i = 13 Then
                        Combo2.Text = sProfileAttrDtlsAry(1, i) & ""
                    ElseIf i = 15 Then
                        Combo5.Text = sProfileAttrDtlsAry(1, i) & ""
                    ElseIf i = 16 Then
                        Combo4.Text = sProfileAttrDtlsAry(1, i) & ""
                    Else
                        FG1.Rows = iRowNumber + 1
                        FG1.Row = iRowNumber
                        FG1.col = 0
                        FG1.Text = sProfileAttrDtlsAry(3, i)
                        FG1.col = 1
                        FG1.Text = sProfileAttrDtlsAry(1, i)
                        iRowNumber = iRowNumber + 1
                    End If
                    FG1.col = 2
                    FG1.Text = i
                End If
                
            End If
        Next
            
    Else
        funGrabProfileForManager
    End If
    KeepThisValueInvoices = Trim(Combo2.Text)
    KeepThisValuePayments = Trim(Combo4.Text)
End Sub


Function funGrabProfileForManager() As Integer
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsProfile4M     As New ADODB.Recordset
    Dim i As Integer
    Dim iRowNumber As Integer
    
    prcInitGrid
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_profile_sp"
    
    'reg_list_user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, Trim(Combo1.Text) & "")
    cmdCommand.Parameters.Append parParameter
        
    Set rsProfile4M = cmdCommand.Execute
    
    
    FG1.Rows = 1
    If Not rsProfile4M.EOF Then
                
        rsProfile4M.MoveFirst
        
On Error Resume Next

            sManageThisProfile(1, 1) = Trim(rsProfile4M![1])
            sManageThisProfile(1, 2) = Trim(rsProfile4M![2])
            sManageThisProfile(1, 3) = Trim(rsProfile4M![3])
            sManageThisProfile(1, 4) = Trim(rsProfile4M![4])
            sManageThisProfile(1, 5) = Trim(rsProfile4M![5])
            sManageThisProfile(1, 6) = Trim(rsProfile4M![6])
            sManageThisProfile(1, 7) = Trim(rsProfile4M![7])
            sManageThisProfile(1, 8) = Trim(rsProfile4M![8])
            sManageThisProfile(1, 9) = Trim(rsProfile4M![9])
            sManageThisProfile(1, 10) = Trim(rsProfile4M![10])
            sManageThisProfile(1, 11) = Trim(rsProfile4M![11])
            sManageThisProfile(1, 12) = Trim(rsProfile4M![12])
            sManageThisProfile(1, 13) = Trim(rsProfile4M![13])
            sManageThisProfile(1, 14) = Trim(rsProfile4M![14])
            sManageThisProfile(1, 15) = Trim(rsProfile4M![15])
            sManageThisProfile(1, 16) = Trim(rsProfile4M![16])
            sManageThisProfile(1, 17) = Trim(rsProfile4M![17])
            sManageThisProfile(1, 18) = Trim(rsProfile4M![18])
            sManageThisProfile(1, 19) = Trim(rsProfile4M![19])
            sManageThisProfile(1, 20) = Trim(rsProfile4M![20])
            sManageThisProfile(1, 21) = Trim(rsProfile4M![21])
            sManageThisProfile(1, 22) = Trim(rsProfile4M![22])
            sManageThisProfile(1, 23) = Trim(rsProfile4M![23])
            sManageThisProfile(1, 24) = Trim(rsProfile4M![24])
            sManageThisProfile(1, 25) = Trim(rsProfile4M![25])
            sManageThisProfile(1, 26) = Trim(rsProfile4M![26])
            sManageThisProfile(1, 27) = Trim(rsProfile4M![27])
            sManageThisProfile(1, 28) = Trim(rsProfile4M![28])
            sManageThisProfile(1, 29) = Trim(rsProfile4M![29])
            sManageThisProfile(1, 30) = Trim(rsProfile4M![30])
            sManageThisProfile(1, 31) = Trim(rsProfile4M![31])
            sManageThisProfile(1, 32) = Trim(rsProfile4M![32])
            sManageThisProfile(1, 33) = Trim(rsProfile4M![33])
            sManageThisProfile(1, 34) = Trim(rsProfile4M![34])
            sManageThisProfile(1, 35) = Trim(rsProfile4M![35])
            sManageThisProfile(1, 36) = Trim(rsProfile4M![36])
            sManageThisProfile(1, 37) = Trim(rsProfile4M![37])
            sManageThisProfile(1, 38) = Trim(rsProfile4M![38])
            sManageThisProfile(1, 39) = Trim(rsProfile4M![39])
            sManageThisProfile(1, 40) = Trim(rsProfile4M![40])
            sManageThisProfile(1, 41) = Trim(rsProfile4M![41])
            sManageThisProfile(1, 42) = Trim(rsProfile4M![42])
            sManageThisProfile(1, 43) = Trim(rsProfile4M![43])
            sManageThisProfile(1, 44) = Trim(rsProfile4M![44])
            sManageThisProfile(1, 45) = Trim(rsProfile4M![45])
            sManageThisProfile(1, 46) = Trim(rsProfile4M![46])
        
            Check1.Value = Trim(rsProfile4M!profiles_enabled) & ""
            Text1.Text = Trim(rsProfile4M!profiles_fullname) & ""
            
            iRowNumber = 1
            For i = 1 To iProfileCount
                If sProfileAttrDtlsAry(4, i) = 1 Then
                    If i = 12 Then
                        Combo3.Text = sManageThisProfile(1, i) & ""
                    ElseIf i = 13 Then
                        Combo2.Text = sManageThisProfile(1, i) & ""
                    ElseIf i = 15 Then
                        Combo5.Text = sManageThisProfile(1, i) & ""
                    ElseIf i = 16 Then
                        Combo4.Text = sManageThisProfile(1, i) & ""
                    Else
                        FG1.Rows = iRowNumber + 1
                        FG1.Row = iRowNumber
                        FG1.col = 0
                        FG1.Text = sProfileAttrDtlsAry(3, i) & ""
                        FG1.col = 1
                        FG1.Text = sManageThisProfile(1, i) & ""
                        iRowNumber = iRowNumber + 1
                    End If
                    FG1.col = 2
                    FG1.Text = i
                End If
            Next
    Else
        'List1.AddItem "No saved settings found, user input required."
    End If
    
On Error GoTo errHandle:
    
    Set rsProfile4M = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Profile Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
End Function

Sub prcInitCombo()
    prcFillCombo
End Sub



Private Sub Combo1_Change()
    sUpdateIndex = ""
    prcClearEditFrame
    prcGetSpecificedProfile
End Sub

Private Sub Combo1_Click()
    If LCase(Trim(Combo1.Text)) <> sUser Then
        Check1.Enabled = True
        Command1.Enabled = True
    Else
        Check1.Enabled = False
        Command1.Enabled = False
    End If
        
    sUpdateIndex = ""
    prcClearEditFrame
    prcGetSpecificedProfile
End Sub

Sub prcClearEditFrame()
    Label9.Caption = ""
    Text2.Text = ""
    Frame4.Enabled = False
End Sub

Sub prcRefresh()
    sUpdateIndex = ""
    prcClearEditFrame
    prcGetSpecificedProfile
End Sub



Private Sub Combo2_Change()
    If Trim(Combo2.Text) <> "years" And Trim(Combo2.Text) <> "months" And Trim(Combo2.Text) <> "days" Then
        If KeepThisValueInvoices <> Trim(Combo2.Text) Then
            Combo2.Text = KeepThisValueInvoices
        End If
    End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)

    KeepThisValueInvoices = Trim(Combo2.Text)
End Sub

Private Sub Combo4_Change()
    If Trim(Combo4.Text) <> "years" And Trim(Combo4.Text) <> "months" And Trim(Combo4.Text) <> "days" Then
        If KeepThisValuePayments <> Trim(Combo4.Text) Then
            Combo4.Text = KeepThisValuePayments
        End If
    End If
    
End Sub

Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
    KeepThisValuePayments = Trim(Combo4.Text)
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    If Trim(Combo1.Text) <> sUser Then
        frmMain.prcUpdateProfile Trim(Combo1.Text), Check1.Value
        frmMain.funGrabProfile
        prcGetSpecificedProfile
    End If
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If Not IsNumeric(Trim(Combo2.Text)) Then
        frmMain.prcUpdateOneAttr 13, Trim(Combo2.Text), Trim(Combo1.Text)
    End If
    If IsNumeric(Trim(Combo3.Text)) Then
        frmMain.prcUpdateOneAttr 12, Trim(Combo3.Text), Trim(Combo1.Text)
    End If
    If Not IsNumeric(Trim(Combo4.Text)) Then
        frmMain.prcUpdateOneAttr 16, Trim(Combo4.Text), Trim(Combo1.Text)
    End If
    If IsNumeric(Trim(Combo5.Text)) Then
        frmMain.prcUpdateOneAttr 15, Trim(Combo5.Text), Trim(Combo1.Text)
    End If
    prcGetSpecificedProfile
End Sub

Private Sub Command4_Click()
    prcResetDateValues4All_EmergencyBtn
    prcGetSpecificedProfile
End Sub

Sub prcResetDateValues4All_EmergencyBtn()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sQuery As String

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then Exit Sub
    
    sQuery = "update qb_features " & _
            " set 12='-4', " & _
            " 13='years', " & _
            " 15='-4', " & _
            " 16='years' "
            
    'Debug.Print sQuery
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
        
    cmdCommand.CommandText = sQuery
    
    cmdCommand.Execute
            
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:

        Screen.MousePointer = vbDefault
        'prcMainLog Now & "-prcLogout-Error: " & err.Description & "~" & err.number & ", agent: " & frmMain.strUserTsr
        Resume Next
        
End Sub

Private Sub Command5_Click()
    If Me.Width = 8955 Then
        Me.Width = 11220
    Else
        Me.Width = 8955
    End If
End Sub

Private Sub FG1_Click()
    sUpdateIndex = ""
    FG1.Row = FG1.RowSel
    FG1.col = 2
    sUpdateIndex = FG1.Text
    FG1.col = 0
    Label9.Caption = "Rule (" & sUpdateIndex & "): " & vbNewLine & FG1.Text
    FG1.col = 1
    Text2.Text = FG1.Text
    Frame4.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Width = 8955
    Me.Height = 5070
    sUpdateIndex = ""
    prcClearEditFrame
    prcInitCombo
    frmMain.prcGrabProfileDtls
    Combo1.Text = sUser
    Check1.Value = 1
    Check1.Enabled = False
    Command1.Enabled = False
    prcFillComboDefault
    KeepThisValueInvoices = Trim(Combo2.Text)
    KeepThisValuePayments = Trim(Combo4.Text)
End Sub

Sub prcFillComboDefault()
    Dim i As Integer
    
    For i = -60 To 0
        Combo3.AddItem i
        Combo5.AddItem i
    Next i
End Sub

Sub prcClearForm()
    Check1.Value = 0
    Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    Combo5.Text = ""
End Sub

Sub prcInitGrid()
    FG1.Clear
    FG1.ColWidth(0) = 4400
    FG1.ColWidth(1) = 700
    FG1.ColWidth(2) = 0
    FG1.Row = 0
    FG1.col = 0
    FG1.Text = "Features"
    FG1.Row = 0
    FG1.col = 1
    FG1.Text = "1 / 0"
End Sub

Sub prcFillCombo()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsProfileDtls     As New ADODB.Recordset
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_profile_names_sp"
        
    Set rsProfileDtls = cmdCommand.Execute
    
    
    If Not rsProfileDtls.EOF Then
        Dim j As Integer
                
        rsProfileDtls.MoveFirst
        While Not rsProfileDtls.EOF
            If Trim(rsProfileDtls!profiles_username) = sUser Then
                Combo1.AddItem Trim(rsProfileDtls!profiles_username) & ""
            Else
                If sSecLvl = 1 Then
                    Combo1.AddItem Trim(rsProfileDtls!profiles_username) & ""
                End If
                If sSecLvl = 2 Then
                    If Trim(rsProfileDtls!profiles_level) > 2 Then
                        Combo1.AddItem Trim(rsProfileDtls!profiles_username) & ""
                    End If
                End If
                If sSecLvl = 3 Then
                    If Trim(rsProfileDtls!profiles_level) > 3 Then
                        Combo1.AddItem Trim(rsProfileDtls!profiles_username) & ""
                    End If
                End If
            End If
            rsProfileDtls.MoveNext
        Wend
    End If
    
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Profile Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub

















Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    sSaveAttrSetting = Trim(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    'If IsNumeric(Trim(Text2.Text)) Then
    '    prcUpdateOneAttr sUpdateIndex, Trim(Text2.Text)
    'Else
    '    Text2.Text = sSaveAttrSetting
    'End If
    
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(Trim(Text2.Text)) Then
        frmMain.prcUpdateOneAttr sUpdateIndex, Trim(Text2.Text), Trim(Combo1.Text)
        frmMain.funGrabProfile
        prcGetSpecificedProfile
    Else
        Text2.Text = sSaveAttrSetting
    End If
    
End Sub
