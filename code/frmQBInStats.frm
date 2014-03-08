VERSION 5.00
Begin VB.Form frmQBInStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QB Statistics"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmQBInStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   5565
   Begin VB.Frame Frame2 
      Caption         =   "Record information"
      Height          =   2595
      Left            =   1860
      TabIndex        =   5
      Top             =   0
      Width           =   3615
      Begin VB.Label Label7 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1500
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   195
         Left            =   1500
         TabIndex        =   11
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label5 
         Height          =   195
         Left            =   1500
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   195
         Left            =   1500
         TabIndex        =   9
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Total:"
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Not Active:"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Active:"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   660
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Categories"
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1755
      Begin VB.OptionButton Option5 
         Caption         =   "Line Items"
         Height          =   195
         Left            =   420
         TabIndex        =   13
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Employees"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1500
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Payments"
         Height          =   195
         Left            =   420
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Invoices"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Customer"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmQBInStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    sFrmQBInStats = 1
    prcGrabCustomerStat
End Sub

Sub prcClearStats()
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmQBInStats = 0
End Sub

Sub prcGrabCustomerStat()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsStats      As New ADODB.Recordset
    
    prcClearStats
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " select count(*) as active from qbx_cust where cust_isactive = 'True' "
    Set rsStats = cmdCommand.Execute
    Label4.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    
    cmdCommand.CommandText = " select count(*) as active from qbx_cust where cust_isactive = 'False' "
    Set rsStats = cmdCommand.Execute
    Label5.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    
    cmdCommand.CommandText = " select count(*) as active from qbx_cust "
    Set rsStats = cmdCommand.Execute
    Label6.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Customer Stats - Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcGrabInvoiceStat()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsStats      As New ADODB.Recordset
    
    prcClearStats
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " select count(*) as active from qbx_inv where inv_enabled = '1' "
    Set rsStats = cmdCommand.Execute
    Label4.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    
    cmdCommand.CommandText = " select count(*) as active from qbx_inv where inv_enabled = '0' "
    Set rsStats = cmdCommand.Execute
    Label5.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    
    cmdCommand.CommandText = " select count(*) as active from qbx_inv "
    Set rsStats = cmdCommand.Execute
    Label6.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Customer Stats - Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcGrabPaymentStat()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsStats      As New ADODB.Recordset
    
    prcClearStats
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " select count(*) as active from qbx_inv_payments "
    Set rsStats = cmdCommand.Execute
    Label4.Caption = Trim(rsStats!Active) & ""
    Label5.Caption = "0"
    Label6.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Customer Stats - Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcGrabLineItemsStat()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsStats      As New ADODB.Recordset
    
    prcClearStats
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " select count(*) as active from qbx_inv_lineitems "
    Set rsStats = cmdCommand.Execute
    Label4.Caption = Trim(rsStats!Active) & ""
    Label5.Caption = "0"
    Label6.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Customer Stats - Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcGrabRepsStat()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsStats      As New ADODB.Recordset
    
    prcClearStats
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " select count(*) as active from qbx_reps where rep_isactive = 'True' "
    Set rsStats = cmdCommand.Execute
    Label4.Caption = Trim(rsStats!Active) & ""
    Label5.Caption = "0"
    Label6.Caption = Trim(rsStats!Active) & ""
    
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Customer Stats - Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsStats = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Private Sub Option1_Click()
    prcGrabCustomerStat
End Sub

Private Sub Option2_Click()
    prcGrabInvoiceStat
End Sub

Private Sub Option3_Click()
    prcGrabPaymentStat
End Sub

Private Sub Option4_Click()
    prcGrabRepsStat
End Sub

Private Sub Option5_Click()
    prcGrabLineItemsStat
End Sub
