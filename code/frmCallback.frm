VERSION 5.00
Begin VB.Form frmCallback 
   Caption         =   "Callback Info"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmCallback.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton Command3 
         Caption         =   "<-   View"
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Top             =   2760
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   6060
         TabIndex        =   17
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save   ->"
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   1215
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Text            =   "frmCallback.frx":030A
         Top             =   2760
         Width           =   4455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000C&
         Caption         =   "Completed"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   1215
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Text            =   "frmCallback.frx":030C
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   3060
         TabIndex        =   2
         Top             =   600
         Width           =   1035
      End
      Begin VB.PictureBox picCal1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4320
         Picture         =   "frmCallback.frx":030E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   600
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   3645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Company:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   660
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7440
         TabIndex        =   10
         Top             =   4080
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Count:"
         Height          =   195
         Left            =   6840
         TabIndex        =   9
         Top             =   4080
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Follow up message:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reason:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Callback Date:"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   660
         Width           =   1050
      End
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   4500
      Width           =   8895
   End
End
Attribute VB_Name = "frmCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


''''''''''''''''''''''''sql part
Public cnMC                 As New ADODB.Connection

Public rsCallBacks          As New ADODB.Recordset

Dim ListID

Sub prcVarInit()
    iCalendarRequest = 0
    Label8.Caption = sCompany
    
End Sub

Private Sub Command1_Click()
    If Trim(Text3.Text) = "" Then
        prcInsertCallback
        prcGrabCallback
    Else
        If Check1.value = 1 Then
            If Trim(Text2.Text) = "" Then
                Dim response
                response = MsgBox("Are you sure you wish to save this 'follow up' without notes?", vbExclamation + vbYesNo, "Followup notes missing!")
                If response = vbYes Then Resume Else Exit Sub
            End If
            prcUpdateCallback
            prcGrabCallback
        Else
            MsgBox "When saving a followup, you must check 'completed' box."
        End If
    End If
End Sub

Private Sub Command2_Click()
    prcClearCallbacks
    prcNew
End Sub

Private Sub Command3_Click()
    prcClearCallbacks
    prcOpenCallbackItem
End Sub

Private Sub Form_Load()
    prcVarInit
    prcClearCallbacks
    If frmInvoiceQry.ListID <> 0 And frmInvoiceQry.ListID <> "" And sCompany <> "" Then
        prcGrabCallback
        prcNew
    Else
        MsgBox "You must choose a company, before you can look at a call back from that company."
        Unload Me
    End If
End Sub

Private Sub List1_Click()
    prcClearCallbacks
    prcOpenCallbackItem
End Sub

Sub prcOpenCallbackItem()
    Dim sItemArray
    
    sItemArray = Split(List1.List(List1.ListIndex), ",")
    
    If UBound(sItemArray) > 0 Then
        If sItemArray(4) = 0 Then
            prcOpenUncompleted
        End If
        Text1.Text = sItemArray(2)
        Text2.Text = sItemArray(3)
        Text3.Text = sItemArray(0)
        txtDate.Text = sItemArray(1)
        Check1.value = sItemArray(4)
    End If
End Sub

Private Sub picCal1_Click()
    iCalendarRequest = 3
    frmCalendar.Show
End Sub





Sub prcGrabCallback()
    Dim response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:
    
    List1.Clear
    
    Set cmdCommand.ActiveConnection = frmInvoiceQry.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_callback_sp"
    
    'reg_list_user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(frmInvoiceQry.ListID) & "")
    cmdCommand.Parameters.Append parParameter
        
    Set rsCallBacks = cmdCommand.Execute
    
    Label4.Caption = rsCallBacks.RecordCount
    If Not rsCallBacks.EOF Then
        rsCallBacks.MoveFirst
        While Not rsCallBacks.EOF
            List1.AddItem rsCallBacks!callback_index & "," & rsCallBacks!callback_callbackdate & "," & rsCallBacks!callback_reason & "," & rsCallBacks!callback_followup & "," & rsCallBacks!callback_completed
            rsCallBacks.MoveNext
        Wend
    End If
    
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Callback Record Opening Error")
            If response = vbYes Then Resume Else Exit Sub
    End Select
End Sub




Sub prcInsertCallback()
    Dim response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:
    
    Set cmdCommand.ActiveConnection = frmInvoiceQry.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "insert_callback_sp"
    
    'listid
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(frmInvoiceQry.ListID) & "")
    cmdCommand.Parameters.Append parParameter
    
    'callbackdate
    Set parParameter = cmdCommand.CreateParameter(, adDate, adParamInput, , Trim(txtDate.Text) & "")
    cmdCommand.Parameters.Append parParameter
    
    'reason
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 2000, Trim(Text1.Text) & "")
    cmdCommand.Parameters.Append parParameter
    
    'user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sUser) & "")
    cmdCommand.Parameters.Append parParameter
        
    cmdCommand.Execute
    
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Inserting Callback Error")
            If response = vbYes Then Resume Else Exit Sub
    End Select
End Sub


Sub prcUpdateCallback()
    Dim response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    Set cmdCommand.ActiveConnection = frmInvoiceQry.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "update_callback_sp"
    
    'index
    Set parParameter = cmdCommand.CreateParameter(, adInteger, adParamInput, , Trim(Text3.Text) & "")
    cmdCommand.Parameters.Append parParameter
    
    'followup
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 2000, Trim(Text2.Text) & "")
    cmdCommand.Parameters.Append parParameter
    
    'completed
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Check1.value & "")
    cmdCommand.Parameters.Append parParameter
    
    'user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sUser) & "")
    cmdCommand.Parameters.Append parParameter
    
    cmdCommand.Execute
        
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "updating Importance Error")
            If response = vbYes Then Resume Else Exit Sub
    End Select
End Sub


Sub prcClearCallbacks()
    Text1.Text = ""
    Text1.Enabled = True
    Text1.BackColor = &H8000000F
    Text2.Text = ""
    Text2.Enabled = False
    Text2.BackColor = &H8000000F
    Text3.Text = ""
    Text3.Enabled = False
    txtDate.Text = ""
    txtDate.Enabled = False
    txtDate.BackColor = &H8000000F
    Check1.value = 0
    Check1.Enabled = False
    Check1.BackColor = &H8000000C
End Sub

Sub prcNew()
    Text1.Enabled = True
    Text1.BackColor = &H80000005
    txtDate.Enabled = True
    txtDate.BackColor = &H80000005
End Sub

Sub prcOpenUncompleted()
    Text2.Enabled = True
    Text2.BackColor = &H80000005
    Check1.Enabled = True
    Check1.BackColor = &H8000000F
End Sub
