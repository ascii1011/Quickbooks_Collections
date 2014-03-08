VERSION 5.00
Begin VB.Form frmAdmin_Reset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resetting Tool"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4230
   Begin VB.Frame Frame1 
      Caption         =   "Reset records"
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3180
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   375
         Left            =   2220
         TabIndex        =   1
         Top             =   1560
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAdmin_Reset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    prcResetInvValue
End Sub


Sub prcResetInvValue()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qbx_inv set inv_enabled = '1' where inv_enabled <> '1' and inv_enabled <> '0' "
            
    Set rsSingleRemark = cmdCommand.Execute
    
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

Private Sub Command2_Click()
    Unload Me
End Sub
