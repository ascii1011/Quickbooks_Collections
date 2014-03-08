VERSION 5.00
Begin VB.Form frmSecurityLock 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Lock"
      Height          =   795
      Left            =   60
      Picture         =   "frmSecurityLock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmSecurityLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo SecErr:

    frmMain.prcDisabled
    frmMain.prcEnabledView
    frmMain.prcSetSecurityLevel
    Unload Me
    Exit Sub
    
SecErr:
    frmMain.prcDisabled
    MsgBox "An error has occured, this application will now close." & vbNewLine & "Please re-open the application."
    Pause 2
    Unload frmMain
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = frmMain.Width - 1120
    Me.top = frmMain.Height - 2350
    sFrmSecurityLock = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmSecurityLock = 0
End Sub
