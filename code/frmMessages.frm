VERSION 5.00
Begin VB.Form frmMessages 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      Begin VB.Label Label1 
         Height          =   2355
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   4275
      End
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

On Error Resume Next
    
    Frame1.top = 0
    Frame1.Left = 60
    Frame1.Width = Me.Width - 200
    Frame1.Height = Me.Height - 500
    
    Label1.top = 300
    Label1.Left = 120
    Label1.Width = Me.Width - 500
    Label1.Height = Me.Height - 1000
    
End Sub
