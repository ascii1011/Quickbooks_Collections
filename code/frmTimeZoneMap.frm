VERSION 5.00
Begin VB.Form frmTimeZoneMap 
   Caption         =   "Time Zone Map"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   6765
   Begin VB.PictureBox Picture1 
      Height          =   4875
      Left            =   60
      Negotiate       =   -1  'True
      Picture         =   "frmTimeZoneMap.frx":0000
      ScaleHeight     =   4815
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   60
      Width           =   6615
   End
End
Attribute VB_Name = "frmTimeZoneMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Width = 6885
    Me.Height = 5535
End Sub
