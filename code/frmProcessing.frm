VERSION 5.00
Begin VB.Form frmProcessing 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   4695
      Begin VB.Label Label1 
         Caption         =   "Please wait while Collections is processing information."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   1
         Top             =   180
         Width           =   3915
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   240
         Picture         =   "frmProcessing.frx":0000
         Top             =   300
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    sFrmProcessing = 1
End Sub

Private Sub Form_Load()
    Me.Left = 300
    Me.top = 200
    If sProcessingText <> "" Then
        Text1.Text = sProcessingText
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sProcessingText = ""
    sFrmProcessing = 0
End Sub

