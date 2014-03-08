VERSION 5.00
Begin VB.Form frmQueryFilter 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   1635
      Begin VB.CommandButton Command1 
         Caption         =   "<--"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "View Zero Bal"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   1395
      End
      Begin VB.CheckBox Check2 
         Caption         =   "View Open"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "View Media"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmQueryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    prcUpdate 6, Check1
    frmInvoiceQry.prcRefresh
End Sub


Sub prcUpdate(sindex As String, cCheck As CheckBox)
    frmMain.prcUpdateOneAttr sindex, cCheck.value, sUser
    frmMain.funGrabProfile
    prcRefresh
End Sub

Sub prcRefresh()
    Check1.value = sProfileAttrDtlsAry(1, 6)
    Check2.value = sProfileAttrDtlsAry(1, 10)
    Check3.value = sProfileAttrDtlsAry(1, 11)
End Sub

Private Sub Check2_Click()
    prcUpdate 10, Check2
    frmInvoiceQry.prcRefresh
End Sub

Private Sub Check3_Click()
    prcUpdate 11, Check3
    frmInvoiceQry.prcRefresh
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.top = frmInvoiceQry.top
    Me.Left = frmInvoiceQry.Left + frmInvoiceQry.Width
    Me.Width = 1665
    Me.Height = 1665
    sFrmQueryFilter = 1
    prcRefresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmQueryFilter = 0
End Sub
