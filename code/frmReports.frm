VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReports2 
   Caption         =   "Form2"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   9750
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   120
      TabIndex        =   5
      Top             =   6900
      Width           =   8055
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   8340
      TabIndex        =   4
      Text            =   "71058"
      Top             =   3720
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   12091
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmReports.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmReports.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmReports.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.TextBox Text1 
         Height          =   6315
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Text            =   "frmReports.frx":0054
         Top             =   420
         Width           =   7815
      End
      Begin VB.ListBox List1 
         Height          =   6105
         Left            =   -74820
         TabIndex        =   2
         Top             =   420
         Width           =   7755
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8340
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmReports2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    List2.AddItem "->beginsession"
    QBFC_OpenConnectionBeginSession
    List2.AddItem "<-beginsession"
    
    List2.AddItem "->Fill"
    QBFC_FillReportList
    List2.AddItem "<-Fill"
                  
    List2.AddItem "->closesession"
    QBFC_EndSessionCloseConnection
    List2.AddItem "<-closesession"
End Sub

