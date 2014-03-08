VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrintPage 
   Caption         =   "Printing"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   Icon            =   "frmPrintPage.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11925
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   13361
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Preview a Letter"
      TabPicture(0)   =   "frmPrintPage.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "WB1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Modify a Letter"
      TabPicture(1)   =   "frmPrintPage.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Variables"
      TabPicture(2)   =   "frmPrintPage.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   7095
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   9375
         Begin TabDlg.SSTab SSTab2 
            Height          =   6855
            Left            =   60
            TabIndex        =   27
            Top             =   180
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   12091
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   423
            TabCaption(0)   =   "Main"
            TabPicture(0)   =   "frmPrintPage.frx":035E
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame6"
            Tab(0).Control(1)=   "Frame5"
            Tab(0).Control(2)=   "Frame2"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Fax"
            TabPicture(1)   =   "frmPrintPage.frx":037A
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Frame8"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame8 
               Caption         =   "General Information"
               Height          =   6435
               Left            =   120
               TabIndex        =   52
               Top             =   300
               Width           =   8895
               Begin VB.TextBox Text18 
                  Height          =   315
                  Left            =   5760
                  TabIndex        =   79
                  Top             =   1740
                  Width           =   2775
               End
               Begin VB.TextBox Text17 
                  Height          =   315
                  Left            =   5760
                  TabIndex        =   77
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.TextBox Text16 
                  Height          =   315
                  Left            =   5760
                  TabIndex        =   75
                  Top             =   1020
                  Width           =   2775
               End
               Begin VB.TextBox Text15 
                  Height          =   315
                  Left            =   5760
                  TabIndex        =   73
                  Top             =   660
                  Width           =   2775
               End
               Begin VB.TextBox Text14 
                  Height          =   315
                  Left            =   5760
                  TabIndex        =   71
                  Top             =   300
                  Width           =   2775
               End
               Begin VB.TextBox Text13 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   69
                  Top             =   1740
                  Width           =   2775
               End
               Begin VB.TextBox Text5 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   67
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.TextBox Text12 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   65
                  Top             =   1020
                  Width           =   2775
               End
               Begin VB.TextBox Text11 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   63
                  Top             =   660
                  Width           =   2775
               End
               Begin VB.TextBox Text7 
                  Height          =   2775
                  Left            =   180
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   62
                  Top             =   2940
                  Width           =   8475
               End
               Begin VB.TextBox Text4 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   60
                  Top             =   300
                  Width           =   2775
               End
               Begin VB.CheckBox Check5 
                  Caption         =   "PLEASE RECYCLE"
                  Height          =   195
                  Left            =   6960
                  TabIndex        =   59
                  Top             =   2280
                  Width           =   1695
               End
               Begin VB.CheckBox Check4 
                  Caption         =   "PLEASE REPLY"
                  Height          =   195
                  Left            =   5220
                  TabIndex        =   58
                  Top             =   2280
                  Width           =   1875
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "PLEASE COMMENT"
                  Height          =   195
                  Left            =   3240
                  TabIndex        =   57
                  Top             =   2280
                  Width           =   1875
               End
               Begin VB.CheckBox Check2 
                  Caption         =   "FOR REVIEW"
                  Height          =   195
                  Left            =   1560
                  TabIndex        =   56
                  Top             =   2280
                  Width           =   1335
               End
               Begin VB.CheckBox Check1 
                  Caption         =   "URGENT"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   55
                  Top             =   2280
                  Width           =   1035
               End
               Begin VB.CommandButton Command19 
                  Caption         =   "Reset Defaults"
                  Height          =   315
                  Left            =   180
                  TabIndex        =   54
                  Top             =   5880
                  Width           =   1275
               End
               Begin VB.CommandButton Command8 
                  Caption         =   "Update"
                  Height          =   315
                  Left            =   7800
                  TabIndex        =   53
                  Top             =   5880
                  Width           =   855
               End
               Begin VB.Label Label21 
                  Caption         =   "NOTES / COMMENTS:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   81
                  Top             =   2700
                  Width           =   2175
               End
               Begin VB.Label Label20 
                  Caption         =   "Sender's Fax:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   80
                  Top             =   1800
                  Width           =   1155
               End
               Begin VB.Label Label19 
                  Caption         =   "Sender's Phone:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   78
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.Label Label18 
                  Caption         =   "Pages:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   76
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label17 
                  Caption         =   "Date:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   74
                  Top             =   720
                  Width           =   735
               End
               Begin VB.Label Label16 
                  Caption         =   "From:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   72
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label Label12 
                  Caption         =   "RE:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   70
                  Top             =   1800
                  Width           =   1155
               End
               Begin VB.Label Label8 
                  Caption         =   "Phone Number:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   68
                  Top             =   1440
                  Width           =   1155
               End
               Begin VB.Label Label11 
                  Caption         =   "Fax Number:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   66
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label10 
                  Caption         =   "Company:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   64
                  Top             =   720
                  Width           =   735
               End
               Begin VB.Label Label3 
                  Caption         =   "To:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   61
                  Top             =   360
                  Width           =   735
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Customer Information"
               Height          =   6435
               Left            =   -74880
               TabIndex        =   37
               Top             =   300
               Width           =   3975
               Begin VB.TextBox Text20 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   83
                  Top             =   1500
                  Width           =   2715
               End
               Begin VB.TextBox Text19 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   82
                  Top             =   1080
                  Width           =   2715
               End
               Begin VB.CommandButton Command9 
                  Caption         =   "Update"
                  Height          =   315
                  Left            =   2940
                  TabIndex        =   46
                  Top             =   4380
                  Width           =   855
               End
               Begin VB.CommandButton Command10 
                  Caption         =   "Reset Defaults"
                  Height          =   315
                  Left            =   180
                  TabIndex        =   45
                  Top             =   4380
                  Width           =   1275
               End
               Begin VB.TextBox Text8 
                  Height          =   1275
                  Left            =   1080
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   44
                  Top             =   1860
                  Width           =   2715
               End
               Begin VB.TextBox Text9 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   43
                  Top             =   3240
                  Width           =   1335
               End
               Begin VB.TextBox Text10 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   42
                  Top             =   3720
                  Width           =   1335
               End
               Begin VB.TextBox Text2 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   41
                  Top             =   720
                  Width           =   2715
               End
               Begin VB.TextBox Text1 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   40
                  Top             =   300
                  Width           =   2715
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   270
                  Left            =   2460
                  Picture         =   "frmPrintPage.frx":0396
                  ScaleHeight     =   240
                  ScaleWidth      =   240
                  TabIndex        =   39
                  ToolTipText     =   "Today's callbacks"
                  Top             =   3240
                  Width           =   270
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   270
                  Left            =   2460
                  Picture         =   "frmPrintPage.frx":07D8
                  ScaleHeight     =   240
                  ScaleWidth      =   240
                  TabIndex        =   38
                  ToolTipText     =   "Today's callbacks"
                  Top             =   3720
                  Width           =   270
               End
               Begin VB.Label Label23 
                  Caption         =   "Fax:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   85
                  Top             =   1560
                  Width           =   735
               End
               Begin VB.Label Label22 
                  Caption         =   "Phone:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   84
                  Top             =   1140
                  Width           =   735
               End
               Begin VB.Label Label13 
                  Caption         =   "Address:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   51
                  Top             =   1920
                  Width           =   675
               End
               Begin VB.Label Label14 
                  Caption         =   "Datestamp:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   50
                  Top             =   3300
                  Width           =   975
               End
               Begin VB.Label Label15 
                  Caption         =   "Due Date:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   49
                  Top             =   3780
                  Width           =   795
               End
               Begin VB.Label Label2 
                  Caption         =   "Contact:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   48
                  Top             =   780
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Caption         =   "Company:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   47
                  Top             =   360
                  Width           =   735
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Modern Consumer Information"
               Height          =   4395
               Left            =   -70860
               TabIndex        =   31
               Top             =   300
               Width           =   4935
               Begin VB.CommandButton Command12 
                  Caption         =   "Update"
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   35
                  Top             =   3900
                  Width           =   855
               End
               Begin VB.CommandButton Command11 
                  Caption         =   "Reset Defaults"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   34
                  Top             =   3900
                  Width           =   1275
               End
               Begin VB.TextBox Text6 
                  Height          =   1455
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   33
                  Top             =   720
                  Width           =   4695
               End
               Begin VB.ComboBox Combo4 
                  Height          =   315
                  Left            =   900
                  TabIndex        =   32
                  Top             =   300
                  Width           =   3915
               End
               Begin VB.Label Label7 
                  Caption         =   "Variables:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   36
                  Top             =   360
                  Width           =   735
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "General Information"
               Height          =   435
               Left            =   -70860
               TabIndex        =   28
               Top             =   4740
               Visible         =   0   'False
               Width           =   4935
               Begin VB.CommandButton Command15 
                  Caption         =   "Update"
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   30
                  Top             =   300
                  Width           =   855
               End
               Begin VB.CommandButton Command14 
                  Caption         =   "Reset Defaults"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   29
                  Top             =   300
                  Width           =   1275
               End
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   7095
         Left            =   -74940
         TabIndex        =   7
         Top             =   60
         Width           =   9375
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   8220
            Picture         =   "frmPrintPage.frx":0C1A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   22
            ToolTipText     =   "Today's callbacks"
            Top             =   1680
            Width           =   270
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Reset"
            Height          =   315
            Left            =   8340
            TabIndex        =   21
            ToolTipText     =   "Resets Variables back to the original values"
            Top             =   4320
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            Caption         =   "?"
            Height          =   255
            Left            =   8220
            TabIndex        =   18
            Top             =   300
            Width           =   255
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Top             =   300
            Width           =   6195
         End
         Begin VB.CommandButton Command7 
            Caption         =   "?"
            Height          =   255
            Left            =   8220
            TabIndex        =   15
            Top             =   1260
            Width           =   255
         End
         Begin VB.CommandButton Command5 
            Caption         =   "?"
            Height          =   255
            Left            =   8220
            TabIndex        =   14
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Update"
            Height          =   315
            Left            =   8340
            TabIndex        =   11
            Top             =   5040
            Width           =   855
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Top             =   840
            Width           =   6195
         End
         Begin VB.TextBox Text3 
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   8
            Top             =   1200
            Width           =   7995
         End
         Begin VB.Label Label9 
            Caption         =   "Customer Variables:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label Label5 
            Caption         =   "Parts of this letter to edit:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   1755
         End
      End
      Begin SHDocVwCtl.WebBrowser WB1 
         Height          =   7095
         Left            =   -74940
         TabIndex        =   5
         Top             =   60
         Width           =   9375
         ExtentX         =   16536
         ExtentY         =   12515
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printing a Letter"
      Height          =   7575
      Left            =   9600
      TabIndex        =   0
      Top             =   60
      Width           =   2235
      Begin VB.CommandButton Command22 
         Caption         =   "Email"
         Height          =   315
         Left            =   660
         TabIndex        =   88
         Top             =   2280
         Width           =   915
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Print Properties"
         Height          =   375
         Left            =   420
         TabIndex        =   87
         Top             =   5280
         Width           =   1395
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Print Preview"
         Height          =   375
         Left            =   420
         TabIndex        =   86
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Frame Frame7 
         Caption         =   "How to"
         Height          =   1995
         Left            =   480
         TabIndex        =   23
         Top             =   3000
         Width           =   1275
         Begin VB.CommandButton Command18 
            Caption         =   "Edit Variables"
            Height          =   435
            Left            =   180
            TabIndex        =   26
            Top             =   1440
            Width           =   915
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Modify Letters"
            Height          =   435
            Left            =   180
            TabIndex        =   25
            Top             =   900
            Width           =   915
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Preview/ Printing"
            Height          =   435
            Left            =   180
            TabIndex        =   24
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   1995
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Preview"
         Height          =   315
         Left            =   660
         TabIndex        =   3
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
         Height          =   375
         Left            =   420
         TabIndex        =   2
         Top             =   6240
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Left            =   660
         TabIndex        =   1
         Top             =   7080
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Choose a Letter Template:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   900
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice #:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPrintPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sFilename As String

Dim sInvInfo As String
Dim sInvDtls As String
Dim iInvCount               As Integer
Dim iInvLineItemCount       As Integer
Dim iInvBoolmark            As Integer

'letter parts
Dim sCurrentLetterName As String
Dim sCurrentLetterPartsBackground() As String
Dim sCurrentLetterPartsToModify() As String
Dim sCurrentLetterPartModifying As String
Dim iLetterCount As Integer
Dim sCurrentLetterTitle As String


'letter variables
Dim sDefaultModConAddress() As String
Dim sCurrentModConAddress() As String

Dim sDefaultFaxVars() As String
Dim sCurrentFaxVars() As String

Dim sDefaultCustomerCompany As String
Dim sCurrentCustomerCompany As String

Dim sDefaultCustomerContact As String
Dim sCurrentCustomerContact As String

Dim sDefaultCustomerPhone As String
Dim sCurrentCustomerPhone As String

Dim sDefaultCustomerBillingAddress As String
Dim sCurrentCustomerBillingAddress As String
Dim sCurrentCustomerBillingAddressWeb As String

Dim sDefaultCustomerShippingAddress As String
Dim sCurrentCustomerShippingAddress As String
Dim sCurrentCustomerShippingAddressWeb As String

Dim sDefaultCustomerTimestamp As String
Dim sCurrentCustomerTimestamp As String

Dim sDefaultCustomerDueDate As String
Dim sCurrentCustomerDueDate As String

Dim sCurrentCustomerNumberOfInvoices As String
Dim sCurrentCustomerListOfInvoices As String

Dim sCurrentCustomerTotalInvoiceBalance As String

Dim sCountOfInvoicesOverDue As String
Dim sListOfInvoicesOverDue As String
Dim sListOfInvoicesOverDueWithDetails As String
Dim sBalanceOfInvoicesOverDue As String

'invoice General Info
Dim sCurrentCustomerTxndate As String
Dim sCurrentCustomerInvoiceNum As String
Dim sCurrentCustomerFullName As String
Dim sCurrentCustomerInvoiceDueDate As String
Dim sCurrentCustomerSalesRep As String
Dim sCurrentCustomerArAccountRef_Name As String
Dim sCurrentCustomerAppliedAmount As String
Dim sCurrentCustomerBalanceRemaining As String
Dim sCurrentCustomerMsgRef_Name As String
Dim sCurrentCustomerSubtotal As String

'invoice LineItems
Dim sCurrentCustomerInvoiceDetailSummary As String

Sub prcResetDefaultLetterParts()
    sCurrentModConAddress(0, 1) = sDefaultModConAddress(0, 1)
    prcDisplayModConVars
End Sub




Private Sub Combo1_Change()
    prcInitCustomerInvoiceInfo
End Sub

Private Sub Combo1_Click()
    prcInitCustomerInvoiceInfo
End Sub

Private Sub Combo2_Click()
    sCurrentLetterName = Trim(Combo2.Text)
    prcPullLetterPartsFromName
    prcGrabAllModifyInfo
End Sub

Sub prcUpdateCurrentLetterPartToModify()
    Dim i As Integer
    
    'if combo3 entry is a proper name, then store as the current letter part to modify
    For i = 0 To UBound(sCurrentLetterPartsToModify)
        If Trim(Combo3.Text) = Trim(sCurrentLetterPartsToModify(i, 0)) Then
            sCurrentLetterPartModifying = Trim(sCurrentLetterPartsToModify(i, 0))
            prcDisplayLetterBreakDownPart
            Exit Sub
        End If
    Next i
    sCurrentLetterPartModifying = Trim(sCurrentLetterPartsToModify(0, 0))
    prcDisplayLetterBreakDownPart
End Sub

Private Sub Combo3_Change()
    prcUpdateCurrentLetterPartToModify
End Sub

Private Sub Combo3_Click()
    prcUpdateCurrentLetterPartToModify
End Sub

Private Sub Combo4_Click()
    prcDisplayModConVars
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command10_Click()
    prcDefaultCustomerVarsReset
    prcDisplayCustomerVars
End Sub

Private Sub Command11_Click()
    prcDefaultModConVarsReset
    prcDisplayModConVars
End Sub

Private Sub Command12_Click()
    prcUpdateModConVars
End Sub

Sub prcUpdateModConVars()
    Dim i As Integer
    
    For i = 0 To UBound(sDefaultModConAddress) - 1
        If Trim(Combo4.Text) = Trim(sCurrentModConAddress(i, 0)) Then
            sCurrentModConAddress(i, 1) = Trim(Text6.Text)
        End If
    Next i
End Sub

Private Sub Command13_Click()
    MsgBox "These are all the variables available for use."
End Sub

Private Sub Command16_Click()

    frmMessages.Width = 4500
    frmMessages.Height = 2000
    frmMessages.Caption = "How to use"
    frmMessages.Frame1.Caption = "Preview/Print a Letter:"
    frmMessages.Label1.Caption = "Step 1: choose an 'invoice #'" & vbNewLine & _
        "Step 2: choose a Letter Template" & vbNewLine & _
        "Step 3: Press the 'Preview' button." & vbNewLine & _
        "Step 4: If you wish to Print, simply press the 'Print' button."
        
End Sub
Private Sub Command17_Click()

    frmMessages.Width = 9000
    frmMessages.Height = 3500
    frmMessages.Caption = "How to use"
    frmMessages.Frame1.Caption = "Modifying a Letter:"
    frmMessages.Label1.Caption = "Step 1: Choose a Letter Template Other than 'Invoice'." & vbNewLine & _
        "Step 2: Press the 'Modify a Letter' tab at the bottom." & vbNewLine & _
        "Step 3: Choose a part of the letter to edit from the second dropdown list." & vbNewLine & _
        "Step 4: The text box below it will display the contents of that Letter Part." & vbNewLine & _
        "Step 5: Edit the text and press the 'Update' Button." & vbNewLine & _
        "Step 6: Last, press the 'Preview' button to see the changes." & vbNewLine & _
        vbNewLine & vbNewLine & _
        "Note: If you see something like '$CustomerCompanyAddress' in a part of the letter, it will be replaced by the" & vbNewLine & _
        "corresponding company's address.  You may add more from the dropdown list by copying them and pasting them." & vbNewLine & _
        "You may also remove them and substitute your own text." & vbNewLine & _
        "Note: '$CustomerCompanyAddress' must be spelled exactly in this manner in order for it to work."

End Sub
Private Sub Command18_Click()

    frmMessages.Width = 8500
    frmMessages.Height = 2000
    frmMessages.Caption = "How to use"
    frmMessages.Frame1.Caption = "Editing a Variable:"
    frmMessages.Label1.Caption = "Step 1: Choose a Letter Template Other than 'Invoice'." & vbNewLine & _
        "Step 2: Choose the 'Variables' tab at the bottom." & vbNewLine & _
        "Step 3: Edit Any of the fields and press the 'Update' button associated with that section." & vbNewLine & _
        "Step 4: Last, press the 'Preview' button to see the changes." & vbNewLine

End Sub

Private Sub Command19_Click()
    prcDefaultFaxVarsReset
    prcDisplayFaxVars
End Sub

Private Sub Command2_Click()
    'WB1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
    WB1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
    'WB1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command20_Click()
    WB1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command21_Click()
    WB1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command22_Click()
    frmMain.prcShowFrmEmailCustomer
End Sub

Private Sub Command3_Click()
    Dim iFileExistance
    
    prcPreviewLetter Trim(Combo2.Text)
    SSTab1.Tab = 0
        
    iFileExistance = DoesFileExist("c:\doc.html")
    
    If iFileExistance = True Then
        Command2.Enabled = True
        Command22.Enabled = True
    Else
        Command2.Enabled = False
        Command22.Enabled = False
    End If
End Sub

Sub prcGrabInvoiceIndexes()
    Dim i As Integer
    'Dim dtsCount
    
    For i = 0 To UBound(InvoicesInfo) - 1
        If InvoicesInfo(i, 1) = Trim(Combo1.Text) Then
            sInvInfo = i
            i = UBound(InvoicesInfo)
        End If
    Next
    
    
    'dtsCount = InvoicesDetails(sInvInfo, 0, 6)
    'For i = 0 To dtsCount
        'If InvoicesDetails(sInvInfo, 0, 0) = Trim(Combo1.Text) Then
        '    sInvDtls = InvoicesDetails(sInvInfo, 0, 6)
        'End If
    'Next
        
    
End Sub

Private Sub Command4_Click()
    prcUpdateLetterPartCurrentlyModified
End Sub

Sub prcUpdateLetterPartCurrentlyModified()
    Dim i As Integer
    
    For i = 0 To UBound(sCurrentLetterPartsToModify) - 1
        If Trim(Combo3.Text) = Trim(sCurrentLetterPartsToModify(i, 0)) Then
            sCurrentLetterPartsToModify(i, 1) = Trim(Text3.Text)
        End If
    Next i
End Sub

Private Sub Command5_Click()
    MsgBox "This drop down menu shows the different parts (names) of the letter that can be edited."
End Sub

Private Sub Command6_Click()
    sCurrentLetterName = Trim(Combo2.Text)
    prcPullLetterPartsFromName
    prcGrabAllModifyInfo
End Sub

Private Sub Command7_Click()
    MsgBox "This shows the text that can be edited for a given (letter) part name."
End Sub



Private Sub Command8_Click()
    prcUpdateFaxVars
    prcDisplayFaxVars
End Sub

Private Sub Command9_Click()
    prcUpdateCustomerVars
End Sub

Private Sub Form_Load()

    prcGrabInvoice
    prcInitVars
    
    WB1.navigate sGHtml_printing
        
    prcGrabLetterNames 'grabs all the names of the letters
    
    prcPullLetterPartsFromName 'grabs all the parts of a letter's name
    prcGrabAllModifyInfo
    
End Sub

Sub prcGrabAllModifyInfo()
    prcDisplayletterBreakDownList 'display the editible parts in the drop down on the modify tab
    prcDisplayLetterVariables '
    
    prcUpdateCurrentLetterPartToModify
End Sub

Sub prcInitVars()
    Me.top = 50
    Me.Left = 50
    Me.Width = 12015
    Me.Height = 8235
    
    Command2.Enabled = False
    SSTab1.Tab = 0
    iLetterCount = 0
    
    'regrab user info
    frmMain.prcGetUserInfo
    
    prcInitModConVars
    prcDefaultModConVarsReset
    prcDisplayModConVars
    
    prcInitCustomerVars
    prcDefaultCustomerVarsReset
    prcDisplayCustomerVars
    
    prcInitFaxVars
    prcDefaultFaxVarsReset
    prcDisplayFaxVars
End Sub


Sub prcDisplayFaxVars()
    
    Text4.Text = sCurrentFaxVars(0)
    Text11.Text = sCurrentFaxVars(1)
    Text12.Text = sCurrentFaxVars(2)
    Text5.Text = sCurrentFaxVars(3)
    Text13.Text = sCurrentFaxVars(4)
    Text14.Text = sCurrentFaxVars(5)
    Text15.Text = sCurrentFaxVars(6)
    Text16.Text = sCurrentFaxVars(7)
    Text17.Text = sCurrentFaxVars(8)
    Text18.Text = sCurrentFaxVars(9)
    
    Check1.Value = sCurrentFaxVars(10)
    Check2.Value = sCurrentFaxVars(11)
    Check3.Value = sCurrentFaxVars(12)
    Check4.Value = sCurrentFaxVars(13)
    Check5.Value = sCurrentFaxVars(14)
    
    Text7.Text = sCurrentFaxVars(15)

End Sub

Sub prcUpdateFaxVars()
    
    sCurrentFaxVars(0) = Text4.Text
    sCurrentFaxVars(1) = Text11.Text
    sCurrentFaxVars(2) = Text12.Text
    sCurrentFaxVars(3) = Text5.Text
    sCurrentFaxVars(4) = Text13.Text
    sCurrentFaxVars(5) = Text14.Text
    sCurrentFaxVars(6) = Text15.Text
    sCurrentFaxVars(7) = Text16.Text
    sCurrentFaxVars(8) = Text17.Text
    sCurrentFaxVars(9) = Text18.Text
    
    sCurrentFaxVars(10) = Check1.Value
    sCurrentFaxVars(11) = Check2.Value
    sCurrentFaxVars(12) = Check3.Value
    sCurrentFaxVars(13) = Check4.Value
    sCurrentFaxVars(14) = Check5.Value
    
    sCurrentFaxVars(15) = Text7.Text

End Sub


Sub prcInitFaxVars()

    ReDim sDefaultFaxVars(16)
    ReDim sCurrentFaxVars(16)
    
    sDefaultFaxVars(0) = CustomerMainInfo(2) 'contact = To:
    sDefaultFaxVars(1) = CustomerMainInfo(1) 'company
    sDefaultFaxVars(2) = CustomerMainInfo(4) 'fax
    sDefaultFaxVars(3) = CustomerMainInfo(3) 'phone
    sDefaultFaxVars(4) = "" 'RE:
    
    sDefaultFaxVars(5) = aryRep_Profile(1) 'From:
    sDefaultFaxVars(6) = Now 'today
    sDefaultFaxVars(7) = "" 'pages
    sDefaultFaxVars(8) = aryRep_Profile(4) 'sender's phone
    sDefaultFaxVars(9) = aryRep_Profile(5) 'sender's fax
    
    sDefaultFaxVars(10) = 0 'phone
    sDefaultFaxVars(11) = 0 'phone
    sDefaultFaxVars(12) = 0 'phone
    sDefaultFaxVars(13) = 0 'phone
    sDefaultFaxVars(14) = 0 'phone
    
    sDefaultFaxVars(15) = "" 'notes/comments
End Sub


Sub prcInitModConVars()
    
    ReDim sDefaultModConAddress(4, 2)
    ReDim sCurrentModConAddress(4, 2)
    sDefaultModConAddress(0, 0) = "$ModConAddress1"
    sDefaultModConAddress(0, 1) = "270 Lafayette street · suite 201 · New york · ny 10012"
    sDefaultModConAddress(1, 0) = "$ModConAddress2"
    sDefaultModConAddress(1, 1) = "Modern Consumer" & vbNewLine & "270 Lafayette street, suite 201" & vbNewLine & "New york, NY 10012" & vbNewLine & "(212) 387-9616" & vbNewLine & "(212) 214-0561 Fax"
    sDefaultModConAddress(2, 0) = "$ModConImage1"
    sDefaultModConAddress(2, 1) = "<img width=400 height=90 src=" & sGImage_Modcon_C & ">"
    sDefaultModConAddress(3, 0) = "$ModConFooter1"
    sDefaultModConAddress(3, 1) = "<div style=""border-top:double windowtext 2.25pt;border-left:none;border-right:none;padding:2.0pt 0in 1.0pt 0in"">" & _
                                "<center><p class=DocumentLabel style=""FONT: bold;FONT-SIZE:10px;border:none;padding:0in"">212-387-9616 · Toll Free 866-387-9616 · Fax 212-214-0561" & _
                                "</p></center></div>"
    
End Sub


Sub prcInitCustomerVars()
    Dim sMerge As String
    Dim i As Integer
    Dim tmpDate As Date
        
    tmpDate = Format(Now, "mm/dd/yyyy")
    
    'company
    If CustomerMainInfo(1) <> "" Then
        sDefaultCustomerCompany = CustomerMainInfo(1)
    Else
        Frame1.Enabled = False
        Exit Sub
    End If
    
    'contact
    If CustomerMainInfo(2) = "" Then
        CustomerMainInfo(2) = "To whom it may Concern"
    End If
    sDefaultCustomerContact = CustomerMainInfo(2)
        
    'datestamp
    CustomerMainInfo(7) = tmpDate
    sDefaultCustomerDueDate = CustomerMainInfo(7)
    
    'datedue
    CustomerMainInfo(8) = DateAdd("d", tmpDate, 7)
    sDefaultCustomerTimestamp = CustomerMainInfo(8)
    
    'billingaddress
    If sBilling = "" Then
        sBilling = "No address given"
    End If
    sDefaultCustomerBillingAddress = sBilling
    
    'shippingaddress
    If sShipping = "" Then
        sShipping = "No address given"
    End If
    sDefaultCustomerShippingAddress = sShipping
    
    sCurrentCustomerTotalInvoiceBalance = CustomerMainInfo(6)
    
    prcInitCustomerInvoiceInfo
    
End Sub

Sub prcInitCustomerInvoiceInfo()
    Dim i As Integer
    Dim bFound As Boolean
    
On Error GoTo Err:
    
    bFound = False
    
    sCurrentCustomerTxndate = "" 'txndate
    sCurrentCustomerInvoiceNum = "" 'invoicenumtemp
    sCurrentCustomerFullName = "" 'fullname
    sCurrentCustomerInvoiceDueDate = "" 'duedate
    sCurrentCustomerSalesRep = "" 'salesrep_name
    sCurrentCustomerArAccountRef_Name = "" 'araccountref_name
    sCurrentCustomerAppliedAmount = "" 'appliedamount
    sCurrentCustomerBalanceRemaining = "" 'balanceremaining
    sCurrentCustomerMsgRef_Name = "" 'customermsgref_name
    sCurrentCustomerSubtotal = "" 'subtotal
    
    sCurrentCustomerInvoiceDetailSummary = "" 'All Line items, formatted for the webpage to be displayed.
    
    sCurrentCustomerNumberOfInvoices = iInvBoolmark + 1
    sCurrentCustomerListOfInvoices = sCurrentCustomerListOfInvoices
    
    For i = 0 To iInvBoolmark
        If Trim(Combo1.Text) = InvoicesInfo(i, 1) Then
            sInvInfo = i
            bFound = True
            
            sCurrentCustomerTxndate = InvoicesInfo(i, 0) 'txndate
            sCurrentCustomerInvoiceNum = InvoicesInfo(i, 1) 'invoicenumtemp
            sCurrentCustomerFullName = InvoicesInfo(i, 2) 'fullname
            sCurrentCustomerInvoiceDueDate = InvoicesInfo(i, 3) 'duedate
            sCurrentCustomerSalesRep = InvoicesInfo(i, 4) 'salesrep_name
            sCurrentCustomerArAccountRef_Name = InvoicesInfo(i, 5) 'araccountref_name
            sCurrentCustomerAppliedAmount = InvoicesInfo(i, 6) 'appliedamount
            sCurrentCustomerBalanceRemaining = InvoicesInfo(i, 7) 'balanceremaining
            sCurrentCustomerMsgRef_Name = InvoicesInfo(i, 8) 'customermsgref_name
            sCurrentCustomerSubtotal = InvoicesInfo(i, 9) 'subtotal
            
            Dim k As Integer
            
            If IsNumeric(InvoicesInfo(i, 10)) Then
                For k = 0 To InvoicesInfo(i, 10) - 1
                    sCurrentCustomerInvoiceDetailSummary = sCurrentCustomerInvoiceDetailSummary & InvoicesDetails(i, k, 0) 'invoiceNumTemp
                    sCurrentCustomerInvoiceDetailSummary = sCurrentCustomerInvoiceDetailSummary & InvoicesDetails(i, k, 1) 'line_quantity
                    sCurrentCustomerInvoiceDetailSummary = sCurrentCustomerInvoiceDetailSummary & InvoicesDetails(i, k, 2) 'line_desc
                    sCurrentCustomerInvoiceDetailSummary = sCurrentCustomerInvoiceDetailSummary & InvoicesDetails(i, k, 3) 'line_rate
                    sCurrentCustomerInvoiceDetailSummary = sCurrentCustomerInvoiceDetailSummary & InvoicesDetails(i, k, 4) 'line_amount
                    sCurrentCustomerInvoiceDetailSummary = sCurrentCustomerInvoiceDetailSummary & InvoicesDetails(i, k, 5) 'line_itemref_fullname
                Next k
                Command3.Enabled = True
            Else
                Command3.Enabled = False
            End If
        End If
    Next i
    Exit Sub
    
Err:
    frmMain.prcLogIt sUser, "prcInitCustomerInvoiceInfo:" & Err.Number & vbNewLine & Err.Description
            
End Sub


Sub prcDefaultFaxVarsReset()
    Dim i As Integer
    
    For i = 0 To UBound(sDefaultFaxVars) - 1
        sCurrentFaxVars(i) = sDefaultFaxVars(i)
    Next
End Sub


Sub prcDefaultModConVarsReset()
    Dim i As Integer

    For i = 0 To UBound(sDefaultModConAddress)
        If i = 0 Then
            Combo4.Text = sDefaultModConAddress(i, 0)
        End If
        sCurrentModConAddress(i, 0) = sDefaultModConAddress(i, 0)
        sCurrentModConAddress(i, 1) = sDefaultModConAddress(i, 1)
        Combo4.AddItem sDefaultModConAddress(i, 0)
    Next i
    
End Sub


Sub prcDefaultCustomerVarsReset()
    
    sCurrentCustomerCompany = sDefaultCustomerCompany
    sCurrentCustomerContact = sDefaultCustomerContact
    sCurrentCustomerPhone = sDefaultCustomerPhone
    sCurrentCustomerTimestamp = sDefaultCustomerTimestamp
    sCurrentCustomerDueDate = sDefaultCustomerDueDate

    sCurrentCustomerBillingAddress = sDefaultCustomerBillingAddress
    sCurrentCustomerBillingAddressWeb = sDefaultCustomerBillingAddress 'Replace(sDefaultCustomerBillingAddress, vbNewLine, "<br>")
    
    sCurrentCustomerShippingAddress = sDefaultCustomerShippingAddress
    sCurrentCustomerShippingAddressWeb = sDefaultCustomerShippingAddress 'Replace(sDefaultCustomerShippingAddress, vbNewLine, "<br>")
    
End Sub


Sub prcDisplayModConVars()
    Dim i As Integer
    
    For i = 0 To UBound(sCurrentModConAddress) - 1
        If Trim(Combo4.Text) = sCurrentModConAddress(i, 0) Then
            Text6.Text = sCurrentModConAddress(i, 1)
        End If
    Next i
End Sub


Sub prcDisplayCustomerVars()
    
    Text1.Text = sCurrentCustomerCompany
    Text2.Text = sCurrentCustomerContact
    Text8.Text = sCurrentCustomerBillingAddress
    Text9.Text = sCurrentCustomerDueDate
    Text10.Text = sCurrentCustomerTimestamp
    
End Sub

Sub prcUpdateCustomerVars()
    Command9.Enabled = False
    sCurrentCustomerCompany = Trim(Text1.Text)
    sCurrentCustomerContact = Trim(Text2.Text)
    sCurrentCustomerBillingAddress = Replace(Trim(Text8.Text), vbNewLine, "<br>")
    sCurrentCustomerDueDate = Trim(Text9.Text)
    sCurrentCustomerTimestamp = Trim(Text10.Text)
    Command9.Enabled = True
End Sub

'Sub prcPopulateGeneralLetterList()
'    Dim Response
'    Dim cmdCommand              As New ADODB.Command
'    Dim parParameter            As New ADODB.Parameter
'    Dim rsGrabLetterNames           As New ADODB.Recordset
'    Dim i As Integer
'
'On Error GoTo errHandle:'

    'MousePointer = vbHourglass
    'SQL_ReConnect_old frmMain.cnMC
    'If frmMain.cnMC.State <> 1 Then
    '    Exit Sub
    'End If
    
    'Set cmdCommand.ActiveConnection = frmMain.cnMC
    'cmdCommand.CommandType = adCmdText
    'cmdCommand.CommandText = " select * from qbx_letter where letter_active = '1' "
        
    'Set rsGrabLetterNames = cmdCommand.Execute
        
    'If Not rsGrabLetterNames.EOF Then
    '    Combo2.Clear
    '    i = 0
    '    ReDim sDefaultLetterParts(rsGrabLetterNames.RecordCount, 9)
    '    iLetterCount = rsGrabLetterNames.RecordCount
    '    rsGrabLetterNames.MoveFirst
    '    Combo2.Text = Trim(rsGrabLetterNames!letter_name)
    '    While Not rsGrabLetterNames.EOF
    '        Combo2.AddItem Trim(rsGrabLetterNames!letter_name)
    '        sDefaultLetterParts(i, 0) = Trim(rsGrabLetterNames!letter_name) & ""
    '        sDefaultLetterParts(i, 1) = Trim(rsGrabLetterNames!letter_description) & ""
    '        sDefaultLetterParts(i, 2) = Trim(rsGrabLetterNames!letter_part_header) & ""
    '        sDefaultLetterParts(i, 3) = Trim(rsGrabLetterNames!letter_part_addresse) & ""
    '        sDefaultLetterParts(i, 4) = Trim(rsGrabLetterNames!letter_part_body) & ""
    '        sDefaultLetterParts(i, 5) = Trim(rsGrabLetterNames!letter_part_closing) & ""
    '        sDefaultLetterParts(i, 6) = Trim(rsGrabLetterNames!letter_part_footer) & ""
    '        sDefaultLetterParts(i, 7) = Trim(rsGrabLetterNames!letter_part_addon_after) & ""
    '        sDefaultLetterParts(i, 8) = Trim(rsGrabLetterNames!letter_part_addon_message) & ""
    '        i = i + 1
    '        rsGrabLetterNames.MoveNext
    '    Wend
    'End If
    
    'Set rsGrabLetterNames = Nothing
    'Set parParameter = Nothing
    'Set cmdCommand = Nothing
    'MousePointer = vbDefault
    'Exit Sub
    
'errHandle:
'    Select Case (Err.Number)
'        Case Else
'            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "rsGrabLetterNames")
'            If Response = vbYes Then Resume Else Exit Sub
'    End Select
'    Set parParameter = Nothing
'    Set cmdCommand = Nothing
'    Screen.MousePointer = vbDefault
'End Sub



Sub prcPullLetterPartsFromName()
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim rsGrabLetterNames           As New ADODB.Recordset
    Dim i As Integer
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_letter where letter_name = '" & sCurrentLetterName & "' "
        
    Set rsGrabLetterNames = cmdCommand.Execute
        
    If rsGrabLetterNames.RecordCount > 0 Then
        ReDim sCurrentLetterPartsToModify(7, 2)
        ReDim sCurrentLetterPartsBackground(8, 2)
        rsGrabLetterNames.MoveFirst
        
        sCurrentLetterTitle = Trim(rsGrabLetterNames!letter_title) & ""
        
        sCurrentLetterPartsBackground(0, 0) = "start-NoV"
        sCurrentLetterPartsBackground(0, 1) = Trim(rsGrabLetterNames!letter_part_start) & ""
        sCurrentLetterPartsToModify(0, 0) = "ModCon Address"
        sCurrentLetterPartsToModify(0, 1) = Trim(rsGrabLetterNames!letter_part_header) & ""
        
        sCurrentLetterPartsBackground(1, 0) = "date-NoV"
        sCurrentLetterPartsBackground(1, 1) = Trim(rsGrabLetterNames!letter_part_date_between) & ""
        sCurrentLetterPartsToModify(1, 0) = "Date"
        sCurrentLetterPartsToModify(1, 1) = Trim(rsGrabLetterNames!letter_part_date) & ""
        
        sCurrentLetterPartsBackground(2, 0) = "addresse-NoV"
        sCurrentLetterPartsBackground(2, 1) = Trim(rsGrabLetterNames!letter_part_addresse_between) & ""
        sCurrentLetterPartsToModify(2, 0) = "Customer Address"
        sCurrentLetterPartsToModify(2, 1) = Trim(rsGrabLetterNames!letter_part_addresse) & ""
        
        sCurrentLetterPartsBackground(3, 0) = "dear-NoV"
        sCurrentLetterPartsBackground(3, 1) = Trim(rsGrabLetterNames!letter_part_dear_between) & ""
        sCurrentLetterPartsToModify(3, 0) = "Dear"
        sCurrentLetterPartsToModify(3, 1) = Trim(rsGrabLetterNames!letter_part_dear) & ""
        
        sCurrentLetterPartsBackground(4, 0) = "body-NoV"
        sCurrentLetterPartsBackground(4, 1) = Trim(rsGrabLetterNames!letter_part_body_between) & ""
        sCurrentLetterPartsToModify(4, 0) = "Letter Body"
        sCurrentLetterPartsToModify(4, 1) = Trim(rsGrabLetterNames!letter_part_body) & ""
        
        sCurrentLetterPartsBackground(5, 0) = "sincerely-NoV"
        sCurrentLetterPartsBackground(5, 1) = Trim(rsGrabLetterNames!letter_part_closing_between) & ""
        sCurrentLetterPartsToModify(5, 0) = "Sincerely"
        sCurrentLetterPartsToModify(5, 1) = Trim(rsGrabLetterNames!letter_part_closing) & ""
        
        sCurrentLetterPartsBackground(6, 0) = "footer-NoV"
        sCurrentLetterPartsBackground(6, 1) = Trim(rsGrabLetterNames!letter_part_footer_between) & ""
        sCurrentLetterPartsToModify(6, 0) = "Closing"
        sCurrentLetterPartsToModify(6, 1) = Trim(rsGrabLetterNames!letter_part_footer) & ""
        
        sCurrentLetterPartsBackground(7, 0) = "end-NoV"
        sCurrentLetterPartsBackground(7, 1) = Trim(rsGrabLetterNames!letter_part_end) & ""
        
        'sCurrentLetterPartsToModify(12, 0) = "name"
        'sCurrentLetterPartsToModify(12, 1) = Trim(rsGrabLetterNames!letter_part_addon_after) & ""
        'sCurrentLetterPartsToModify(13, 0) = "name"
        'sCurrentLetterPartsToModify(13, 1) = Trim(rsGrabLetterNames!letter_part_addon_message) & ""
        'sCurrentLetterPartsToModify(0, 0) = "name"
        'sCurrentLetterPartsToModify(0, 1) = Trim(rsGrabLetterNames!letter_name) & ""
        'sCurrentLetterPartsToModify(1, 0) = "name"
        'sCurrentLetterPartsToModify(1, 1) = Trim(rsGrabLetterNames!letter_description) & ""
        
    End If
    
    Set rsGrabLetterNames = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "rsGrabLetterNames")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub


Function funDisplayLetterPreview() As String
    Dim i As Integer
    Dim sLetterPreview As String
    
    sLetterPreview = ""
    For i = 0 To UBound(sCurrentLetterPartsToModify) - 1
        sLetterPreview = sLetterPreview & Trim(sCurrentLetterPartsBackground(i, 1))
        sLetterPreview = sLetterPreview & Trim(sCurrentLetterPartsToModify(i, 1))
    Next i
    sLetterPreview = sLetterPreview & Trim(sCurrentLetterPartsBackground(i, 1))
    
    sLetterPreview = funConvertVariablesToData(sLetterPreview)
    
    funDisplayLetterPreview = sLetterPreview
End Function

Sub prcDisplayLetterVariables()
    Combo7.Clear
    Combo7.Text = "$CustomerName"
    Combo7.AddItem "$CustomerName"
    Combo7.AddItem "$CustomerContact"
    Combo7.AddItem "$CustomerAddress"
    Combo7.AddItem "$CustomerDateStamp"
    Combo7.AddItem "$CustomerDueDate"
    
    Combo7.AddItem "$NumberOfCurrentInvoice"
    Combo7.AddItem "$BalanceOfCurrentInvoice"
    
    Combo7.AddItem "$CountOfInvoicesOverDue"
    Combo7.AddItem "$ListOfInvoicesOverDue"
    Combo7.AddItem "$ListOfInvoicesOverDueWithDetails"
    Combo7.AddItem "$BalanceOfInvoicesOverDue"
    
    Combo7.AddItem "$CountOfAllInvoices"
    Combo7.AddItem "$ListOfAllInvoices"
    Combo7.AddItem "$BalanceOfAllInvoices"
    
    Combo7.AddItem "$ModConAddress1"
    Combo7.AddItem "$ModConAddress2"
End Sub

Function funConvertVariablesToData(sLetterVariables As String) As String

    sLetterVariables = Replace(sLetterVariables, "$CustomerName", sCurrentCustomerCompany)
    sLetterVariables = Replace(sLetterVariables, "$CustomerContact", sCurrentCustomerContact)
    sLetterVariables = Replace(sLetterVariables, "$CustomerAddress", sCurrentCustomerBillingAddress)
    sLetterVariables = Replace(sLetterVariables, "$CustomerDateStamp", sCurrentCustomerTimestamp)
    sLetterVariables = Replace(sLetterVariables, "$CustomerDueDate", sCurrentCustomerDueDate)

    sLetterVariables = Replace(sLetterVariables, "$NumberOfCurrentInvoice", sCurrentCustomerInvoiceNum)
    sLetterVariables = Replace(sLetterVariables, "$BalanceOfCurrentInvoice", sCurrentCustomerBalanceRemaining)
    
    sLetterVariables = Replace(sLetterVariables, "$CountOfInvoicesOverDue", sCountOfInvoicesOverDue)
    sLetterVariables = Replace(sLetterVariables, "$ListOfInvoicesOverDueWithDetails", sListOfInvoicesOverDueWithDetails)
    sLetterVariables = Replace(sLetterVariables, "$ListOfInvoicesOverDue", sListOfInvoicesOverDue)
    sLetterVariables = Replace(sLetterVariables, "$BalanceOfInvoicesOverDue", sBalanceOfInvoicesOverDue)
    
    sLetterVariables = Replace(sLetterVariables, "$CountOfAllInvoices", sCurrentCustomerNumberOfInvoices)
    sLetterVariables = Replace(sLetterVariables, "$ListOfAllInvoices", sCurrentCustomerListOfInvoices)
    sLetterVariables = Replace(sLetterVariables, "$BalanceOfAllInvoices", sCurrentCustomerTotalInvoiceBalance)
        
    sLetterVariables = Replace(sLetterVariables, "$ModConAddress1", sCurrentModConAddress(0, 1))
    sLetterVariables = Replace(sLetterVariables, "$ModConAddress2", sCurrentModConAddress(1, 1))
    sLetterVariables = Replace(sLetterVariables, "$ModConImage1", sCurrentModConAddress(2, 1))
    sLetterVariables = Replace(sLetterVariables, "$ModConFooter1", sCurrentModConAddress(3, 1))
    
    'sLetterVariables = Replace(sLetterVariables, "$CurrentLetterTitle", sCurrentLetterTitle)
    sLetterVariables = Replace(sLetterVariables, "$CurrentLetterTitle", "")
    
    'fax vars to convert
    Dim i As Integer
       
    For i = 0 To UBound(sCurrentFaxVars)
        
            If i > 9 And i < 15 Then
                If sCurrentFaxVars(i) = "1" Then
                    sCurrentFaxVars(i) = " checked "
                ElseIf sCurrentFaxVars(i) = "0" Then
                    sCurrentFaxVars(i) = ""
                End If
            Else
                If sCurrentFaxVars(i) = "" Then
                    sCurrentFaxVars(i) = "&nbsp;"
                End If
            End If
        
        sLetterVariables = Replace(sLetterVariables, "$sCurrentFaxVars(" & i & ")", sCurrentFaxVars(i))
        
            If i > 9 And i < 15 Then
                If sCurrentFaxVars(i) = " checked " Then
                    sCurrentFaxVars(i) = "1"
                ElseIf sCurrentFaxVars(i) = "" Then
                    sCurrentFaxVars(i) = "0"
                End If
            Else
                If sCurrentFaxVars(i) = "&nbsp;" Then
                    sCurrentFaxVars(i) = ""
                End If
            End If
    Next i
    
    sLetterVariables = Replace(sLetterVariables, vbNewLine, "<br>")
    Dim stemp As String
    
    stemp = "l""""k"
    stemp = Replace(stemp, """""", "'")
    sLetterVariables = Replace(sLetterVariables, """""", "'")
    
    funConvertVariablesToData = sLetterVariables
End Function

Sub prcDisplayletterBreakDownList()
    Dim i As Integer
    Combo3.Clear
    For i = 0 To UBound(sCurrentLetterPartsToModify) - 1
        If i = 0 Then
            sCurrentLetterPartModifying = Trim(sCurrentLetterPartsToModify(i, 0))
            Combo3.Text = sCurrentLetterPartModifying
        End If
        Combo3.AddItem Trim(sCurrentLetterPartsToModify(i, 0))
    Next i
    
End Sub

Sub prcDisplayLetterBreakDownPart()
    Dim i As Integer
    
    For i = 0 To UBound(sCurrentLetterPartsToModify)
        If sCurrentLetterPartModifying = Trim(sCurrentLetterPartsToModify(i, 0)) Then
            Text3.Text = Trim(sCurrentLetterPartsToModify(i, 1))
            Exit Sub
        End If
    Next i
        
End Sub



Sub prcGrabLetterNames()
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim rsGrabLetterNames           As New ADODB.Recordset
    Dim sSQL                    As String
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    sSQL = " select * from qbx_letter where letter_active = '1' or letter_title = 'Fax Cover' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    Set rsGrabLetterNames = cmdCommand.Execute
        
    If Not rsGrabLetterNames.EOF Then
        Combo2.Clear
        rsGrabLetterNames.MoveFirst
        Combo2.Text = Trim(rsGrabLetterNames!letter_name) & ""
        sCurrentLetterName = Trim(rsGrabLetterNames!letter_name) & ""
        While Not rsGrabLetterNames.EOF
            Combo2.AddItem Trim(rsGrabLetterNames!letter_name) & ""
            rsGrabLetterNames.MoveNext
        Wend
    End If
    
    Set rsGrabLetterNames = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "rsGrabLetterNames")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub




Sub CreateAfile(sPath As String, sFilename As String, sPage As String)
    Dim fs, f, A
    Dim strTemp As String

On Error GoTo errhandler
    
    'If sPage = "threat" Then
    '    strTemp = funCollectionThreatLetter
    'ElseIf sPage = "" Then
    '    strTemp = funInvoiceLetter
    'ElseIf sPage = "kill" Then
    '    strTemp = "."
    'End If
    If Trim(Combo2.Text) = "Invoice" Then
        strTemp = funInvoiceLetter
    ElseIf Trim(Combo2.Text) = "Facsimile Transmittal Sheet" Then
        strTemp = funFaxLetter
    Else
        strTemp = funDisplayLetterPreview
    End If
    
    sEmailAttachmentMessage = strTemp
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set A = fs.CreateTextFile(sPath & sFilename, True)
    
    A.WriteLine (strTemp)
    A.Close
    Exit Sub
    
errhandler:
    frmMain.prcLogIt sUser, "CreateAfile:" & Err.Number & vbNewLine & Err.Description
    'MsgBox "CreateAfile:" & Err.Number & vbNewLine & Err.Description
'f = FreeFile
    'Open sPath & "CreateFileError.txt" For Append As #f
    'Print #f, Now & " Err = " & Err.Description & " " & Err.number
    'f.Close
    
End Sub



Sub prcCreateOrderFormFile(sPath As String, sFilename As String, sPage As String, sFileSave As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Dim fs, f
    Dim strTemp As String

On Error GoTo errhandler
    
    'If sPage = "threat" Then
    '    strTemp = funCollectionThreatLetter
    'ElseIf sPage = "" Then
    '    strTemp = funInvoiceLetter
    'ElseIf sPage = "kill" Then
    '    strTemp = "."
    'End If
    If Trim(Combo2.Text) = "Invoice" Then
        strTemp = funInvoiceLetter
    ElseIf Trim(Combo2.Text) = "Facsimile Transmittal Sheet" Then
        strTemp = funFaxLetter
    Else
        strTemp = funDisplayLetterPreview
    End If
    
    sEmailAttachmentMessage = strTemp
    
    'ElseIf Trim(Combo2.Text) = "General Collections" Then
    '    strTemp = funCollectionThreatLetter
    '    Frame3.Enabled = False
    '    Frame4.Enabled = False
    
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If sFileSave = "input" Then
        Set f = fs.OpenTextFile(sPath & sFilename, ForWriting, 0)
    'Else
        'Set f = fs.OpenTextFile(sPath & sFilename, ForAppending, TristateFalse)
    End If
    
    f.Write strTemp
    'f.Close
    Exit Sub
    
errhandler:
    frmMain.prcLogIt sUser, "prcCreateOrderFormFile:" & Err.Number & vbNewLine & Err.Description
    'MsgBox "prcCreateOrderFormFile:" & Err.Number & vbNewLine & Err.Description
'f = FreeFile
    'Open sPath & "CreateFileError.txt" For Append As #f
    'Print #f, Now & " Err = " & Err.Description & " " & Err.number
    'f.Close
    
End Sub



Sub prcPreviewLetter(sPage As String)

    Dim iFileExistance As Boolean
    Dim sPath As String
    Dim sFilename As String
    Dim sTempFileName As String
    Dim i As Integer
    Dim iFileGood As Integer
    Dim strFullPath As String
    
On Error GoTo Err:
    
    sPath = "c:\"
    i = 1
    iFileGood = 0
        
    sTempFileName = "doc.html"
    'Check File exists
    strFullPath = sPath & sTempFileName
    iFileExistance = DoesFileExist(strFullPath)
    
    If iFileExistance = True Then
        prcCreateOrderFormFile sPath, sTempFileName, sPage, "input"
    Else
        CreateAfile sPath, sTempFileName, sPage
    End If
        
    'check file exists
    iFileExistance = DoesFileExist(strFullPath)
    
    If iFileExistance = True Then
        'display file
        sFilename = strFullPath
        WB1.navigate sFilename
    End If
    Exit Sub
    
Err:
    frmMain.prcLogIt sUser, "prcPreviewLetter:" & Err.Number & vbNewLine & Err.Description
    'MsgBox "prcPreviewLetter:" & Err.Number & vbNewLine & Err.Description
    
End Sub




Function funCollectionThreatLetter() As String
    Dim sCollectionLetter As String
    
    '<img width=400 height=90 src=""Z:\Software\internal\images\modernconsumer_c.gif"">
    
    sCollectionLetter = "<html><head><meta http-equiv=Content-Type content=""text/html; charset=windows-1252""><meta name=Generator content=""Microsoft Word 11 (filtered)""><title>Collections</title></head><body lang=EN-US link=blue vlink=purple>" _
                        & "<div class=Section1><div style='border-top:double windowtext 2.25pt;border-left:none;border-bottom:double windowtext 2.25pt;border-right:none;padding:1.0pt 0in 1.0pt 0in'>" _
                        & "<center><p class=DocumentLabel style='FONT: bold;FONT-SIZE:14px;border:none;padding:0in'><img width=400 height=90 src=" & sGImage_Modcon_C & "><br>270 Lafayette street · suite 201 · New york · ny 10012</p></center>" _
                        & "</div><p class=MsoNormal>" & CustomerMainInfo(7) & "</p><p class=MsoNormal>" & sBilling & "</p><p class=MsoNormal>Re: Payment due to Modern Consumer $" & CustomerMainInfo(6) & "</p><p class=MsoNormal>" & CustomerMainInfo(2) & ":</p>" _
                        & "<p class=MsoNormal> <b>" & CustomerMainInfo(1) & "</b>  is late on payment to Modern Consumer.  Unless we receive payment in full by " & CustomerMainInfo(8) & " we must place the account into collections.  Prevent this by sending your check for payment in full.  You may send payment overnight using  Modern Consumer’s FedEx account number 2690-7080-3.</p>" _
                        & "<p class=MsoNormal>Your dealership has a signed agreement with Modern Consumer to pay for Driverloans.com leads sent to them through our lead management system at <a href=""http://dealers.driverloans.com/"">http://dealers.driverloans.com</a>. You may call 866-387-9616 for your dealer ID and password.</p><p class=MsoNormal>Please direct all inquiries to Ines Cedeno in accounting at 646-442-0111 or 866-387-9616 x0111.</p><p class=MsoNormal>Sincerely,</p><p class=MsoNormal>Modern Consumer LLC</p></div><p class=MsoNormal>&nbsp;</p><p class=MsoNormal>&nbsp;</p><p class=MsoNormal>&nbsp;</p><p class=MsoNormal>&nbsp;</p><p class=MsoNormal>&nbsp;</p><div class=Section1>" _
                        & "<div style='border-top:double windowtext 2.25pt;border-left:none;border-right:none;padding:2.0pt 0in 1.0pt 0in'><center><p class=DocumentLabel style='FONT: bold;FONT-SIZE:10px;border:none;padding:0in'>212-387-9616 · Toll Free 866-387-9616 · Fax 212-214-0561</p></center></div></body></html>"

    funCollectionThreatLetter = sCollectionLetter
            
            
    '<p class=MsoNormal>&nbsp;</p>
End Function


Sub prcGrabInvoice()
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim rsGrabInvoice           As New ADODB.Recordset
    Dim j As Integer
    Dim strXtraSQL As String
    Dim sTempDueDate As Date
    Dim sTempNow As Date
    Dim dTempBalance As Double
    
On Error GoTo errHandle:
    iInvBoolmark = 0

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    If sProfileAttrDtlsAry(1, 44) = 0 Then
        strXtraSQL = " and sign(inv_balanceremaining) != '-1' "
    End If
    
    If sProfileAttrDtlsAry(1, 46) = 0 Then
        strXtraSQL = strXtraSQL & " and inv_balanceremaining <> '0.00' "
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_inv where inv_customerref_listid = '" & ListID & "' " & strXtraSQL
        
    Set rsGrabInvoice = cmdCommand.Execute
        
    If Not rsGrabInvoice.EOF Then
        sCountOfInvoicesOverDue = 0
        dTempBalance = 0
        sListOfInvoicesOverDueWithDetails = ""
        iInvCount = rsGrabInvoice.RecordCount
        ReDim InvoicesInfo(iInvCount, 15)
        ReDim InvoicesDetails(iInvCount, 30, 7)
        j = 0
        rsGrabInvoice.MoveFirst
        
        InvoiceNumTemp = Trim(rsGrabInvoice!inv_refnumber)
        Combo1.Text = InvoiceNumTemp
        While Not rsGrabInvoice.EOF

            InvoiceNumTemp = Trim(rsGrabInvoice!inv_refnumber)
            'items in the invoices: date, invoice #, terms, due date, rep, account, total, payments/credits, current balance, total balance
            InvoicesInfo(j, 0) = Trim(rsGrabInvoice!inv_txndate) & ""
            InvoicesInfo(j, 1) = InvoiceNumTemp
            If j = 0 Then
                sCurrentCustomerListOfInvoices = sCurrentCustomerListOfInvoices & "(" & InvoiceNumTemp
            Else
                sCurrentCustomerListOfInvoices = sCurrentCustomerListOfInvoices & ", " & InvoiceNumTemp
            End If
            Combo1.AddItem InvoiceNumTemp
            InvoicesInfo(j, 2) = Trim(rsGrabInvoice!inv_termsref_fullname) & ""
            InvoicesInfo(j, 3) = Trim(rsGrabInvoice!inv_duedate) & ""
            InvoicesInfo(j, 4) = Trim(rsGrabInvoice!inv_salesrepref_fullname) & ""
            InvoicesInfo(j, 5) = Trim(rsGrabInvoice!inv_araccountref_fullname) & ""
            InvoicesInfo(j, 6) = Trim(rsGrabInvoice!inv_appliedamount) & ""
            InvoicesInfo(j, 7) = Trim(rsGrabInvoice!inv_balanceremaining) & ""
            InvoicesInfo(j, 8) = Trim(rsGrabInvoice!inv_customermsgref_fullName) & ""
            InvoicesInfo(j, 9) = Trim(rsGrabInvoice!inv_subtotal) & ""
            
            'checking for overdue invoices
            sTempDueDate = InvoicesInfo(j, 3)
            sTempNow = Format(Now, "M/d/YYYY")
            If sTempNow > sTempDueDate Then
                sCountOfInvoicesOverDue = sCountOfInvoicesOverDue + 1
                dTempBalance = dTempBalance + InvoicesInfo(j, 7)
                If sCountOfInvoicesOverDue = 1 Then
                    sListOfInvoicesOverDue = "(" & InvoiceNumTemp
                    sListOfInvoicesOverDueWithDetails = sListOfInvoicesOverDueWithDetails & "<table><tr><td width=60 align=center>Inv. No.</td><td width=80 align=center>Inv. Date</td><td width=80 align=center>Due Date</td><td width=100 align=right>Inv. Amount</td><td width=100 align=right>Balance</td></tr>" & _
                        "<tr><td align=center>" & InvoiceNumTemp & "</td><td align=center>" & InvoicesInfo(j, 0) & "</td><td align=center>" & InvoicesInfo(j, 3) & "</td><td align=right>$" & InvoicesInfo(j, 7) & "</td><td align=right>$" & InvoicesInfo(j, 7) & "</td></tr>"
                Else
                    sListOfInvoicesOverDue = ", " & sListOfInvoicesOverDue
                    sListOfInvoicesOverDueWithDetails = sListOfInvoicesOverDueWithDetails & _
                        "<tr><td align=center>" & InvoiceNumTemp & "</td><td align=center>" & InvoicesInfo(j, 0) & "</td><td align=center>" & InvoicesInfo(j, 3) & "</td><td align=right>$" & InvoicesInfo(j, 7) & "</td><td align=right>$" & InvoicesInfo(j, 7) & "</td></tr>"
                End If
                'MsgBox sListOfInvoicesOverDueWithDetails
            End If
            
            
            
            iInvBoolmark = j
            prcGrabInvoiceLineItems rsGrabInvoice!inv_txnid
            
            j = j + 1
            rsGrabInvoice.MoveNext
        Wend
        sCurrentCustomerListOfInvoices = sCurrentCustomerListOfInvoices & ")"
        sListOfInvoicesOverDue = sListOfInvoicesOverDue & ")"
        
        sBalanceOfInvoicesOverDue = dTempBalance
        sBalanceOfInvoicesOverDue = funFormatDecimal(sBalanceOfInvoicesOverDue)
        
        If sListOfInvoicesOverDueWithDetails <> "" Then
            sListOfInvoicesOverDueWithDetails = sListOfInvoicesOverDueWithDetails & _
                "<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td align=left>Total OverDue:</td><td align=right>$" & sBalanceOfInvoicesOverDue & "</td></tr></table>"
        Else
            sListOfInvoicesOverDueWithDetails = "Empty"
        End If
                
    Else
        Combo1.Text = ""
        Combo1.Enabled = False
        'Option2.Value = True
        'Option1.Enabled = False
    End If
    
    Set rsGrabInvoice = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "prcGrabInvoice Bit Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub



Sub prcGrabInvoiceLineItems(txnId As String)
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim rsGrabInvoiceLineItems  As New ADODB.Recordset
    Dim k                       As Integer
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_inv_lineitems where inv_txnid_link = '" & txnId & "' "
        
    Set rsGrabInvoiceLineItems = cmdCommand.Execute
    
    If Not rsGrabInvoiceLineItems.EOF Then
    
        k = 0
        rsGrabInvoiceLineItems.MoveFirst
        
        While Not rsGrabInvoiceLineItems.EOF
        
            InvoicesDetails(iInvBoolmark, k, 0) = InvoiceNumTemp
            InvoicesDetails(iInvBoolmark, k, 1) = Trim(rsGrabInvoiceLineItems!inv_line_quantity)
            InvoicesDetails(iInvBoolmark, k, 2) = Trim(rsGrabInvoiceLineItems!inv_line_desc)
            InvoicesDetails(iInvBoolmark, k, 3) = Trim(rsGrabInvoiceLineItems!inv_line_rate)
            InvoicesDetails(iInvBoolmark, k, 4) = Trim(rsGrabInvoiceLineItems!inv_line_amount)
            InvoicesDetails(iInvBoolmark, k, 5) = Trim(rsGrabInvoiceLineItems!inv_line_itemref_fullname)
            InvoicesInfo(iInvBoolmark, 10) = Trim(rsGrabInvoiceLineItems.RecordCount)
            
            k = k + 1
            rsGrabInvoiceLineItems.MoveNext
        Wend
        
    End If
    
    Set rsGrabInvoiceLineItems = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Grab Customer Bit Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Function funFaxLetter() As String
    Dim sFaxLetter As String
    
On Error Resume Next
    
    sFaxLetter = "<html><head><meta http-equiv=Content-Type content=""text/html; charset=windows-1252""><meta name=Generator content=""Microsoft Word 11 (filtered)""><title>$CurrentLetterTitle</title></head><body lang=EN-US link=blue vlink=purple>" _
            & "<div class=Section1><div style=""border-top:double windowtext 2.25pt;border-left:none;border-bottom:double windowtext 2.25pt;border-right:none;padding:1.0pt 0in 1.0pt 0in""><p class=DocumentLabel style=""FONT: bold;FONT-SIZE:14px;border:none;padding:0in"">" _
            & "<center>$ModConImage1</p></div><br><br><br><br><div style=""FONT-SIZE: 14px;FONT-STYLE: Normal;FONT-FAMILY: Verdana, Arial, Helvetica;"">Facsimile Transmittal Sheet</div><br><br>" _
            & "<div style=""border-top:double windowtext 2.25pt;border-left:none;border-bottom:double windowtext 2.25pt;border-right:none;padding:1.0pt 0in 1.0pt 0in"">" _
            & "<table border=1 width=670>" _
            & "<tr><td width=330><b>TO:</b> $sCurrentFaxVars(0)</td><td width=10>&nbsp;</td><td width=330><b>FROM:</b> $sCurrentFaxVars(5)</td></tr><tr><td></td><td></td><td></td></tr><tr><td><b>COMPANY:</b> $sCurrentFaxVars(1)</td><td>&nbsp;</td><td><b>DATE:</b> $sCurrentFaxVars(6)</td></tr><tr><td></td><td></td><td></td></tr><tr><td><b>FAX NUMBER:</b> $sCurrentFaxVars(2)</td><td>&nbsp;</td><td><b>NO. PAGES:</b> $sCurrentFaxVars(7)</td></tr><tr><td></td><td></td><td></td></tr><tr><td><b>PHONE NUMBER:</b> $sCurrentFaxVars(3)</td><td>&nbsp;</td><td><b>SENDER'S PHONE:</b> $sCurrentFaxVars(8)</td></tr><tr><td></td><td></td><td></td></tr><tr><td><b>RE:</b> $sCurrentFaxVars(4)</td><td>&nbsp;</td><td><b>SENDER'S FAX:</b> $sCurrentFaxVars(9)</td></tr><tr><td></td><td></td><td></td></tr>" _
            & "</table></div><br>" _
            & "<table border=0 WIDTH=670><tr>" _
            & "<td><div style=""FONT-SIZE: 10px;FONT-STYLE: Normal;FONT-FAMILY: Arial, Helvetica;""><input type=""checkbox"" $sCurrentFaxVars(10)>Urgent</div></td>" _
            & "<td><div style=""FONT-SIZE: 10px;FONT-STYLE: Normal;FONT-FAMILY: Arial, Helvetica;""><input type=""checkbox"" $sCurrentFaxVars(11)>FOR REVIEW</div></td>" _
            & "<td><div style=""FONT-SIZE: 10px;FONT-STYLE: Normal;FONT-FAMILY: Arial, Helvetica;""><input type=""checkbox"" $sCurrentFaxVars(12)>PLEASE COMMENT</div></td>" _
            & "<td><div style=""FONT-SIZE: 10px;FONT-STYLE: Normal;FONT-FAMILY: Arial, Helvetica;""><input type=""checkbox"" $sCurrentFaxVars(13)>PLEASE REPLY</div></td>" _
            & "<td><div style=""FONT-SIZE: 10px;FONT-STYLE: Normal;FONT-FAMILY: Arial, Helvetica;""><input type=""checkbox"" $sCurrentFaxVars(14)>PLEASE RECYCLE</div></td></tr></table>" _
            & "<br><div style=""border-top:double windowtext 2.25pt;border-left:none;border-bottom:none 2.25pt;border-right:none;padding:1.0pt 0in 1.0pt 0in"">" _
            & "<table border=0 WIDTH=670 height=200><tr><td valign=top align=left><div style=""FONT-SIZE: 14px;FONT-STYLE: Normal;FONT-FAMILY: Arial, Helvetica;"">NOTES/COMMENTS:<BR>$sCurrentFaxVars(15)</div></td></tr></table></div><br><br><br><br><br>" _
            & "$ModConFooter1</center></body></html>"
            
    sFaxLetter = funConvertVariablesToData(sFaxLetter)
    
    funFaxLetter = sFaxLetter
    
End Function


Function funInvoiceLetter() As String
    Dim sInvoiceLetter As String
    Dim i As Integer
    
    If InvoicesInfo(sInvInfo, 2) = "" Then
        InvoicesInfo(sInvInfo, 2) = "&nbsp;"
    End If
    If InvoicesInfo(sInvInfo, 3) = "" Then
        InvoicesInfo(sInvInfo, 3) = "&nbsp;"
    End If
    If InvoicesInfo(sInvInfo, 4) = "" Then
        InvoicesInfo(sInvInfo, 4) = "&nbsp;"
    End If
    If InvoicesInfo(sInvInfo, 8) = "" Then
        InvoicesInfo(sInvInfo, 8) = "&nbsp;"
    End If
    
    
    
    sInvoiceLetter = "<html><head><title></title><style> .bdr {border-right-style: solid;border-right-width:1.0pt;} .bdb {border-bottom-style: solid;border-bottom-width: 1px;} .vars {FONT-SIZE: 12px; COLOR: black; FONT-FAMILY: Times New Roman;} .invoice {FONT-SIZE: 32px; COLOR: black; FONT-FAMILY: sans-serif; MARGIN-LEFT: 0px; MARGIN-RIGHT: 0px; FONT: bold;} .modcon {FONT-SIZE: 18px; COLOR: black; FONT-FAMILY: Times New Roman;} .closestatment {FONT-SIZE: 11px; COLOR: black; FONT-FAMILY: Times New Roman; FONT: bold;} .moneytitle {FONT-SIZE: 14px; COLOR: black; FONT-FAMILY: sans-serif; FONT: bold;}</style>" _
                    & "</head><body><table border=0 width=96%><tr><td width=50% align=left><img width=260 height=40 src=""Z:\Software\internal\images\modernconsumer_c.gif""><span class=modcon><br>270 Lafayette Street, Suite 201<br>New York, NY 10012<br>(212) 387-9616<br>(212) 214-0561</span></td><td align=right><span class=invoice>Invoice</span><br>" _
                    & "<!---address---><table border=1 width=50%><tr><td align=center>Date</td><td align=center>Invoice #</td></tr><tr><td align=center>" & InvoicesInfo(sInvInfo, 0) & "</td><td align=center>" & InvoicesInfo(sInvInfo, 1) & "</td></tr></table></td></tr><tr><td align=center><table border=1 width=80%><tr><td align=left>&nbsp;&nbsp;Bill To</td></tr><tr><td align=left valign=top height=100><span class=vars>" & sBilling & "</span></td></tr></table>" _
                    & "</td><td align=center><br><br><table border=1 width=80%><tr><td align=left>&nbsp;&nbsp;Deliver To</td></tr><tr><td align=left valign=top height=100><span class=vars>" & sShipping & "</span></td></tr></table><br><br></td></tr><tr><td colspan=2><table border=0 width=100% cellpadding=0 cellspacing=0><tr><td><table border=0 cellpadding=0 cellspacing=0 width=100%><tr><td width=300>&nbsp;</td><td>" _
                    & "<!---info---><table border=1 cellpadding=0 cellspacing=0 width=100%><tr><td align=center width=145>Terms</td><td align=center width=145>Due Date</td><td align=center width=145>Rep</td><td align=center width=145>Account #</td></tr><tr><td align=center><span class=vars>" & InvoicesInfo(sInvInfo, 2) & "</span></td><td align=center><span class=vars>" & InvoicesInfo(sInvInfo, 3) & "</span></td><td align=center><span class=vars>" & InvoicesInfo(sInvInfo, 4) & "</span></td><td align=center><span class=vars>" & sAccountNumber & "</span></td></tr></table>" _
                    & "</td></tr></table></td></tr><tr><td><table border=1 cellpadding=0 cellspacing=0 width=100%><tr><td><table border=0 cellpadding=0 cellspacing=0 width=100%><tr><td class=bdr align=center width=198>Item</td><td class=bdr align=center width=100>Quantity</td><td class=bdr align=center width=348>Description </td><td class=bdr align=center width=96>Rate</td><td align=center width=120>Amount</td></tr></table></td></tr><tr><td>" _
                    & "<!---items-------><table class=vars border=0 cellpadding=0 cellspacing=0 width=100%>"
                 
    '                & "</td></tr></table></td></tr><tr><td><table border=1 cellpadding=0 cellspacing=0 width=100%><tr><td><table border=0 cellpadding=0 cellspacing=0 width=100%><tr><td class=bdr align=center width=200>Item</td><td class=bdr align=center width=100>Quantity</td><td class=bdr align=center width=346>Description </td><td class=bdr align=center width=100>Rate</td><td align=center width=120>Amount</td></tr></table></td></tr><tr><td>" _

    'iInvBoolmark = 0
    'For i = 0 To iInvCount
    '    If Trim(InvoicesInfo(i, 1)) = Trim(Combo1.Text) Then
    '        iInvBoolmark = i
    '        i = iInvCount + 1
    '    End If
    'Next i


    If Trim(InvoicesInfo(sInvInfo, 10)) = 0 Then
        i = 0
    Else
        For i = 0 To Trim(InvoicesInfo(sInvInfo, 10)) - 1
            sInvoiceLetter = sInvoiceLetter & "<tr><td class=bdr align=center width=173>" & InvoicesDetails(sInvInfo, i, 5) & "</td><td class=bdr align=center width=120>" & InvoicesDetails(sInvInfo, i, 1) & "</td><td class=bdr align=center width=330>" & InvoicesDetails(sInvInfo, i, 2) & "</td><td class=bdr align=center width=90>" & InvoicesDetails(sInvInfo, i, 3) & "</td><td align=center width=120>" & InvoicesDetails(sInvInfo, i, 4) & "</td></tr>"
        Next
    End If
    
    For i = i To 22
        'sInvoiceLetter = sInvoiceLetter & "<tr><td align=center width=200>&nbsp;</td><td align=center width=100>&nbsp;</td><td align=center width=350>&nbsp;</td><td align=center width=100>&nbsp;</td><td align=center width=120>&nbsp;</td></tr>"
        sInvoiceLetter = sInvoiceLetter & "<tr><td class=bdr align=center>&nbsp;</td><td class=bdr align=center>&nbsp;</td><td class=bdr align=center>&nbsp;</td><td class=bdr align=center>&nbsp;</td><td align=center>&nbsp;</td></tr>"
    Next
    
    sInvoiceLetter = sInvoiceLetter & "</table><!---end items---->" _
                    & "</td></tr></table></td></tr><tr><td><table border=1 cellpadding=0 cellspacing=0 width=100%><tr><td width=500 height=42><span class=vars>" & InvoicesInfo(sInvInfo, 8) & "</span></td><td rowspan=2><table border=0 width=240 cellpadding=0 cellspacing=0><tr><td class=bdb align=left height=40 width=140><span class=moneytitle>Total</span></td><td class=bdb width=100 align=right><span class=vars>$" & InvoicesInfo(sInvInfo, 9) & "</span>&nbsp;</td></tr><tr><td class=bdb align=left height=40>" _
                    & "<span class=moneytitle>Payments/Credits</span></td><td  class=bdb align=right><span class=vars>$" & InvoicesInfo(sInvInfo, 6) & "&nbsp;<span></td></tr><tr><td class=bdb align=left height=40><span class=moneytitle>Current Balance</span></td><td class=bdb align=right><span class=vars>$" & InvoicesInfo(sInvInfo, 7) & "&nbsp;</span></td></tr><tr><td align=left height=40><span class=moneytitle>Total Balance</span></td><td align=right><span class=vars>$" & CustomerMainInfo(6) & "&nbsp;</span></td></tr>" _
                    & "</table></td></tr><tr><td><table border=0 width=100% cellpadding=0 cellspacing=0><tr><td align=left height=40><span class=closestatment>For verification, you can view your leads at http://dealers.driverloans.com.  Please <br>contact us for user ID and Password.</span>" _
                    & "</td></tr><tr><td align=left height=40><span class=closestatment>For billing questions, please call 646-442-0111.</span></td></tr><tr><td align=left height=40><span class=closestatment>Thank you!</span></td></tr></table></td></tr></table></td></tr></table></td></tr></table></body></html>"

    'InvoicesInfo(sInvInfo, 5) &

    funInvoiceLetter = sInvoiceLetter
            
End Function

Private Sub Form_Initialize()
    sFrmPrintPage = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Me.Width = 12015
    SSTab1.Height = Me.Height - 600
    WB1.Height = Me.Height - 1020
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'prcCreateOrderFormFile "c:\", "doc.html", "kill", "input"
    sFrmPrintPage = 0
End Sub

Private Sub Picture3_Click()
    iCalendarRequest = 3
    frmCalendar.Calendar1.Value = Trim(Text3.Text)
    Load frmCalendar
    frmCalendar.Show
End Sub

