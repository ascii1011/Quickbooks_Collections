VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmEmailImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import E-Mail"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmEmailImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   7830
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Text            =   "Inbox"
         Top             =   2880
         Width           =   4815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Folder"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   2880
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   3600
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Import"
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   3600
         Width           =   915
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Delete Mail From Outlook Afterwards"
         Height          =   195
         Left            =   2460
         TabIndex        =   1
         Top             =   3780
         Width           =   3015
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2475
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4366
         _Version        =   393216
         Cols            =   5
         AllowUserResizing=   3
      End
      Begin VB.Label Label1 
         Caption         =   "ID:"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Current Folder:"
         Height          =   195
         Left            =   1320
         TabIndex        =   7
         Top             =   3000
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmEmailImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' late bind this object variable, since it could be various item types
Dim olTempItem As Object
Dim olTempSession As Object
    
Dim olJobSourceFolder As Outlook.mapiFolder
Dim olMoveFromFolder As Outlook.mapiFolder

Dim sTempCustomerListID As String
Dim sTmpCustInfo(7) As String


Private Sub Command1_Click()
    Dim sOlFolder As String
    
    sOlFolder = Trim(Text2.Text)
    
On Error GoTo Err:
    If sOlFolder <> "" Then
        Set olTempItem = olJobSourceFolder.Items(sOlFolder)
        prcImportMail olTempItem
    End If
    Exit Sub
    
Err:
    Exit Sub
End Sub

Sub prcImportMail(ByVal olItem)
    Dim sMsg As String
    
On Error GoTo EmailErr:
    
        sMsg = "<Email Entry>" & vbNewLine & "Date: " & Trim(olItem.CreationTime) & vbNewLine
        sMsg = sMsg & "From: " & Trim(olItem.SenderEmailAddress) & vbNewLine
        sMsg = sMsg & "To: " & sUser & vbNewLine
        sMsg = sMsg & "Subject: " & Trim(olItem.Subject) & vbNewLine
        sMsg = sMsg & "Body: " & Left(Trim(olItem.body) & vbNewLine, 2000)
            
    frmInvoiceQry.Text6.Text = sMsg
    
    If Check2.Value = 1 Then
        olItem.Delete
        prcDisplayFolder 2
        MsgBox "Email was deleted"
    End If
    
    Text2.Text = ""
    Check2.Value = 0
    'sTmpCustInfo(5) = sMsg
    
    'If sTmpCustInfo(5) <> "" Then
    '    frmInvoiceQry.prcInsertNote sTmpCustInfo
    '    frmInvoiceQry.prcGrabNote
    'End If
    
    Exit Sub

EmailErr:
    MsgBox "An Error occured while processing, please try again."
    Exit Sub
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    prcDisplayFolder 1
    'Text1.Text = prcGrabEmailfolder
End Sub

Function prcGrabEmailfolder()
    Dim olApp As Outlook.Application
    Dim olSession As Outlook.NameSpace
    Set olApp = Application
    Set olSession = olApp.GetNamespace("MAPI")
    Set olMoveFromFolder = olSession.PickFolder
    If olMoveFromFolder Is Nothing Then
    
    Else
        prcGrabEmailfolder = olMoveFromFolder
    End If
End Function

Private Sub FG1_Click()
    FG1.col = 0
    Text2.Text = Trim(FG1.Text)
End Sub

Private Sub Form_Load()
    sFrmEmailImport = 1
    sTmpCustInfo(0) = CustomerMainInfo(1)
    sTmpCustInfo(1) = CustomerMainInfo(6)
    sTmpCustInfo(2) = CustomerMainInfo(9)
    sTmpCustInfo(3) = ""
    sTmpCustInfo(4) = ""
    sTmpCustInfo(5) = ""
    sTmpCustInfo(6) = ListID
    Frame1.Caption = "Importing for (" & Trim(CustomerMainInfo(1)) & ")"
    'prcGrabDefaultFolder
    prcDisplayFolder 1
End Sub

Sub prcDisplayFolder(iSwitch As Integer)
    Dim lMainIndex As Long
    Dim lTotalCount As Long
    
    'declare folders
    'Dim olJobSourceFolder As Outlook.mapiFolder
    Dim ol As Outlook.Application
    Dim olSession As Outlook.NameSpace
    Dim MyFolder1
    
    Dim sEntryID As String
    
On Error GoTo MailErr:

    'init params
    Set ol = New Outlook.Application
    Set olSession = ol.GetNamespace("MAPI")
    
    If iSwitch = 1 Then
        Set olJobSourceFolder = olSession.PickFolder
        Set olTempSession = olJobSourceFolder
    ElseIf iSwitch = 2 Then
        Set olJobSourceFolder = olTempSession
    End If
    
    prcInitTables
    Text1.Text = ""
    Text2.Text = ""
    Check2.Value = 0
    
    Text1.Text = olJobSourceFolder
    lTotalCount = olJobSourceFolder.Items.Count
             
    'FG1.Rows = lTotalCount + 1
    
    For lMainIndex = 1 To lTotalCount
                
        If lMainIndex > lTotalCount Then
            Exit For
        End If
                    
        Set olTempItem = olJobSourceFolder.Items(lMainIndex)
                              
        FG1.Rows = lMainIndex + 1
        FG1.Row = lMainIndex
        FG1.col = 0
        FG1.Text = lMainIndex 'olTempItem.EntryID
        FG1.col = 1
        FG1.Text = olTempItem.CreationTime
        FG1.col = 2
        FG1.Text = olTempItem.SenderEmailAddress
        FG1.col = 3
        FG1.Text = olTempItem.Subject
        FG1.col = 4
        FG1.Text = olTempItem.body
        
    Next lMainIndex
                        
    
    Exit Sub
    
MailErr:
    prcInitTables
    Text1.Text = ""
    Text2.Text = ""
    Check2.Value = 0
    Exit Sub
End Sub









Sub prcInitTables()
    
    FG1.ColWidth(0) = 500
    FG1.ColWidth(1) = 1300
    FG1.ColWidth(2) = 1900
    FG1.ColWidth(3) = 2600
    FG1.ColWidth(4) = 500
    FG1.Rows = 1
    FG1.Row = 0
    FG1.col = 0
    FG1.Text = "ID"
    FG1.col = 1
    FG1.Text = "Date"
    FG1.col = 2
    FG1.Text = "From"
    FG1.col = 3
    FG1.Text = "Subject"
    FG1.col = 4
    FG1.Text = "Body"
    
End Sub












Private Sub Form_Unload(Cancel As Integer)
    sFrmEmailImport = 0
End Sub
