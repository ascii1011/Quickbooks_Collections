VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmEmailCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "frmEmailCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11115
   Begin SHDocVwCtl.WebBrowser wb2 
      Height          =   4995
      Left            =   1080
      TabIndex        =   15
      Top             =   1320
      Width           =   9075
      ExtentX         =   16007
      ExtentY         =   8811
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
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Height          =   7515
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      Begin VB.CommandButton Command4 
         Caption         =   "Text"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3900
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Html"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   11
         Top             =   6420
         Width           =   9075
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   6900
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   7860
         Picture         =   "frmEmailCustomer.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6900
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   4995
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   1320
         Width           =   9075
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   960
         Width           =   9075
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1020
         TabIndex        =   3
         Top             =   600
         Width           =   9075
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   240
         Width           =   9075
      End
      Begin VB.Label Label5 
         Caption         =   "Attachment:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   6480
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Message:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Subject:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmailCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Response
    Dim sMsgStr As String
    
    Frame1.Enabled = False
    frmProcessing.Show
    
    sMsgStr = Replace(Trim(Text4.Text), vbNewLine, "<br>")
    
    If (InStr(1, Trim(Text2.Text), "@j2send.com") <> 0) Then
        Response = MsgBox("This email address will be sent to this customer as a fax." & vbNewLine & "Do you wish to continue?", vbExclamation + vbYesNo, "Email a fax (Confirmation).")
        If Response = vbYes Then
            SendOCEmail Trim(Text2.Text), Trim(Text1.Text), sMsgStr, "", Trim(Text3.Text)
            Unload Me
            Unload frmProcessing
        Else
            Unload frmProcessing
        End If
    Else
        SendOCEmail Trim(Text2.Text), Trim(Text1.Text), sMsgStr, "", Trim(Text3.Text)
        Unload Me
        Unload frmProcessing
    End If
    
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    wb2.Visible = True
    Text4.Visible = False
End Sub

Private Sub Command4_Click()
    Text4.Visible = True
    wb2.Visible = False
End Sub

Private Sub Form_Load()
    Me.Width = 11205
    Me.Height = 8055
    If sUser <> "" Then
        prcLoadVars
    Else
        MsgBox "Please logout of 'Modcon Collections' application and log back in."
        Unload Me
    End If
    
End Sub

Sub prcLoadVars()
    Dim stemp
        
    Text1.Text = sUser & "@modernconsumer.com"
    Text2.Text = Trim(frmInvoiceQry.Text10.Text)
    Text3.Text = "RE: Collections Message"
    'aryRep_Profile(0)
    stemp = vbNewLine & vbNewLine & _
        vbNewLine & aryRep_Profile(1) & _
        vbNewLine & "Email: " & aryRep_Profile(3) & _
        vbNewLine & "Phone: " & aryRep_Profile(4) & _
        vbNewLine & "Fax: " & aryRep_Profile(5) & _
        vbNewLine & "Collections Department" & _
        vbNewLine & "Modern Consumer LLC."
    
    'stemp = vbNewLine & vbNewLine & _
    '    vbNewLine & UCase(Left(sUser, 1)) & ". " & UCase(Right(sUser, (Len(sUser) - 1))) & _
    '    vbNewLine & sUser & "@modernconsumer.com" & _
    '    vbNewLine & "Collections Department" & _
    '    vbNewLine & "Modern Consumer LLC."
        
    If sEmailAttachmentMessage <> "" Then
        Text4.Text = sEmailAttachmentMessage & stemp
        prcPreviewLetter ""
        wb2.Visible = True
        Text4.Visible = False
    Else
        Text4.Text = stemp
        Command3.Enabled = False
        Text4.Visible = True
        wb2.Visible = False
    End If
    
    Label5.Visible = False
    Text5.Visible = False

End Sub


Sub prcPreviewLetter(sPage As String)
    Dim iFileExistance As Boolean
    Dim sPath As String
    Dim sFilename As String
    Dim sTempFileName As String
    Dim i As Integer
    Dim iFileGood As Integer
    Dim strFullPath As String
    
    sPath = "c:\"
    i = 1
    iFileGood = 0
        
    sTempFileName = "doc.html"
    'Check File exists
    strFullPath = sPath & sTempFileName
    iFileExistance = DoesFileExist(strFullPath)
            
    If iFileExistance = True Then
        'display file
        sFilename = strFullPath
        wb2.navigate sFilename
    Else
        Command3.Enabled = False
        Command4.Value = True
    End If
    
End Sub

Private Sub Form_Initialize()
    sFrmEmailCustomer = 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    sFrmEmailCustomer = 0
End Sub

