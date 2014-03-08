VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmReporting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   Icon            =   "frmReporting.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10800
   ScaleWidth      =   11205
   Begin VB.Frame Frame3 
      Caption         =   "Criteria:"
      Height          =   1935
      Left            =   60
      TabIndex        =   23
      Top             =   2160
      Width           =   11055
      Begin VB.CommandButton Command10 
         Caption         =   "Edit"
         Height          =   315
         Left            =   2760
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command8 
         Caption         =   "New"
         Height          =   315
         Left            =   1440
         TabIndex        =   30
         Top             =   1320
         Width           =   915
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Height          =   315
         Left            =   2760
         TabIndex        =   29
         Top             =   1320
         Width           =   915
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   3780
         TabIndex        =   28
         Top             =   180
         Width           =   7095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Remove"
         Height          =   315
         Left            =   2760
         TabIndex        =   27
         Top             =   780
         Width           =   915
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmReporting.frx":0442
         Left            =   180
         List            =   "frmReporting.frx":045B
         TabIndex        =   25
         Text            =   "Rep"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Height          =   1620
         Left            =   3780
         TabIndex        =   32
         Top             =   180
         Width           =   7095
      End
      Begin VB.Label Label7 
         Caption         =   "Properties:"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Choose a property to customize your search by and click ""New""."
         Height          =   435
         Left            =   180
         TabIndex        =   24
         Top             =   360
         Width           =   2475
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1275
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Text            =   "frmReporting.frx":049B
      Top             =   9420
      Width           =   11055
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   4875
      Left            =   60
      TabIndex        =   6
      Top             =   4140
      Width           =   11055
      ExtentX         =   19500
      ExtentY         =   8599
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
   Begin VB.Frame Frame2 
      Caption         =   "Order Results By:"
      Height          =   1155
      Left            =   60
      TabIndex        =   13
      Top             =   1020
      Width           =   11055
      Begin VB.CheckBox Check2 
         Caption         =   "Put Notes at bottom"
         Height          =   195
         Left            =   7980
         TabIndex        =   33
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmReporting.frx":04A1
         Left            =   1440
         List            =   "frmReporting.frx":04AB
         TabIndex        =   20
         Text            =   "Ascending"
         Top             =   720
         Width           =   1275
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmReporting.frx":04C6
         Left            =   120
         List            =   "frmReporting.frx":04DF
         TabIndex        =   19
         Text            =   "Rep"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "View Total Balance"
         Height          =   195
         Left            =   5640
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   4875
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   315
         Left            =   2820
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         Height          =   315
         Left            =   3540
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   10020
         TabIndex        =   14
         Text            =   "3"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Show this amount of notes:"
         Height          =   195
         Left            =   7980
         TabIndex        =   21
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton Command3 
         Height          =   555
         Left            =   10320
         Picture         =   "frmReporting.frx":051F
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   555
      End
      Begin VB.CommandButton Command4 
         Height          =   555
         Left            =   10320
         Picture         =   "frmReporting.frx":0961
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Report"
         Height          =   375
         Left            =   5820
         TabIndex        =   10
         Top             =   300
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
         Height          =   375
         Left            =   7560
         TabIndex        =   9
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By User:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   555
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   540
         Width           =   1995
      End
   End
   Begin VB.Label Label4 
      Height          =   195
      Left            =   6900
      TabIndex        =   8
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Total Balance:   $"
      Height          =   195
      Left            =   5580
      TabIndex        =   7
      Top             =   9120
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Count:"
      Height          =   195
      Left            =   8400
      TabIndex        =   5
      Top             =   9120
      Width           =   555
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   9060
      TabIndex        =   4
      Top             =   9120
      Width           =   1815
   End
End
Attribute VB_Name = "frmReporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPartOne As String
Dim sPartTwo As String
Dim sPartThree As String
Dim sImportance As String
Dim sInvStr As String
Dim sTitleStr As String
Dim sWebStr As String

Dim dTxnDate3 As Date
Dim dTxnDate4 As Date

Dim rsGrabCustomers     As New ADODB.Recordset

Dim sSQLMatchCriteria   As String
Dim sRepList()            As String

Dim sFrameTypeE As Integer

Sub prcSetPageVars()

    Report.Tpl_Html = "<html><head><title></title>" & _
        "<style type=""text/css"">" & _
        ".cust {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 10px;font-weight: bold;}" & _
        ".inv {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
        ".msg {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
        ".msgbody {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
        ".title {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 12px;font-weight: bold;}" & _
        "</style></head><body>"
        
    Report.Tpl_Title = funGetTpl_Title
        
    Report.Tpl_Body = funGetTpl
    
End Sub

Function funGetTpl_Title() As String

        funGetTpl_Title = "<table border=1 width=100% align=center><tr><td><table border=0 width=100% align=center>"
        funGetTpl_Title = funGetTpl_Title & "<tr><td class=title align=left>Today: " & Format(Now, "mm/dd/yyyy") & "</td><td class=title align=left>Created By: " & sUser & "</td></tr>"
        If Option1.Value = False Then
            funGetTpl_Title = funGetTpl_Title & "<tr><td class=title align=left>Report by user: " & Trim(Combo1.Text) & "</td>"
        Else
            funGetTpl_Title = funGetTpl_Title & "<tr><td class=title align=left>Report by: All</td>"
        End If
        
        funGetTpl_Title = funGetTpl_Title & "<td class=title align=left>Dates: " & Report.t_dTxnDate4 & " - " & Format(Now, "mm/dd/yyyy") & "</td></tr>"
        funGetTpl_Title = funGetTpl_Title & "<tr><td colspan=2 class=title align=right>Record Count:</td><td class=title>" & ifound & "</td></tr></table></td></tr></table><br>"
            
End Function

Function funGetTpl_Type() As String

    funGetTpl = ""
    
    If Report.Tpl_Type = "normal" Then
        funGetTpl = "<table border=1 width=100% align=left>" & _
                        "<tr><td>{Cust}</td><td rowspan=2>{Notes}</td></tr>" & _
                        "<tr><td>" & _
                            "<table border=1 width=100% align=left><tr>" & _
                            "<td>{Imp}</td>" & _
                            "<td>{Inv}</td>" & _
                            "</table>" & _
                        "</td></tr>" & _
                    "</table>"
                    
    ElseIf Report.Tpl = "horiz-notes" Then
        funGetTpl = "<table border=1 width=100% align=left>" & _
                        "<tr><td>{Cust}</td></tr>" & _
                        "<tr><td>" & _
                            "<table border=1 width=100% align=left><tr>" & _
                            "<td>{Imp}</td>" & _
                            "<td>{Inv}</td>" & _
                            "</table>" & _
                        "</td></tr>" & _
                        "<tr><td>{Notes}</td></tr>" & _
                    "</table>"
                    
    End If

End Function


Sub prcProduceReport()

    Report.Tpl = "normal"
    prcGetCustomerRecords
    prcSetPageVars
    
End Sub

Sub prcGetCustomerRecords()
    
    prcGrabCustomers
    prcCycleThroughCustomers
    
End Sub

Sub prcSet_tpl_Main_vars()
    Report.t_numItems = rsGrabCustomers.RecordCount
    Report.t_ifound = 0
    Report.t_TBal = 0
    Label4.Caption = 0
    Report.t_dTxnDate4 = Format(Now, "mm/dd/yyyy")
    Report.t_dTxnDate3 = Format(Now, "mm/dd/yyyy")
End Sub

Sub prcClr_tpl_Cust_vars()
    Report.c_txnId = ""
    Report.c_Rep = ""
    Report.c_Name = ""
    Report.c_Bal = ""
    Report.c_Contact = ""
    Report.c_Phone = ""
    Report.c_Status = ""
    Report.c_WebStatus = ""
    Report.c_Importancelvl = ""
    Report.c_Upfront = ""
End Sub

Sub prcSet_tpl_Cust_vars()
    prcClr_tpl_Cust_vars
    Report.c_txnId = Trim(rsGrabCustomers!cust_listid)
    Report.c_Rep = Trim(rsGrabCustomers!cust_salesrepref_fullname)
    Report.c_Name = Trim(rsGrabCustomers!cust_fullname)
    Report.c_Bal = Trim(rsGrabCustomers!cust_totalbalance)
    Report.c_Contact = Trim(rsGrabCustomers!cust_contact)
    Report.c_Phone = Trim(rsGrabCustomers!cust_phone1)
    Report.c_Status = LCase(Trim(rsGrabCustomers!cust_jobstatus))
    Report.c_WebStatus = Trim(rsGrabCustomers!cust_webstatus)
    Report.c_Importancelvl = Trim(rsGrabCustomers!importance_name) & ""
    If Report.c_Importancelvl = "" Then
        Report.c_Importancelvl = "empty"
    End If
    Report.c_Upfront = Trim(rsGrabCustomers!importance_upfront) & ""
    If Report.c_Upfront = "" Then
        Report.c_Upfront = "empty"
    End If
End Sub


Sub prcCycleThroughCustomers()
    Dim sID As String

    If rsGrabCustomers.State = 1 Then
        prcSet_tpl_Main_vars
        
        If Not rsGrabCustomers.EOF Then
            rsGrabCustomers.MoveFirst
            While Not rsGrabCustomers.EOF
                sID = ""
                sID = Trim(rsGrabCustomers!cust_listid) & ""
                
                If sID <> "" Then
                    prcSet_tpl_Cust_vars
                    prcGetCustomerParts sID
                End If
                
                rsGrabCustomers.MoveNext
            Wend
        End If
    End If
    
End Sub

Sub prcGetCustomerParts(sID As String)
    

    'get customer info
    Report.Cust = pfunRptg_GetCust(sID)
    
    'get Importance info
    Report.Imp = pfunRptg_GetImp(sID)
        
    'get Invoice info
    Report.Inv = pfunRptg_GetInv(sID)
    
    'get note records....
    Report.Note = pfunRptg_GetNotes(sID, Reporting.iNoteCount)


End Sub

Function funRptgTmpl(tplID)



End Function



Sub prcGrabCustomers()
    Dim Response
    Dim cmdCommand          As New ADODB.Command
    Dim parParameter        As New ADODB.Parameter
    Dim sSQLXtra            As String
    Dim sOrderBy            As String
    Dim sWhichOrder             As String
    Dim sOrderBigParts()        As String
    Dim sOrderSmallParts()      As String
    Dim i As Integer, j As Integer
    Dim sCriteria               As String
    Dim sSimilarCriteria()      As String
    Dim sSimilarFound           As Integer

On Error GoTo errHandle:

    sSQLXtra = ""
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    Set rsGrabCustomers = Nothing
    
    If Combo1.Text <> "" And Option1.Value = False Then
        sSQLXtra = " and cust_salesrepref_fullname = '" & Trim(Combo1.Text) & "' "
    End If
    
    Dim sCriteriaSplit() As String
    Dim sCriteriaSubSplit() As String
    
    
    If List1.ListCount > 0 Then
        'sSimilarCriteria(i, 0) = sql statement
        'sSimilarCriteria(i, 1) = parameter
        'sSimilarCriteria(i, 2) = (1 = tree root, 0 = duplicate)
        'sSimilarCriteria(i, 3) = if sql statment is a "not" or "!" then change "or" to "and"
        ReDim sSimilarCriteria(List1.ListCount, 4)
        For i = 0 To List1.ListCount - 1
            'If i = 0 Then
            '    sCriteria = List1.List(i)
            'Else
                sCriteria = sCriteria & " and " & List1.List(i)
            'End If
        Next i
    End If
    
    
    'recursive split to find out all the fields to order by
    If Trim(Text1.Text) <> "" Then
    
        sOrderBigParts = Split(Trim(Text1.Text), ";")
        'If UBound(sOrderBigParts) > 0 Then
        
            For i = 0 To UBound(sOrderBigParts)
                'split small parts apart to input into sql statement
                
                sOrderSmallParts = Split(Trim(sOrderBigParts(i)), ",")
                If UBound(sOrderSmallParts) = 1 Then
                
                    If sOrderSmallParts(0) = "Rep" Then
                        sOrderBy = " cust_salesrepref_fullname "
                    ElseIf sOrderSmallParts(0) = "Importance" Then
                        sOrderBy = " importance_type "
                    ElseIf sOrderSmallParts(0) = "Customer" Then
                        sOrderBy = " cust_fullname "
                    ElseIf sOrderSmallParts(0) = "Balance" Then
                        sOrderBy = " cust_totalbalance_money "
                    ElseIf sOrderSmallParts(0) = "Phone" Then
                        sOrderBy = " cust_phone1 "
                    ElseIf sOrderSmallParts(0) = "Contact" Then
                        sOrderBy = " cust_contact "
                    ElseIf sOrderSmallParts(0) = "Status" Then
                        sOrderBy = " cust_webstatus "
                    End If
                        
                    'If Combo3.Text = "Ascending" Then
                    '    sWhichOrder = " asc "
                    'ElseIf Combo3.Text = "Descending" Then
                    '    sWhichOrder = " desc "
                    'End If
                    If sWhichOrder = "" Then
                        sWhichOrder = "Order By " & sOrderBy & " " & sOrderSmallParts(1)
                    Else
                        sWhichOrder = sWhichOrder & ", " & sOrderBy & " " & sOrderSmallParts(1)
                    End If
                End If
            Next i
        'End If
    Else
        sWhichOrder = ""
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select cust_listid, cust_salesrepref_fullname, cust_fullname, cust_webstatus " & _
            " , cust_totalbalance, cust_totalbalance_money, cust_contact, cust_phone1, cust_jobstatus, importance_name, importance_type, importance_upfront " & _
            " from qbx_cust " & _
            " where cust_jobstatus <> 'awarded' and cust_fullname <> '' " & _
            " and sign(cust_totalbalance) <> '-1' and cust_totalbalance <> '0.00' " & sSQLXtra & " " & sCriteria & " " & sWhichOrder
                                
    Text2.Text = cmdCommand.CommandText
         'Exit Sub
    Set rsGrabCustomers = cmdCommand.Execute
        
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
    Set rsGrabCustomers = Nothing
End Sub

Private Sub prcReportRequest()
    Dim numItems As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sName As String
    Dim sRep
    Dim sBal As Currency
    Dim sContact
    Dim sPhone
    Dim sStatus As String
    Dim sWebStatus As String
    Dim sImportancelvl As String
    Dim sUpfront As String
    Dim txnId As String
    Dim ifound As Integer
    Dim sTBal As Currency
    Dim CustTableWidth As String
    Dim NoteTableWidth As String
    Dim NoteAtBottom As Boolean
    
On Error Resume Next
    
    prcGrabCustomers
    'Exit Sub
    
    If Check2.Value = 0 Then
        NoteAtBottom = False
        CustTableWidth = "70%"
        NoteTableWidth = "30%"
    Else
        NoteAtBottom = True
        CustTableWidth = "100%"
        NoteTableWidth = "100%"
    End If
    
    '''''
    
    If rsGrabCustomers.State = 1 Then
        numItems = rsGrabCustomers.RecordCount
        
        
        sPartOne = "<html><head><title></title>"
        sPartOne = sPartOne & "<style type=""text/css"">"
        sPartOne = sPartOne & ".cust {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 10px;font-weight: bold;}"
        sPartOne = sPartOne & ".inv {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}"
        sPartOne = sPartOne & ".msg {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}"
        sPartOne = sPartOne & ".msgbody {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}"
        sPartOne = sPartOne & ".title {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 12px;font-weight: bold;}"
        sPartOne = sPartOne & "</style></head><body>"
        sPartTwo = "<table border=1 width=100% align=left>"
        
        ifound = 0
        sTBal = 0
        Label4.Caption = 0
        dTxnDate4 = Format(Now, "mm/dd/yyyy")
        dTxnDate3 = Format(Now, "mm/dd/yyyy")
        
        If Not rsGrabCustomers.EOF Then
        
            rsGrabCustomers.MoveFirst
            While Not rsGrabCustomers.EOF
                           
                txnId = Trim(rsGrabCustomers!cust_listid)
                sRep = Trim(rsGrabCustomers!cust_salesrepref_fullname)
                sName = Trim(rsGrabCustomers!cust_fullname)
                sBal = Trim(rsGrabCustomers!cust_totalbalance)
                sContact = Trim(rsGrabCustomers!cust_contact)
                sPhone = Trim(rsGrabCustomers!cust_phone1)
                sStatus = LCase(Trim(rsGrabCustomers!cust_jobstatus))
                sWebStatus = Trim(rsGrabCustomers!cust_webstatus)
                sImportancelvl = Trim(rsGrabCustomers!importance_name) & ""
                If sImportancelvl = "" Then
                    sImportancelvl = "empty"
                End If
                sUpfront = Trim(rsGrabCustomers!importance_upfront) & ""
                If sUpfront = "" Then
                    sUpfront = "empty"
                End If
                'If custChildNode.nodeName = "JobStatus" Then
                'If LCase(custChildNode.Text) = "awarded" Then
                        
                
                    'If sName <> "" And sStatus <> "awarded" And InStr(1, sBal, "-") = 0 And sBal <> "0.00" Then
                    
                        If Option1.Value = False Then
                            If sRep = Trim(Combo1.Text) Then
                                sPartTwo = sPartTwo & "<tr><td><table border=1 width=100% align=left>"
                                sPartTwo = sPartTwo & "<tr>" & _
                                                "<td width=" & CustTableWidth & " height=40 valign=top>" & _
                                                    "<table border=1 width=100%><tr>" & _
                                                        "<td width=170 class=cust>" & sName & "</td>" & _
                                                        "<td width=120 class=cust>" & sContact & "</td>" & _
                                                        "<td width=120 align=center class=cust>" & sPhone & "</b></td>" & _
                                                        "<td width=30 align=center class=cust>" & sRep & "</td>" & _
                                                    "</tr></table>" & _
                                                "</td>" & _
                                                "<td width=" & NoteTableWidth & " rowspan=4 valign=top align=left>" & _
                                                    "<table border=0 width=100%><tr>" & _
                                                        "<td valign=top><table border=1>"
                                
                                'prcGrab_Importance txnId
                                sImportance = "<table border=1 width=100>"
                                sImportance = sImportance & "<tr><td class=msg width=50>Importance:</td><td class=msg width=100>" & sImportancelvl & "</td></tr>"
                                sImportance = sImportance & "<tr><td class=msg width=50>Upfront:</td><td class=msg width=100>" & sUpfront & "</td></tr>"
                                sImportance = sImportance & "<tr><td class=msg width=50>Status:</td><td class=msg width=100>" & sWebStatus & "</td></tr>"
                                sImportance = sImportance & "</table>"
                                'If sImportance = "" Then
                                    'sImportance = "<table border=1 width=100>"
                                    'sImportance = sImportance & "<tr><td class=msg width=50>Type:</td><td class=msg width=100>empty</td></tr>"
                                    'sImportance = sImportance & "<tr><td class=msg width=50>Upfront:</td><td class=msg width=100>empty</td></tr></table>"
                                'End If
                                
                                'grab all messages for this customer that start with a specific date
                                'grab all invoices and return the date for the lowest date
                                prcGrab_Messages txnId, QBFC_Invoices2(txnId)
                                    
                                sPartTwo = sPartTwo & "</table></td></tr></table></td></tr>"
                                                            
                                                            
                                sPartTwo = sPartTwo & sInvStr
                                                        
                                sPartTwo = sPartTwo & "<tr><td align=right width=70%><table border=0 width=70%><tr><td width=100 align=left>&nbsp;</td><td width=100 align=left class=cust>Balance:</td><td width=100 align=center class=cust>$" & funFormatCurr2String(sBal) & "</td></tr></table></td></tr>"
                                sPartTwo = sPartTwo & "<tr><td width=70%>&nbsp;<br></td></tr>"
                                sPartTwo = sPartTwo & "</table></td></tr>"
                                
                                ifound = ifound + 1
                                Label1.Caption = ifound
                                Label1.Refresh
                                sTBal = sTBal + sBal
                                
                                If dTxnDate4 > dTxnDate3 Then
                                    dTxnDate4 = dTxnDate3
                                End If
                            End If
                        Else
                            
                            sPartTwo = sPartTwo & "<tr><td><table border=0 width=100% align=left>"
                            sPartTwo = sPartTwo & "<tr><td width=70%><table border=1 width=100%><tr><td width=170 class=cust>" & sName & "</td><td width=120 class=cust>" & sContact & "</td><td width=120 align=center class=cust>" & sPhone & "</b></td><td width=30 align=center class=cust>" & sRep & "</td></tr></table></td>"
                            sPartTwo = sPartTwo & "<td rowspan=4 valign=top align=left><table border=0 width=100%><tr><td><table border=1>"
                    
                            sImportance = "<table border=1 width=100>"
                            sImportance = sImportance & "<tr><td class=msg width=50>Importance:</td><td class=msg width=100>" & sImportancelvl & "</td></tr>"
                            sImportance = sImportance & "<tr><td class=msg width=50>Upfront:</td><td class=msg width=100>" & sUpfront & "</td></tr>"
                            sImportance = sImportance & "<tr><td class=msg width=50>Status:</td><td class=msg width=100>" & sWebStatus & "</td></tr>"
                            sImportance = sImportance & "</table>"
                            'prcGrab_Importance txnId
                            
                            'If sImportance = "" Then
                            '    sImportance = "<table border=1 width=100>"
                            '    sImportance = sImportance & "<tr><td class=msg width=50>Type:</td><td class=msg width=100>empty</td></tr>"
                            '    sImportance = sImportance & "<tr><td class=msg width=50>Upfront:</td><td class=msg width=100>empty</td></tr></table>"
                            'End If
                            
                            'grab all messages for this customer that start with a specific date
                            'grab all invoices and return the date for the lowest date
                            prcGrab_Messages txnId, QBFC_Invoices2(txnId)
                            
                            sPartTwo = sPartTwo & "</table></td></tr></table></td></tr>"
                                                        
                                                        
                            sPartTwo = sPartTwo & sInvStr
                                                        
                            sPartTwo = sPartTwo & "<tr><td align=right width=70%><table border=0 width=70%><tr><td width=100 align=left>&nbsp;</td><td width=100 align=left class=cust>Balance:</td><td width=100 align=center class=cust>$" & funFormatCurr2String(sBal) & "</td></tr></table></td></tr>"
                            sPartTwo = sPartTwo & "<tr><td width=70%>&nbsp;<br></td></tr>"
                            sPartTwo = sPartTwo & "</table></td></tr>"
                                
                            ifound = ifound + 1
                            Label1.Caption = ifound
                            Label1.Refresh
                            sTBal = sTBal + sBal
                            
                            If dTxnDate4 > dTxnDate3 Then
                                dTxnDate4 = dTxnDate3
                            End If
                        End If
                   ' End If
                rsGrabCustomers.MoveNext
            Wend
        End If
        
        sTitleStr = "<table border=1 width=100% align=center><tr><td><table border=0 width=100% align=center>"
        sTitleStr = sTitleStr & "<tr><td class=title align=left>Today: " & Format(Now, "mm/dd/yyyy") & "</td><td class=title align=left>Created By: " & sUser & "</td></tr>"
        If Option1.Value = False Then
            sTitleStr = sTitleStr & "<tr><td class=title align=left>Report by user: " & Trim(Combo1.Text) & "</td>"
        Else
            sTitleStr = sTitleStr & "<tr><td class=title align=left>Report by: All</td>"
        End If
        
        sTitleStr = sTitleStr & "<td class=title align=left>Dates: " & dTxnDate4 & " - " & Format(Now, "mm/dd/yyyy") & "</td></tr>"
        sTitleStr = sTitleStr & "<tr><td colspan=2 class=title align=right>Record Count:</td><td class=title>" & ifound & "</td></tr></table></td></tr></table><br>"
            
        If Check1.Value = 1 Then
            Label4.Caption = funFormatCurr2String(sTBal)
            'totaling the bottom
            sWebStr = sWebStr & "<tr><td width=70%><table border=0 align=right width=70%><td width=100 align=center class=cust>Total Balance</td><td width=100 align=center class=cust>$" & Trim(Label4.Caption) & "<hr></td><td width=30%>&nbsp;</td></tr></table></td></tr>"
        End If
        
        sPartThree = sPartThree & "</table></body></html>"
        
        sWebStr = sPartOne & sTitleStr & sPartTwo & sWebStr & sPartThree
        prcCreateReportFile "c:\", "report.html", sWebStr
            
        WB1.navigate "c:\1report.html"
        Command2.Enabled = True
            
        prcCleanUpReporting
    
    End If
    
    numItems = 0
    i = 0
    j = 0
    sName = ""
    sRep = ""
    sBal = ""
    sContact = ""
    sPhone = ""
    sStatus = ""
    txnId = ""
    ifound = 0
    sTBal = 0
End Sub


Sub prcCleanUpReporting()
    sPartOne = ""
    sPartTwo = ""
    sPartThree = ""
    sImportance = ""
    sInvStr = ""
    sTitleStr = ""
    sWebStr = ""
End Sub

Sub prcGrab_Messages_new(sTxnID As String, dTxnDate2 As Date)
    Dim rNotes As Sql_Results_Struct
    
    dTxnDate2 = DateAdd("m", -1, dTxnDate2)
    
    rNotes.Query = "  select * from qb_note " & _
            " where note_listid = " & sTxnID & " and note_datestamp > " & dTxnDate2 & _
            " order by note_datestamp desc "
    SQL_Query_auto rNotes.Query, rNotes.Data
            
    If Not rNotes.Data.EOF Then
        rNotes.Data.MoveFirst
        While Not rNotes.Data.EOF
            sPartTwo = sPartTwo & "<tr><td width=30 class=msg>" & Trim(rsReportNotes!note_created_by) & "</td>"
            sPartTwo = sPartTwo & "<td class=msg>" & Trim(rsReportNotes!note_datestamp) & "</td></tr>"
            sPartTwo = sPartTwo & "<tr><td colspan=2 class=msgbody>" & sTmpMsg & vbNewLine & Addmsgs_Output(sTxnID) & "</td></tr>"
            rNotes.Data.MoveNext
        Wend
    End If
    
    SQL_Close_Clear rNotes.Data
End Sub


Function Addmsgs_Output(Ni As String) As String
    Dim Addmsgs As Sql_Results_Struct
    
    Addmsgs.Query = " select * from qb_note_addmsg " & _
                    " where note_index = '" & Ni & "' " & _
                    " order by nadd_datestamp asc "
    SQL_Query_auto Addmsgs.Query, Addmsgs.Data
        
    If Not Addmsgs.Data.EOF Then
        Addmsgs.Data.MoveFirst
        While Not Addmsgs.Data.EOF
            Addmsgs_Output = Addmsgs_Output & vbNewLine & vbNewLine
            Addmsgs_Output = Addmsgs_Output & "****Updaded By " & Trim(Addmsgs.Data!nadd_created_by) & " at " & Trim(Addmsgs.Data!nadd_datestamp) & "****" & vbNewLine
            Addmsgs_Output = Addmsgs_Output & Trim(Addmsgs.Data!nadd_msg) & ""
            Addmsgs.Data.MoveNext
        Wend
    End If
    
    SQL_Close_Clear Addmsgs.Data
End Function

Function Convert_SQL2Text(sMesg As String) As String

    sMesg = Replace(Trim(sMesg), "*##*", "'")
    sMesg = Replace(sMesg, "$++$", """")
                
    Convert_SQL2Text = sMesg
    
End Function

Sub prcGrab_Messages(sTxnID As String, dTxnDate2 As Date)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsReportNotes   As New ADODB.Recordset
    Dim icount          As Integer

On Error GoTo errHandle:

    If Trim(Combo4.Text) <> 0 Then
    
    SQL_ReConnect_old frmMain.cnMC
        If frmMain.cnMC.State <> 1 Then
            Exit Sub
        End If
        
        dTxnDate2 = DateAdd("m", -1, dTxnDate2)
        
        Set cmdCommand.ActiveConnection = frmMain.cnMC
        cmdCommand.CommandType = adCmdStoredProc
        cmdCommand.CommandText = "grab_note_by_date_sp"
        
        'cust txn id
        Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sTxnID) & "")
        cmdCommand.Parameters.Append parParameter
        
        'date
        Set parParameter = cmdCommand.CreateParameter(, adDate, adParamInput, , Trim(dTxnDate2) & "")
        cmdCommand.Parameters.Append parParameter
            
        Set rsReportNotes = cmdCommand.Execute
        
        icount = 0
        If Not rsReportNotes.EOF Then
            Dim sTmpMsg As String
            rsReportNotes.MoveFirst
            rsReportNotes.MoveLast
            While Not rsReportNotes.BOF
                        
                sTmpMsg = Replace(Trim(rsReportNotes!note_msg), "*##*", "'")
                sTmpMsg = Replace(sTmpMsg, "$++$", """")
                sPartTwo = sPartTwo & "<tr><td width=30 class=msg>" & Trim(rsReportNotes!note_created_by) & "</td>"
                sPartTwo = sPartTwo & "<td class=msg>" & Trim(rsReportNotes!note_datestamp) & "</td></tr>"
                sPartTwo = sPartTwo & "<tr><td colspan=2 class=msgbody>" & sTmpMsg & "</td></tr>"
                
                icount = icount + 1
                
                If icount = Trim(Combo4.Text) Then
                    rsReportNotes.MoveFirst
                End If
                rsReportNotes.MovePrevious
                
            Wend
        Else
            'List1.AddItem "No saved settings found, user input required."
        End If
    
    End If
        
    Set rsReportNotes = Nothing
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
    Set rsReportNotes = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcGrab_Importance(sTxnID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsReportImportance   As New ADODB.Recordset

    sImportance = ""
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_importance_sp"
    
    'cust txn id
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sTxnID) & "")
    cmdCommand.Parameters.Append parParameter
        
    Set rsReportImportance = cmdCommand.Execute
    
    If Not rsReportImportance.EOF Then
        sImportance = "<table border=1 width=100>"
        
        rsReportImportance.MoveFirst
        While Not rsReportImportance.EOF
            If Trim(rsReportImportance!importance_name) = "" Then
                sImportance = sImportance & "<tr><td class=msg width=50>Type:</td><td class=msg width=100>empty</td></tr>"
            Else
                sImportance = sImportance & "<tr><td class=msg width=50>Type:</td><td class=msg width=100>" & Trim(rsReportImportance!importance_name) & "</td></tr>"
            End If
            
            If Trim(rsReportImportance!UpFront) = "" Then
                sImportance = sImportance & "<tr><td class=msg width=50>Upfront:</td><td class=msg width=100>empty</td></tr>"
            Else
                sImportance = sImportance & "<tr><td class=msg width=50>Upfront:</td><td class=msg width=100>" & Trim(rsReportImportance!UpFront) & "</td></tr>"
            End If
            rsReportImportance.MoveNext
        Wend
        sImportance = sImportance & "</table>"
        
    Else
        'List1.AddItem "No saved settings found, user input required."
    End If
    
    Set rsReportImportance = Nothing
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
    Set rsReportImportance = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Public Function QBFC_Invoices2(txnId As String) As Date
    Dim Response
    Dim cmdCommand          As New ADODB.Command
    Dim parParameter        As New ADODB.Parameter
    Dim rsGrabInv           As New ADODB.Recordset

On Error GoTo errHandle:

    sInvStr = ""
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select inv_txndate, inv_refnumber, inv_balanceremaining from qbx_inv where inv_customerref_listid = '" & txnId & "' and inv_balanceremaining <> '0.00' order by inv_txndate asc "
        
    Set rsGrabInv = cmdCommand.Execute
    
    If Not rsGrabInv.EOF Then
    
        Dim dTxnDate As Date
        Dim dTemp As Date
        dTxnDate = DateAdd("yyyy", 1, Now)
        
        sInvStr = sInvStr & "<tr><td align=right width=70% valign=top>"
        sInvStr = sInvStr & "<table border=0 width=100%><tr><td align=left width=150 valign=top>" & sImportance & "</td><td align=right>"
        sInvStr = sInvStr & "<table border=1><tr><td width=80 align=center valign=top class=inv>Date</td><td width=80 align=center class=inv>Invoice#</td><td width=100 align=center class=inv>Open Balance</td></tr>"
    
        rsGrabInv.MoveFirst
        While Not rsGrabInv.EOF
        
            sInvStr = sInvStr & "<tr><td align=center class=inv>" & Trim(rsGrabInv!inv_txndate) & "</td><td align=center class=inv>" & Trim(rsGrabInv!inv_refnumber) & "</td><td align=center class=inv>$" & Trim(rsGrabInv!inv_balanceremaining) & "</td></tr>"
            dTemp = Trim(rsGrabInv!inv_txndate)
            
            If dTxnDate > dTemp Then
                dTxnDate = dTemp
                dTxnDate3 = dTemp
            End If
            rsGrabInv.MoveNext
        Wend
        
        sInvStr = sInvStr & "</table></td></tr></table></td></tr>"
    End If
        
    QBFC_Invoices2 = dTxnDate
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function



Sub prcCreateReportFile(sPath As String, sFilename As String, sPage As String)
    Dim f As Integer
    Dim strTemp As String

On Error GoTo errhandler
               
    strTemp = sPath & "1" & sFilename
    
    f = FreeFile
    Open strTemp For Output As #f
    Print #f, sWebStr
    Close #f
    Exit Sub
    
errhandler:
    Close #f
    MsgBox "An error occured by creating the reporting file."
    
End Sub

Sub prcGrabReps()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsGrabReps      As New ADODB.Recordset
    Dim iRepCount       As Integer

    Combo1.Text = ""
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select rep_initial from qbx_reps where rep_isactive = 'True' order by rep_initial asc "
        
    Set rsGrabReps = cmdCommand.Execute
    
    If Not rsGrabReps.EOF Then
        ReDim sRepList(rsGrabReps.RecordCount)
        iRepCount = 0
        rsGrabReps.MoveFirst
        While Not rsGrabReps.EOF
            Combo1.AddItem Trim(rsGrabReps!rep_initial)
            sRepList(iRepCount) = Trim(rsGrabReps!rep_initial)
            iRepCount = iRepCount + 1
            rsGrabReps.MoveNext
        Wend
    End If
    
    Set rsGrabReps = Nothing
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




Private Sub Command1_Click()
    Frame1.Enabled = False
    Label1.Caption = 0
    prcReportRequest
        
    Frame1.Enabled = True
End Sub




Private Sub Command2_Click()
    WB1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub


Private Sub Command3_Click()
    Command3.Visible = False
    Command4.Visible = True
    
    Frame2.Visible = True
    Frame3.Visible = True
    '9015 - 8535
    '480 + 4395
    '4875 + 4140 = 9015
    WB1.Left = 60
    WB1.top = 4140
    'WB1.top = 2100
    WB1.Width = 11055
    WB1.Height = 4875
End Sub

Private Sub Command4_Click()
    Command4.Visible = False
    Command3.Visible = True
    
    Frame2.Visible = False
    Frame3.Visible = False
    
    WB1.Left = 60
    WB1.top = 1080
    WB1.Width = 11055
    WB1.Height = 7935
End Sub

Private Sub Command5_Click()
    Dim sOrder As String
    
    If Trim(Combo3.Text) = "Ascending" Then
        sOrder = "asc"
    Else
        sOrder = "desc"
    End If
    
    If Trim(Text1.Text) = "" Then
        Text1.Text = Trim(Combo2.Text) & "," & sOrder
    Else
        Text1.Text = Trim(Text1.Text) & ";" & Trim(Combo2.Text) & "," & sOrder
    End If
End Sub

Private Sub Command6_Click()
    Text1.Text = ""
End Sub

Private Sub Command7_Click()
    List1.Clear
    List2.Clear
    If List1.ListCount > 0 Then
        Command10.Enabled = True
    Else
        Command10.Enabled = False
    End If
End Sub


Private Sub Command8_Click()
    frmReportingAddNew.Show
End Sub

Private Sub Command9_Click()
    Dim iTmpIndex As Integer
    
    If List2.ListIndex <> -1 Then
        If sFrameTypeE = 1 Then
            iTmpIndex = List2.ListIndex
        Else
            iTmpIndex = List1.ListIndex
        End If
        List1.RemoveItem iTmpIndex
        List2.RemoveItem iTmpIndex
        If List1.ListCount > 0 Then
            Command10.Enabled = True
        Else
            Command10.Enabled = False
        End If
    Else
        MsgBox "Please Select an item to remove."
    End If
End Sub

Private Sub Form_Load()

    prcInit
    prcGrabReps
    'If sUser = "charty" Or sUser = "sleavy" Or sUser = "mfishman" Or sUser = "icedeno" Then
    prcCriteriaPropertiesInit
    Label3.Visible = True
    Label4.Visible = True
    Check1.Visible = True
    List2.Visible = True
    List1.Visible = False
    Label1.Caption = ""
    sFrameTypeE = 1
    'End If
        
    
End Sub

Sub prcCriteriaPropertiesInit()
    Combo5.Clear
        
    If Option1.Value = True Then
        Combo5.Text = "Rep"
        Combo5.AddItem "Rep"
        Combo5.AddItem "Importance"
        Combo5.AddItem "Customer"
        Combo5.AddItem "Balance"
        Combo5.AddItem "Phone"
        Combo5.AddItem "Contact"
    ElseIf Option2.Value = True Then
        Combo5.Text = "Importance"
        Combo5.AddItem "Importance"
        Combo5.AddItem "Customer"
        Combo5.AddItem "Balance"
        Combo5.AddItem "Phone"
        Combo5.AddItem "Contact"
    End If
    
End Sub

Sub prcInit()

    Me.Height = 9885
    Me.Width = 11295
    
    Frame1.Left = 60
    Frame1.top = 0
    Frame1.Width = 11055
    Frame1.Height = 1035
    Option2.Value = True
    
    Frame2.Left = 60
    Frame2.top = 1020
    Frame2.Width = 11055
    Frame2.Height = 1155
    Frame2.Visible = False
    
    Frame3.Left = 60
    Frame3.top = 2160
    Frame3.Width = 11055
    Frame3.Height = 1935
    Frame3.Visible = False
    
    Command3.Visible = True
    Command4.Visible = False
    Command10.Enabled = False

    WB1.Left = 60
    WB1.top = 1080
    WB1.Width = 11055
    WB1.Height = 7935
    WB1.navigate sGHtml_Reporting
    'WB1.navigate "z:\qb\collections\html\reporting.html"
    
    Command2.Enabled = False
    Label3.Visible = False
    Label4.Visible = False
    Check1.Visible = False
    Combo4.AddItem "0"
    Combo4.AddItem "1"
    Combo4.AddItem "2"
    Combo4.AddItem "3"
    Combo4.AddItem "4"
    Combo4.AddItem "5"
    Combo4.AddItem "6"
    Combo4.AddItem "7"
    Combo4.AddItem "8"
    Combo4.AddItem "9"
    Combo4.AddItem "10"
    Combo4.AddItem "All"
End Sub

Sub prcDisableCriteriaProperties()
    Combo5.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
End Sub

Sub prcEnableCriteriaProperties()
    Combo5.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
End Sub


Private Sub Label3_DblClick()
    If Me.Height = 9885 Then
        Me.Height = 11280
    Else
        Me.Height = 9885
    End If
End Sub


Private Sub List1_DblClick()
    List2.Visible = True
    List1.Visible = False
    sFrameTypeE = 1
End Sub


Private Sub List2_DblClick()
    List1.Visible = True
    List2.Visible = False
    sFrameTypeE = 2
End Sub

Private Sub Option1_Click()
    prcCriteriaPropertiesInit
End Sub

Private Sub Option2_Click()
    prcCriteriaPropertiesInit
End Sub


Sub prcOLDcriteria()
    'Dim i As Integer
    'For i = 0 To List1.ListCount - 1
    '        sCriteriaSplit = Split(List1.List(i), " ")
    '
    '        If sCriteriaSplit(2) = "=" Then
    '            sSimilarCriteria(i, 3) = 1
    '        Else
    '            sSimilarCriteria(i, 3) = 0
    '        End If
    '
    '        If i <> 0 Then
    '            sSimilarFound = 0
    '            For j = 0 To UBound(sSimilarCriteria()) - 1
    '                If sSimilarCriteria(j, 2) <> "" And Not IsNull(sSimilarCriteria(j, 2)) Then
    '                    If sSimilarCriteria(j, 2) = 1 Then
    '                        sCriteriaSubSplit = Split(sSimilarCriteria(j, 0), " ")
    '                        If sCriteriaSubSplit(1) = sCriteriaSplit(1) Then
    '                            sSimilarFound = 1
    '                            sSimilarCriteria(i, 0) = List1.List(i)
    '                            sSimilarCriteria(i, 1) = sSimilarCriteria(j, 1)
    '                            sSimilarCriteria(i, 2) = 0
    '                            j = UBound(sSimilarCriteria()) + 1
    '                        End If
    '                    End If
    '                End If
    '            Next j
    '
    '            If sSimilarFound = 0 Then
    '                sSimilarCriteria(i, 0) = List1.List(i)
    '                sSimilarCriteria(i, 1) = i
    '                sSimilarCriteria(i, 2) = 1
    '            End If
    '
    '        Else
    '            sSimilarCriteria(i, 0) = List1.List(i)
    '            sSimilarCriteria(i, 1) = i
    '            sSimilarCriteria(i, 2) = 1
    '        End If
    '
    '    Next i
    '
    '
    '    Dim parts As String
    '    Dim ifound As Integer
    '
    '    For i = 0 To UBound(sSimilarCriteria()) - 1
    '        ifound = 0
    '
    '        If sSimilarCriteria(i, 2) = 1 Then
    '            sCriteria = sCriteria & " and ("
    '            For j = (i + 1) To UBound(sSimilarCriteria())
    '                If sSimilarCriteria(i, 1) = sSimilarCriteria(j, 1) Then
    '                    If ifound = 0 Then
    '                        parts = sSimilarCriteria(i, 0)
    '                    End If
    '                    If sSimilarCriteria(i, 3) = 1 Then
    '                        parts = parts & " or " & sSimilarCriteria(j, 0)
    '                    Else
    '                        parts = parts & " and " & sSimilarCriteria(j, 0)
    '                    End If
    '                    ifound = ifound + 1
    '                End If
    '            Next j
    '
    '            If ifound = 0 Then
    '                sCriteria = sCriteria & sSimilarCriteria(i, 0) & ")"
    '            Else
    '                sCriteria = sCriteria & parts & ")"
    '            End If
    '        End If
    '
    '    Next i
    '
    'End If
   '
    ''Text2.Text = sCriteria
    'Exit Sub
End Sub
