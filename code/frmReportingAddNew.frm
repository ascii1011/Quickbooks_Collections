VERSION 5.00
Begin VB.Form frmReportingAddNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Criteria"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   Icon            =   "frmReportingAddNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9525
   Begin VB.Frame Frame3 
      Height          =   4155
      Left            =   6360
      TabIndex        =   26
      Top             =   60
      Width           =   3075
      Begin VB.CommandButton Command9 
         Caption         =   "Remove"
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   3660
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear"
         Height          =   315
         Left            =   2100
         TabIndex        =   28
         Top             =   3660
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   120
         TabIndex        =   27
         Top             =   540
         Width           =   2835
      End
      Begin VB.ListBox List2 
         Height          =   2790
         Left            =   120
         TabIndex        =   30
         Top             =   540
         Width           =   2835
      End
      Begin VB.Label Label8 
         Caption         =   "CriteriaList:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   60
      TabIndex        =   12
      Top             =   1920
      Width           =   6255
      Begin VB.TextBox Text10 
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   420
         Width           =   6015
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add Statement && Exit"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exit without adding"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   420
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Criteria Statement:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6255
      Begin VB.CommandButton Command13 
         Caption         =   "<"
         Height          =   315
         Left            =   3600
         TabIndex        =   23
         ToolTipText     =   "Lesser then"
         Top             =   780
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   ">"
         Height          =   315
         Left            =   3600
         TabIndex        =   22
         ToolTipText     =   "Greater then"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   1395
      End
      Begin VB.CommandButton Command6 
         Caption         =   "And"
         Height          =   315
         Left            =   4440
         MaskColor       =   &H000000FF&
         TabIndex        =   15
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add ->"
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Top             =   1380
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   4620
         TabIndex        =   9
         Top             =   600
         Width           =   1395
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmReportingAddNew.frx":0442
         Left            =   4620
         List            =   "frmReportingAddNew.frx":0444
         TabIndex        =   8
         Top             =   600
         Width           =   1395
      End
      Begin VB.CommandButton Command5 
         Caption         =   "end"
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Top             =   780
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "start"
         Height          =   315
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "similar"
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   780
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "equal"
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Is"
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         ToolTipText     =   "Whether this statement is of type ""is"" or ""is not"" "
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Criteria Item:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1380
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Property:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Flag to Match:"
         Height          =   195
         Left            =   4620
         TabIndex        =   10
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1380
         Width           =   3975
      End
   End
   Begin VB.Label Label3 
      Height          =   195
      Left            =   3960
      TabIndex        =   7
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   7920
      Width           =   3735
   End
End
Attribute VB_Name = "frmReportingAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFrame1Meat As String
Dim sFrame1Start As String
Dim sFrame1End As String
Dim sFrameTypeE As String

Dim sCumulativeSQL As String
Dim sCumulativeSQLE As String
Dim sRepList() As String
Dim iPropDetailFocus As Integer
Dim iTechFocus As Integer


Sub prcIsNot()
    If Command1.Caption = "Not" Then
        'Option1.Caption = "Equal"
        'Option2.Caption = "Similar to"
        'Option3.Caption = "Start with"
        'Option4.Caption = "End with"
    Else
        'Option1.Caption = "Not Equal"
        'Option2.Caption = "Not Similar to"
        'Option3.Caption = "Not Start with"
        'Option4.Caption = "Not End with"
    End If
End Sub


Private Sub Combo1_Click()
    prcBuildStatements
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    
    prcBuildStatements
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Not" Then
        Command1.Caption = "Is"
    Else
        Command1.Caption = "Not"
    End If
    prcIsNot
    prcBuildStatements
End Sub

Private Sub Command10_Click()
    Unload Me
End Sub

Private Sub Command11_Click()
    frmReporting.List1.AddItem Trim(Text9.Text)
    frmReporting.List2.AddItem Trim(Text10.Text)
    If frmReporting.List1.ListCount > 0 Then
        frmReporting.Command10.Enabled = True
    Else
        frmReporting.Command10.Enabled = False
    End If
    Unload Me
End Sub

Private Sub Command12_Click()
    sFrame1Meat = ">"
    sFrame1Start = ""
    sFrame1End = ""
    sFrameTypeE = 5
    prcBuildStatements
End Sub

Private Sub Command13_Click()
    sFrame1Meat = "<"
    sFrame1Start = ""
    sFrame1End = ""
    sFrameTypeE = 6
    prcBuildStatements
End Sub

Private Sub Command2_Click()
    sFrame1Meat = "="
    sFrame1Start = ""
    sFrame1End = ""
    sFrameTypeE = 1
    prcBuildStatements
End Sub

Private Sub Command3_Click()
    sFrame1Meat = "like"
    sFrame1Start = "%"
    sFrame1End = "%"
    sFrameTypeE = 2
    prcBuildStatements
End Sub

Private Sub Command4_Click()
    sFrame1Meat = "like"
    sFrame1Start = ""
    sFrame1End = "%"
    sFrameTypeE = 3
    prcBuildStatements
End Sub

Private Sub Command5_Click()
    sFrame1Meat = "like"
    sFrame1Start = "%"
    sFrame1End = ""
    sFrameTypeE = 4
    prcBuildStatements
End Sub

Sub prcBuildStatements()
    Dim sProperties As String
    Dim sStatement As String
    Dim sTmp As String
    Dim sFlag As String
    
    'english
    Dim sEName As String
    Dim sEMiddle As String
    Dim sEDetail As String
    Dim sEMath As String
    Dim sEIs    As String
        
    Label5.Caption = ""
    sEName = ""
    sEMiddle = ""
    sEDetail = ""
    
    sTmp = " "
    If Command1.Caption = "Not" Then
        If sFrame1Meat = "=" Then
            sTmp = " !"
        ElseIf sFrame1Meat = "like" Then
            sTmp = " not "
        End If
        sEMiddle = "not"
    Else
        sTmp = " "
        sEMiddle = ""
    End If
    
    
    If iPropDetailFocus = 2 Then
        sFlag = Trim(Text2.Text)
        sEDetail = sFlag
    Else
        sFlag = Trim(Combo1.Text)
        
        sEDetail = sFlag
        If sFlag = "other" Then
            sFlag = ""
            sEDetail = "nothing"
        End If
    End If
    
    sStatement = Trim(Text1.Text)
    If sStatement = "Rep" Then
        sEName = sStatement
        sProperties = "cust_salesrepref_fullname"
    ElseIf sStatement = "Importance" Then
        sEName = sStatement
        sProperties = "importance_name"
    ElseIf sStatement = "Customer" Then
        sEName = sStatement
        sProperties = "cust_fullname"
    ElseIf sStatement = "Balance" Then
        sEName = sStatement
        sProperties = "CONVERT(int, cust_totalbalance_money)"
    ElseIf sStatement = "Phone" Then
        sEName = sStatement
        sProperties = "cust_phone1"
    ElseIf sStatement = "Contact" Then
        sEName = sStatement
        sProperties = "cust_contact"
    ElseIf sStatement = "Status" Then
        sEName = sStatement
        sProperties = "cust_webstatus"
    End If
    
    If sFrameTypeE = 1 Then
        sEMath = "equal to"
        sEIs = "is "
    ElseIf sFrameTypeE = 2 Then
        sEMath = "contains"
        sEIs = ""
    ElseIf sFrameTypeE = 3 Then
        sEMath = "begins with"
        sEIs = ""
    ElseIf sFrameTypeE = 4 Then
        sEMath = "ends with"
        sEIs = ""
    ElseIf sFrameTypeE = 5 Then
        sEMath = "greater then"
        sEIs = "is "
    ElseIf sFrameTypeE = 6 Then
        sEMath = "lesser then"
        sEIs = "is "
    End If
    
    Label5.Caption = sProperties & sTmp & sFrame1Meat & " '" & sFrame1Start & sFlag & sFrame1End & "'"
    Label6.Caption = sEName & " " & sEIs & sEMiddle & " " & sEMath & " " & sEDetail
    'Text10.Text = sEName & " is " & sEMiddle & " " & sEMath & " " & sEDetail
    
    'cumulative
    Dim i As Integer
    Dim sDuplicateFlag As Boolean
    sDuplicateFlag = False
    sCumulativeSQL = "( "
    
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = Trim(Label5.Caption) Then
            sDuplicateFlag = True
        End If
        'Else
            If i = 0 And i < List1.ListCount Then
                sCumulativeSQL = sCumulativeSQL & List1.List(i)
            Else
                If Trim(List1.List(i)) <> "" Then
                    sCumulativeSQL = sCumulativeSQL & " " & LCase(Command6.Caption) & " " & Trim(List1.List(i))
                End If
            End If
        'End If
    Next i
    If List1.ListCount < 1 Then
        sDuplicateFlag = True
    End If
    
    sCumulativeSQLE = "( "
    For i = 0 To List2.ListCount - 1
        If List2.List(i) = Trim(Label6.Caption) Then
            sDuplicateFlag = True
        End If
        'Else
            If i = 0 Then
                sCumulativeSQLE = sCumulativeSQLE & List2.List(i)
            Else
                If Trim(List2.List(i)) <> "" Then
                    sCumulativeSQLE = sCumulativeSQLE & " " & LCase(Command6.Caption) & " " & Trim(List2.List(i))
                End If
            End If
        'End If
    Next i
        
    If sDuplicateFlag = False Then
        Text9.Text = sCumulativeSQL & " " & LCase(Command6.Caption) & " " & Trim(Label5.Caption) & " )"
    Else
        If List1.ListCount > 0 Then
            Text9.Text = sCumulativeSQL & " )"
        Else
            Text9.Text = sCumulativeSQL & " " & Trim(Label5.Caption) & " )"
        End If
    End If
    
    'If List1.ListCount < 1 Then
    '    Text9.Text = sCumulativeSQL & " " & Trim(Label5.Caption) & " )"
    'Else
    '    'If List1.ListCount = 1 Then
    '    If sDuplicateFlag = False Then
    '        Text9.Text = sCumulativeSQL & " " & LCase(Command6.Caption) & " " & Trim(Label5.Caption) & " )"
    '    Else
    '        Text9.Text = sCumulativeSQL & " " & Trim(Label5.Caption) & " )"
    '    End If
    '    'Else
    '        'Text9.Text = sCumulativeSQL & " " & LCase(Command6.Caption) & " " & Trim(Label5.Caption) & " )"
    '    'End If
    'End If
    
    If sDuplicateFlag = False Then
        Text10.Text = sCumulativeSQLE & " " & LCase(Command6.Caption) & " " & Trim(Label6.Caption) & " )"
    Else
        If List1.ListCount > 0 Then
            Text10.Text = sCumulativeSQLE & " )"
        Else
            Text10.Text = sCumulativeSQLE & " " & Trim(Label6.Caption) & " )"
        End If
    End If
    
    'If List2.ListCount < 1 Then
    '    Text10.Text = sCumulativeSQLE & " " & Trim(Label6.Caption) & " )"
    'Else
    '    If sDuplicateFlag = True Then
    '        Text10.Text = sCumulativeSQLE & " " & Trim(Label6.Caption) & " )"
    '    Else
    '        Text10.Text = sCumulativeSQLE & " " & LCase(Command6.Caption) & " " & Trim(Label6.Caption) & " )"
    '    End If
    'End If
    
    Label2.Caption = ""
End Sub



Sub prcCriteriaPropertiesChanged()
    Dim sProperties As String
    
    Combo1.Visible = True
    Text2.Visible = False
    sProperties = Trim(Text1.Text)
    Combo1.Clear
    iPropDetailFocus = 2
    Command12.Visible = False
    Command13.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    
    If sProperties = "Rep" Then
        iPropDetailFocus = 1
        
        prcGrabReps
        
        Dim i As Integer
        
        For i = 0 To UBound(sRepList)
            If i = 0 Then
                Combo1.Text = sRepList(i)
            End If
            Combo1.AddItem sRepList(i)
        Next i
    ElseIf sProperties = "Importance" Then
        iPropDetailFocus = 1
        Combo1.Visible = True
        Text2.Visible = False
                       
        Combo1.Clear
        For i = 0 To UBound(aryGImportLvl) - 1
            If i = 0 Then
                Combo1.Text = aryGImportLvl(i, 2)
            End If
            Combo1.AddItem aryGImportLvl(i, 2)
        Next i
        
        'Combo1.Text = "Collections"
        'Combo1.AddItem "Collections"
        'Combo1.AddItem "High"
        'Combo1.AddItem "Medium"
        'Combo1.AddItem "Low"
        'Combo1.AddItem "Sales"
        'Combo1.AddItem "Discuss"
        'Combo1.AddItem "Priority"
        'Combo1.AddItem "Other"
    ElseIf sProperties = "Customer" Then
        Combo1.Visible = False
        Text2.Visible = True
        Command3.Visible = True
        Command4.Visible = True
        Command5.Visible = True
    ElseIf sProperties = "Balance" Then
        Text2.Text = 0
        Command12.Visible = True
        Command13.Visible = True
        Combo1.Visible = False
        Text2.Visible = True
    ElseIf sProperties = "Phone" Then
        Combo1.Visible = False
        Text2.Visible = True
        Command3.Visible = True
        Command4.Visible = True
        Command5.Visible = True
    ElseIf sProperties = "Contact" Then
        Combo1.Visible = False
        Text2.Visible = True
        Command3.Visible = True
        Command4.Visible = True
        Command5.Visible = True
    End If
End Sub






Private Sub Command6_Click()
    If Command6.Caption = "And" Then
        Command6.Caption = "Or"
    Else
        Command6.Caption = "And"
    End If
    prcBuildStatements
End Sub

Private Sub Command7_Click()
    Dim sDuplicateFlag As Boolean
    Dim i As Integer
    sDuplicateFlag = False
    
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = Trim(Label5.Caption) Then
            sDuplicateFlag = True
            i = List1.ListCount + 1
        End If
    Next i
    
    For i = 0 To List2.ListCount - 1
        If List2.List(i) = Trim(Label6.Caption) Then
            'sDuplicateFlag = True
            i = List2.ListCount + 1
        End If
    Next i
    
    If sDuplicateFlag = False Then
        List1.AddItem Trim(Label5.Caption)
        List2.AddItem Trim(Label6.Caption)
        Label3.Caption = "Added!"
    Else
        Label3.Caption = "Duplicate - Not Added!"
    End If
    
    prcCheck4Count
    prcBuildStatements
End Sub

Private Sub Command8_Click()
    List1.Clear
    List2.Clear
    prcCheck4Count
    prcBuildStatements
End Sub

Private Sub Command9_Click()
    Dim iTmpIndex As Integer
    If sFrameTypeE = 1 Then
        iTmpIndex = List2.ListIndex
    Else
        iTmpIndex = List1.ListIndex
    End If
    List1.RemoveItem iTmpIndex
    List2.RemoveItem iTmpIndex
    prcCheck4Count
    prcBuildStatements
End Sub

Private Sub Form_Load()
    Me.Height = 4785
    Me.Width = 9615
    
    Text1.Text = Trim(frmReporting.Combo5.Text)
    sFrame1Meat = "="
    Command6.Caption = "And"
    iPropDetailFocus = 0
    sFrameTypeE = 1
    
    Text10.Visible = True
    Text10.Enabled = True
    Label6.Visible = True
    List2.Visible = True
    Text9.Visible = False
    Label5.Visible = False
    List1.Visible = False
    iTechFocus = 1
        
    prcCriteriaPropertiesChanged
    prcBuildStatements
    prcCheck4Count
End Sub

Sub prcCheck4Count()
    If List1.ListCount = 0 Or List1.ListCount > 1 Then
        Command6.Enabled = False
        Command6.FontBold = False
        Command6.FontSize = 8
    Else
        Command6.Enabled = True
        Command6.FontBold = True
        Command6.FontSize = 12
    End If
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

Private Sub Label1_Click()
    If iTechFocus = 1 Then
        Text9.Visible = True
        Text9.Enabled = True
        List1.Visible = True
        Label5.Visible = True
        Text10.Visible = False
        Text10.Enabled = False
        List2.Visible = False
        Label6.Visible = False
        iTechFocus = 2
    ElseIf iTechFocus = 2 Then
        Text10.Visible = True
        Text10.Enabled = True
        List2.Visible = True
        Label6.Visible = True
        Text9.Visible = False
        Text9.Enabled = False
        List1.Visible = False
        Label5.Visible = False
        iTechFocus = 1
    End If
End Sub


Private Sub Text2_Change()
    prcBuildStatements
End Sub

