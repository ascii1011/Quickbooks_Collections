VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmImportanceSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importance Settings"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8205
   Begin VB.Frame Frame1 
      Caption         =   "Importance"
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2775
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmImportanceSettings.frx":0000
         Left            =   120
         List            =   "frmImportanceSettings.frx":0002
         TabIndex        =   2
         Text            =   "Create"
         Top             =   360
         Width           =   2535
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   3555
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6271
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ID"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Name"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=423"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=344"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2858"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2778"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   2
         DefColWidth     =   0
         EditDropDown    =   0   'False
         HeadLines       =   1
         FootLines       =   1
         TabAction       =   1
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   12307669
         ScrollTrack     =   -1  'True
         RowDividerColor =   12307669
         RowSubDividerColor=   12307669
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=44,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=27,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=28,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=43,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
         _StyleDefs(65)  =   "Named:id=25:payment"
         _StyleDefs(66)  =   ":id=25,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(67)  =   "Named:id=26:Balance"
         _StyleDefs(68)  =   ":id=26,.parent=25,.fgcolor=&H0&,.borderColor=&H80000007&,.bold=-1,.fontsize=825"
         _StyleDefs(69)  =   ":id=26,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(70)  =   ":id=26,.fontname=MS Sans Serif"
      End
      Begin VB.Label lblCount 
         Height          =   195
         Left            =   1620
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Count:"
         Height          =   195
         Left            =   960
         TabIndex        =   24
         Top             =   4440
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Create Importance"
      Height          =   4395
      Left            =   2880
      TabIndex        =   3
      Top             =   60
      Width           =   5235
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Create"
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   1980
         Width           =   915
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label18 
         Caption         =   "New Importance Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Insert this new importance before "
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1020
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Modify Importance"
      Height          =   4395
      Left            =   2880
      TabIndex        =   9
      Top             =   60
      Width           =   5235
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "Update"
         Height          =   315
         Left            =   2820
         TabIndex        =   13
         Top             =   960
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Importance Active"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   780
         TabIndex        =   15
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Move this importance up or down in the list."
         Height          =   195
         Left            =   780
         TabIndex        =   14
         Top             =   2700
         Width           =   3075
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frmImportanceSettings.frx":0004
         ToolTipText     =   "Move down"
         Top             =   3420
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmImportanceSettings.frx":0446
         ToolTipText     =   "Move up"
         Top             =   2700
         Width           =   480
      End
      Begin VB.Label Label16 
         Caption         =   "Importance Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Remove Importance"
      Height          =   4395
      Left            =   2880
      TabIndex        =   17
      Top             =   60
      Width           =   5235
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Process"
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   3120
         Width           =   915
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   660
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Step 3.  Process Request. (Might take a few minutes.)"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2820
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Step 2.  Choose an Importance that will take over the Importance Level you are removing."
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Step 1.  Choose Importance to remove"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   420
         Width           =   2835
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Count:"
      Height          =   195
      Left            =   2940
      TabIndex        =   27
      Top             =   4560
      Width           =   555
   End
   Begin VB.Label Label8 
      Height          =   195
      Left            =   3480
      TabIndex        =   26
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   $"frmImportanceSettings.frx":0888
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   540
      TabIndex        =   16
      Top             =   5040
      Width           =   6735
   End
End
Attribute VB_Name = "frmImportanceSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsImportance        As New ADODB.Recordset

Dim aryImportance()     As String

Dim iImportSelected As Integer
Dim iHighestID      As Integer

Private Sub Combo1_Change()
    prcChoosing
End Sub

Private Sub Combo1_Click()
    prcChoosing
End Sub

Sub prcDefaultDisable()
    Combo1.Enabled = False
    Combo7.Enabled = False
    Command1.Enabled = False
    Command4.Enabled = False
End Sub

Sub prcChoosing()

    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    
    If Trim(Combo1.Text) = "Create" Then
        Frame2.Visible = True
    ElseIf Trim(Combo1.Text) = "Modify" Then
        Frame3.Visible = True
    ElseIf Trim(Combo1.Text) = "Remove" Then
        Frame4.Visible = True
    Else
        Combo1.Text = "Create"
        Frame2.Visible = True
    End If
    
End Sub

Private Sub Command4_Click()
    If funNewImportanceIsUnique = True Then
        funCreateNewImportance
    Else
        Label8.Caption = "That Importance Name is Not Unique. Please Change."
    End If
End Sub

Function funCreateNewImportance() As Boolean
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Function
    End If
    
    strXtraSQL = " insert into qbx_importance_levels " & _
                " ( import_name, import_id ) " & _
                " values " & _
                " ( '" & Trim(Text1.Text) & "', '" & iHighestID & "' ) "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = strXtraSQL
        
    cmdCommand.Execute
    
    Label8.Caption = "That Importance Name is Unique, Continuing."
    Set cmdCommand = Nothing
    Set parParameter = Nothing
    Exit Function
    
errHandle:
    Label8.Caption = "An error occured while adding Importance entry."
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
End Function

Function funNewImportanceIsUnique() As Boolean
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsDupSearch     As New ADODB.Recordset
    Dim bFound          As Boolean

On Error GoTo errHandle:

    funNewImportanceIsUnique = False
    iHighestID = 0
    
    If rsDupSearch.State = 0 Then Set rsDupSearch = Nothing
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then Exit Function
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_importance_levels  "
        
    Set rsDupSearch = cmdCommand.Execute
    
    If Not rsDupSearch.EOF Then
    
        rsDupSearch.MoveFirst
        While Not rsDupSearch.EOF
        
            If iHighestID < Trim(rsDupSearch!import_id) And Trim(rsDupSearch!import_id) <> 99 Then
                iHighestID = Trim(rsDupSearch!import_id)
            End If
            
            If LCase(Trim(rsDupSearch!import_name)) = LCase(Trim(Text1.Text)) And funNewImportanceIsUnique = False Then
                funNewImportanceIsUnique = True
            End If
            
            rsDupSearch.MoveNext
            
        Wend
        
        'this is for the next import_id number in the table, for creating a new importance
        iHighestID = iHighestID + 1
        
    End If
    
    Set cmdCommand = Nothing
    Set parParameter = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
End Function

Sub prcProcessCreateNewImportance()
    'promote all levels after the new level
    prcPromoteAllLevelsAfterTheNewLevel
End Sub

Sub prcPromoteAllLevelsAfterTheNewLevel()
    Dim importCount As Integer, i As Integer, promotionStart As Integer
    
    promotionStart = 0
    importCount = UBound(aryImportance)
    
    'find place to start promotion
    For i = 0 To importCount - 2
        If aryImportance(i, 1) = Combo7.Text Then
            promotionStart = aryImportance(i, 0)
            i = importCount - 2
        End If
    Next i
    
    For i = importCount - 2 To promotionStart
        prcUpdateImportanceLevelPromotion i
    Next i
    
    
End Sub

Sub prcUpdateImportanceLevelPromotion(Id As Integer)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String
    Dim iPromoteTo      As Integer

On Error GoTo errHandle:

    iPromoteTo = Id + 1
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Sub
    End If
    
    strXtraSQL = " update qbx_importance_levels " & _
                " set import_id = '" & iPromoteTo & "' " & _
                " where import_id = '" & Id & "' and import_name = '" & aryImportance(Id, 1) & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = strXtraSQL
        
    cmdCommand.Execute
    
    Set cmdCommand = Nothing
    Set parParameter = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub


Sub prcUpdateCustImportanceLevelBecauseOfPromotion(Id As Integer, iPromoteTo As Integer)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Sub
    End If
    
    strXtraSQL = " update qbx_importance_levels " & _
                " set import_id = '" & iPromoteTo & "' " & _
                " where import_id = '" & Id & "' and import_name = '" & aryImportance(Id, 1) & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = strXtraSQL
        
    cmdCommand.Execute
    
    Set cmdCommand = Nothing
    Set parParameter = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Me.Width = 8295
    Me.Height = 5445
    sfrmImportanceSettings = 1
    prcInitVals
    
    prcChoosing
    prcGrabAllImportanceLevels
End Sub

Sub prcInitVals()
    
    iImportSelected = 0

    Combo1.Text = "Create"
    'Combo1.AddItem "Create"
    'Combo1.AddItem "Modify"
    'Combo1.AddItem "Remove"
    
    Frame3.Enabled = False
    Frame4.Enabled = False
End Sub

Sub prcGrabAllImportanceLevels()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String
    Dim i               As Integer

On Error GoTo errHandle:

    If rsImportance.State = 1 Then
        Set rsImportance = Nothing
    End If
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Sub
    End If
    
    strXtraSQL = " select * from qbx_importance_levels order by import_id asc "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = strXtraSQL
        
    Set rsImportance = cmdCommand.Execute
    
    lblCount.Caption = rsImportance.RecordCount
    If Not rsImportance.EOF Then
        ReDim aryImportance(rsImportance.RecordCount, 3)
        i = 0
        rsImportance.MoveFirst
        While Not rsImportance.EOF
            aryImportance(i, 0) = Trim(rsImportance!import_id) & ""
            aryImportance(i, 1) = Trim(rsImportance!import_name) & ""
            aryImportance(i, 2) = Trim(rsImportance!import_active) & ""
            
            Combo6.AddItem aryImportance(i, 1)
            
            i = i + 1
            rsImportance.MoveNext
        Wend
    End If
    
    TDBGrid1.ReBind
    
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Note Record Opening Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub





Private Sub Form_Unload(Cancel As Integer)
    sfrmImportanceSettings = 0
End Sub

Private Sub TDBGrid1_DblClick()
    iImportSelected = Trim(TDBGrid1.Columns(0).Value)
End Sub

Private Sub TDBGrid1_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    Dim ColIndex As Integer, col As Integer
    Dim RowsFetched As Integer, Row As Long
    Dim StartRow As Variant
    Dim Response As Long
    Dim cols As Long
    Dim Rows As Long
    Dim Pos As Long
    Dim sOutPutMsg As String

On Error GoTo NoRead
    
    'If bCancelRead Then Exit Sub
    cols = RowBuf.ColumnCount - 1
    Rows = RowBuf.RowCount - 1
    RowsFetched = 0
    
    If IsNull(StartLocation) Then
        If offset < 0 Then
            rsImportance.MoveLast
            rsImportance.MoveNext
        Else
            rsImportance.MoveFirst
            rsImportance.MovePrevious
        End If
        rsImportance.Move offset
    Else
        rsImportance.Move offset, StartLocation
    End If
        
    StartRow = rsImportance.Bookmark
    Pos = rsImportance.AbsolutePosition
    
    For Row = 0 To Rows
        If rsImportance.BOF Or rsImportance.EOF Then Exit For
        For col = 0 To cols
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsImportance!import_id) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsImportance!import_name) & ""
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsImportance.Bookmark
        RowsFetched = RowsFetched + 1
        rsImportance.MoveNext
    Next Row
    RowBuf.RowCount = RowsFetched
    If Pos >= 0 Then ApproximatePosition = Pos
    
    
    Exit Sub

NoRead:
    Select Case (Err.Number)
        Case (3704):
            Exit Sub
        Case (3021):
            RowBuf.RowCount = 0
            'exit sub
            Exit Sub
        Case (-2147217906):
            RowBuf.RowCount = 0
        Case Else:
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & " Try again?" & vbNewLine & "Cancel aborts read", vbExclamation + vbYesNoCancel, "Dial Record DETAIL ERROR")
            Select Case (Response)
                Case (vbYes): Resume
                Case (vbNo): Exit Sub
            End Select
    End Select
End Sub

Private Sub Text1_Change()
    If Trim(Text1.Text) <> "" Then
        Combo7.Enabled = True
        Command4.Enabled = True
    Else
        Combo7.Enabled = False
        Command4.Enabled = False
    End If
End Sub
