VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmRemarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Notes"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "frmRemarks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8685
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   8535
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   15
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   11
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         Height          =   1755
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   420
         Width           =   4935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "New"
         Height          =   375
         Left            =   7620
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7620
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   7620
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   7620
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Message:"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Mod. Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Mod. By:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Created Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Created By:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label22 
         Caption         =   "ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid2 
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   2340
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3096
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID"
      Columns(0).DataField=   ""
      Columns(0).DataWidth=   100
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Message"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Created by"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Creation Date"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Modified By"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Modification Date"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12307669
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=10081"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=10001"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1508"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1429"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2302"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1640"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1561"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2408"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2328"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=30,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=27,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=28,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=29,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=44,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=31,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=32,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=43,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=48,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=45,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=46,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=47,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=52,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=49,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=50,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=51,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
      _StyleDefs(81)  =   "Named:id=25:payment"
      _StyleDefs(82)  =   ":id=25,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(83)  =   "Named:id=26:Balance"
      _StyleDefs(84)  =   ":id=26,.parent=25,.fgcolor=&H0&,.borderColor=&H80000007&,.bold=-1,.fontsize=825"
      _StyleDefs(85)  =   ":id=26,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(86)  =   ":id=26,.fontname=MS Sans Serif"
   End
   Begin VB.Label Label7 
      Height          =   195
      Left            =   7440
      TabIndex        =   19
      Top             =   4140
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Count:"
      Height          =   195
      Left            =   6840
      TabIndex        =   18
      Top             =   4140
      Width           =   495
   End
End
Attribute VB_Name = "frmRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsRemarks As New ADODB.Recordset
Dim bSortOrd As Boolean


Private Sub Command1_Click()
    If Trim(Text13.Text) = "" Then
        prcCreateRemarks
    Else
        prcUpdateRemarks Trim(TDBGrid2.Columns(0).Value)
        
    End If
    prcGrabRemarks
End Sub

Private Sub Command2_Click()
    prcDeleteRemarks Trim(TDBGrid2.Columns(0).Value)
    prcGrabRemarks
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    prcClearRemark
End Sub

Sub prcClearRemark()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text8.Text = ""
    Text13.Text = ""
End Sub

Private Sub Form_Load()
    Me.Width = 8775
    Me.Height = 4800
    sFrmRemarks = 1
    prcGrabRemarks
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmRemarks = 0
End Sub

Sub prcGrabRemarks()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    If rsRemarks.State = 1 Then
        Set rsRemarks = Nothing
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_remarks where remark_delete = '0' order by remark_created_date desc "
            
    Set rsRemarks = cmdCommand.Execute
    
    Label7.Caption = rsRemarks.RecordCount
    If rsRemarks.RecordCount > 0 Then
        TDBGrid2.ApproxCount = rsRemarks.RecordCount
    Else
        Set rsRemarks = Nothing
    End If
    TDBGrid2.ReBind
    
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
    Set rsRemarks = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcUpdateRemarks(sID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qbx_remarks " & _
                            " set remark_msg = '" & Trim(Text8.Text) & "' " & _
                            " , remark_modified_by = '" & sUser & "' " & _
                            " , remark_modified_date = '" & Now & "' " & _
                            " where remark_index = '" & Trim(sID) & "' "
        
    cmdCommand.Execute
    
    prcGrabARemark sID
    
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


Sub prcCreateRemarks()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " insert into qbx_remarks " & _
                            " (remark_msg, remark_created_by, remark_created_date, remark_delete) " & _
                            " values " & _
                            " ('" & Trim(Text8.Text) & "', '" & sUser & "', '" & Now & "', '0') "
        
    cmdCommand.Execute
    
    prcGrabCreatedRemark
    
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


Sub prcDeleteRemarks(sID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qbx_remarks " & _
                            " set remark_delete = '1' " & _
                            " , remark_modified_by = '" & sUser & "' " & _
                            " , remark_modified_date = '" & Now & "' " & _
                            " where remark_index = '" & Trim(sID) & "' "
        
    cmdCommand.Execute
    
    prcClearRemark
    
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


Sub prcGrabARemark(sID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsSingleRemark  As New ADODB.Recordset

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    If rsSingleRemark.State = 1 Then
        Set rsSingleRemark = Nothing
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_remarks where remark_index = '" & sID & "' "
            
    Set rsSingleRemark = cmdCommand.Execute
    
    If Not rsSingleRemark.EOF Then
        If rsSingleRemark.RecordCount = 1 Then
            rsSingleRemark.MoveFirst
            Text13.Text = Trim(rsSingleRemark!remark_index) & ""
            Text1.Text = Trim(rsSingleRemark!remark_created_by) & ""
            Text2.Text = Trim(rsSingleRemark!remark_created_date) & ""
            Text3.Text = Trim(rsSingleRemark!remark_modified_by) & ""
            Text4.Text = Trim(rsSingleRemark!remark_modified_date) & ""
            Text8.Text = Trim(rsSingleRemark!remark_msg) & ""
        End If
    End If
    
    Set rsSingleRemark = Nothing
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
    Set rsSingleRemark = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Sub prcGrabCreatedRemark()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsSingleRemark  As New ADODB.Recordset

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    If rsSingleRemark.State = 1 Then
        Set rsSingleRemark = Nothing
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select top 1 * from qbx_remarks where remark_created_by = '" & sUser & "' order by remark_created_date desc "
            
    Set rsSingleRemark = cmdCommand.Execute
    
    If Not rsSingleRemark.EOF Then
        If rsSingleRemark.RecordCount = 1 Then
            rsSingleRemark.MoveFirst
            Text13.Text = Trim(rsSingleRemark!remark_index) & ""
            Text1.Text = Trim(rsSingleRemark!remark_created_by) & ""
            Text2.Text = Trim(rsSingleRemark!remark_created_date) & ""
            Text3.Text = Trim(rsSingleRemark!remark_modified_by) & ""
            Text4.Text = Trim(rsSingleRemark!remark_modified_date) & ""
            Text8.Text = Trim(rsSingleRemark!remark_msg) & ""
        End If
    End If
    
    Set rsSingleRemark = Nothing
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
    Set rsSingleRemark = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Private Sub TDBGrid2_DblClick()
    prcClearRemark
    prcGrabARemark Trim(TDBGrid2.Columns(0).Value)
End Sub

Private Sub Text13_Change()
    If Trim(Text13.Text) <> "" Then
        Command1.Caption = "Modify"
    Else
        Command1.Caption = "Save"
    End If
End Sub

Private Sub TDBGrid2_Error(ByVal DataError As Integer, Response As Integer)
On Error Resume Next

    Response = 0
End Sub

Private Sub TDBGrid2_HeadClick(ByVal ColIndex As Integer)

On Error Resume Next

    MousePointer = vbHourglass
    If bSortOrd Then
        'prcGrabCustomerInfo rsRemarks.Fields(ColIndex).Name & " ASC "
        rsRemarks.Sort = "[" & rsRemarks.Fields(ColIndex).name & "]  ASC"
        bSortOrd = False
    Else
        'prcGrabCustomerInfo rsRemarks.Fields(ColIndex).Name & " DESC "
        rsRemarks.Sort = "[" & rsRemarks.Fields(ColIndex).name & "]  DESC"
        bSortOrd = True
    End If
    TDBGrid2.col = ColIndex
    rsRemarks.MoveFirst
    TDBGrid2.ReBind
    MousePointer = vbDefault

End Sub

Private Sub TDBGrid2_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
Dim ColIndex As Integer, col As Integer
    Dim RowsFetched As Integer, Row As Long
    Dim StartRow As Variant
    Dim Response As Long
    Dim cols As Long
    Dim Rows As Long
    Dim Pos As Long
    Dim strlen As Integer

On Error GoTo NoRead
    
    'If bCancelRead Then Exit Sub
    cols = RowBuf.ColumnCount - 1
    Rows = RowBuf.RowCount - 1
    RowsFetched = 0
    
    If IsNull(StartLocation) Then
        If offset < 0 Then
            rsRemarks.MoveLast
            rsRemarks.MoveNext
        Else
            rsRemarks.MoveFirst
            rsRemarks.MovePrevious
        End If
        rsRemarks.Move offset
    Else
        rsRemarks.Move offset, StartLocation
    End If
        
    StartRow = rsRemarks.Bookmark
    Pos = rsRemarks.AbsolutePosition
    
    For Row = 0 To Rows
        If rsRemarks.BOF Or rsRemarks.EOF Then Exit For
        For col = 0 To cols
            'strlen = Len(rsRemarks!Status)
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsRemarks!remark_index) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsRemarks!remark_msg) & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsRemarks!remark_created_by) & ""
                Case (3):   RowBuf.Value(Row, 3) = Trim(rsRemarks!remark_created_date) & ""
                Case (4):   RowBuf.Value(Row, 4) = Trim(rsRemarks!remark_modified_by) & ""
                Case (5):   RowBuf.Value(Row, 5) = Trim(rsRemarks!remark_modified_date) & ""
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsRemarks.Bookmark
        RowsFetched = RowsFetched + 1
        rsRemarks.MoveNext
    Next Row
    RowBuf.RowCount = RowsFetched
    If Pos >= 0 Then ApproximatePosition = Pos
    Exit Sub

NoRead:
    'MsgBox err.Number & vbNewLine & err.Description
    Select Case (Err.Number)
        Case (3704):
            Exit Sub
        Case (3021):
            RowBuf.RowCount = 0
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
