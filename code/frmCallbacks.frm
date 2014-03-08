VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmCallbacks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Callbacks"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "frmCallbacks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   11280
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.OptionButton Option3 
         Caption         =   "Marked Items for specified Date"
         Height          =   195
         Left            =   1860
         TabIndex        =   17
         Top             =   240
         Width           =   2595
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   4020
         Width           =   195
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   4500
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   3960
         Width           =   2115
      End
      Begin VB.OptionButton Option2 
         Caption         =   "View All Notes"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Callbacks"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Go to"
         Height          =   375
         Left            =   10080
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   8940
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7800
         Picture         =   "frmCallbacks.frx":030A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         ToolTipText     =   "Today's callbacks"
         Top             =   540
         Width           =   270
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   540
         Width           =   2115
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   2955
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5212
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ID"
         Columns(0).DataField=   ""
         Columns(0).DataWidth=   1
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Time"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Customer"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Message"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Created"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "By"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   4
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Today"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "idx"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=423"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=344"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2170"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2090"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4260"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4180"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=6800"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=6720"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1693"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1614"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1508"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1429"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1111"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1032"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=51,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=48,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=49,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=50,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=59,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=52,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=53,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=54,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=67,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=60,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=61,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=62,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
         _StyleDefs(89)  =   "Named:id=47:Open"
         _StyleDefs(90)  =   ":id=47,.parent=42,.fgcolor=&HFF&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(91)  =   ":id=47,.strikethrough=0,.charset=0"
         _StyleDefs(92)  =   ":id=47,.fontname=MS Sans Serif"
         _StyleDefs(93)  =   "Named:id=68:CA green"
         _StyleDefs(94)  =   ":id=68,.parent=47,.fgcolor=&H8000&"
      End
      Begin VB.Label Label3 
         Caption         =   "Mark for today:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Count:"
         Height          =   195
         Left            =   9600
         TabIndex        =   8
         Top             =   3900
         Width           =   495
      End
      Begin VB.Label Label12 
         Height          =   195
         Left            =   10200
         TabIndex        =   7
         Top             =   3900
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Date:"
         Height          =   195
         Left            =   4980
         TabIndex        =   4
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Showing Callbacks for:"
         Height          =   195
         Left            =   4980
         TabIndex        =   3
         Top             =   240
         Width           =   1635
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FG4 
      Height          =   855
      Left            =   60
      TabIndex        =   9
      Top             =   5280
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   6
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmCallbacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public rsNoteCallBacks          As New ADODB.Recordset
Public tempDate As String

Dim bSortOrd As Boolean
Dim bCheckAvailable As Boolean

Private Sub Check1_Click()
    If bCheckAvailable = True Then
        prcUpdateCallbackToday
    End If
End Sub


Sub prcUpdateCallbackToday()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim lBookmark As Long

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Sub
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " update qb_note " & _
                    " set note_callback_today = '" & Check1.Value & "', " & _
                    " note_callback_today_date = '" & Date & "' " & _
                    " where note_index = '" & Trim(Text3.Text) & "' "
        
    cmdCommand.Execute
        
        
    If Option1.Value = True Then
        prcGrabNoteCallbacksByDate 0
    ElseIf Option2.Value = True Then
        prcGrabNoteCallbacksByDate 1
    ElseIf Option3.Value = True Then
        prcGrabNoteCallbacksByDate 2
    End If
    
    If TDBGrid1.Bookmark = 2 Then
        TDBGrid1.Bookmark = TDBGrid1.Bookmark - 1
    End If
    If TDBGrid1.Bookmark = 3 Then
        TDBGrid1.Bookmark = TDBGrid1.Bookmark - 2
    End If
    If TDBGrid1.Bookmark > 3 Then
        TDBGrid1.Bookmark = TDBGrid1.Bookmark - 3
    End If
    
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

Private Sub Command1_Click()
    If SQL_ReConnect_old(frmMain.cnMC) = False Then
        frmMain.StatusBar1.Panels.Item(6).Text = "Not Connected."
        Exit Sub
    End If
    frmMain.StatusBar1.Panels.Item(6).Text = "Connected."
    Frame1.Enabled = False
    Command2.Enabled = False
    
    Pause 0.08
    
    tempDate = Trim(Text1.Text)
    
    If IsDate(tempDate) Or tempDate = "" Then
        tempDate = CDate(tempDate)
        If Option1.Value = True Then
            prcGrabNoteCallbacksByDate 0
        ElseIf Option2.Value = True Then
            prcGrabNoteCallbacksByDate 1
        ElseIf Option3.Value = True Then
            prcGrabNoteCallbacksByDate 2
        End If
    End If
        
    Frame1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmCallbacks = 0
End Sub

Private Sub Command2_Click()
    prcGrabInfo
End Sub

Sub prcGrabInfo()

On Error Resume Next

    ListID = Trim(TDBGrid1.Columns(0).Value)
    'frmInvoiceQry.prcFind Trim(TDBGrid1.Columns(0).value), 0
    frmInvoiceQry.prcCallCustomerDtls
    
    frmInvoiceQry.SSTab1.Tab = 0
    frmInvoiceQry.Command11.Enabled = True
    frmInvoiceQry.Frame2.Enabled = True
    If Trim(Text1.Text) <> "" Then
    
        'msg
        If Trim(TDBGrid1.Columns(3).Value) <> "" Then
            frmInvoiceQry.Text6.Text = Trim(TDBGrid1.Columns(3).Value) & vbNewLine & vbNewLine
        Else
            frmInvoiceQry.Text6.Text = "<Empty>" & vbNewLine & vbNewLine
        End If
        
        frmInvoiceQry.Text6.Text = frmInvoiceQry.Text6.Text & "By: " & Trim(TDBGrid1.Columns(5).Value)
        
        'created datestamp
        frmInvoiceQry.Text5.Text = Trim(TDBGrid1.Columns(4).Value)
        
        'callback date
        frmInvoiceQry.Text11.Text = (Text1.Text)
        
        'callback time
        frmInvoiceQry.Text12.Text = Trim(TDBGrid1.Columns(1).Value)
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Width = 11370
    Me.Height = 4980
    sFrmCallbacks = 1
    bCheckAvailable = False
    TDBGrid1.FetchRowStyle = True
    Command2.Enabled = False
    Text1.Text = Date
    tempDate = Trim(Text1.Text)
    Option1.Value = True
    
End Sub


Sub prcGrabNoteCallbacksByDate(sType As Integer)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String

On Error GoTo errHandle:
    If rsNoteCallBacks.State = 1 Then
        Set rsNoteCallBacks = Nothing
    End If
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Sub
    End If
    
    
    'If sProfileAttrDtlsAry(1, 6) = 0 Then
    '    strXtraSQL = " and note_company_status != 'awarded' "
    'End If
    'If sProfileAttrDtlsAry(1, 10) = 0 Then
    '    strXtraSQL = strXtraSQL & " and note_company_amount != '' "
    'End If
    'If sProfileAttrDtlsAry(1, 11) = 0 Then
    '    strXtraSQL = strXtraSQL & " and note_company_status != '0' "
    'End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    If sType = 0 Then
        cmdCommand.CommandText = " select * from qb_note " & _
                    " where note_callback_date = '" & tempDate & "' " & strXtraSQL & _
                    " order by note_callback_time asc "
    ElseIf sType = 1 Then
        cmdCommand.CommandText = " select * from qb_note " & _
                    " where note_listid <> '' " & strXtraSQL & _
                    " order by note_datestamp asc "
    ElseIf sType = 2 Then
        cmdCommand.CommandText = " select * from qb_note " & _
                    " where note_callback_today_date = '" & tempDate & "' and note_callback_today = '1' " & strXtraSQL & _
                    " order by note_callback_time asc "
    End If
    'MsgBox cmdCommand.CommandText
        
    Set rsNoteCallBacks = cmdCommand.Execute
    
    Label12.Caption = rsNoteCallBacks.RecordCount
    If Not rsNoteCallBacks.EOF Then
        
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

Private Sub Option1_Click()
    Label1.Visible = True
    Label2.Visible = True
    Text1.Visible = True
    Picture1.Visible = True
    Text2.Text = ""
    Text3.Text = ""
    Check1.Value = 0
    TDBGrid1.Bookmark = 1
    prcGrabNoteCallbacksByDate 0
End Sub

Private Sub Option2_Click()
    Label1.Visible = False
    Label2.Visible = False
    Text1.Visible = False
    Picture1.Visible = False
    Text2.Text = ""
    Text3.Text = ""
    Check1.Value = 0
    TDBGrid1.Bookmark = 1
    prcGrabNoteCallbacksByDate 1
End Sub

Private Sub Option3_Click()
    Label1.Visible = True
    Label2.Visible = True
    Text1.Visible = True
    Picture1.Visible = True
    Text2.Text = ""
    Text3.Text = ""
    Check1.Value = 0
    TDBGrid1.Bookmark = 1
    prcGrabNoteCallbacksByDate 2
End Sub

Private Sub Picture1_Click()
    
    iCalendarRequest = 1
    frmCalendar.Show
End Sub






Private Sub TDBGrid1_Click()
    TDBGrid1.PostMsg 1
    'bCheckAvailable = False
    'Debug.Print rsNoteCallBacks.Bookmark
    'Text2.Text = TDBGrid1.Columns(0).Value
    'Text3.Text = TDBGrid1.Columns(2).Value
    'Check1.Value = TDBGrid1.Columns(6).Value
    'bCheckAvailable = True
End Sub

Private Sub TDBGrid1_DblClick()
    bCheckAvailable = False
    prcGrabInfo
    bCheckAvailable = True
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
On Error Resume Next

    Response = 0
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)

On Error GoTo errHandle
                                   
    rsNoteCallBacks.Bookmark = Bookmark
    If rsNoteCallBacks!note_callback_today = "1" And Trim(rsNoteCallBacks!note_callback_today_date) = Trim(Text1.Text) Then
    
        If LCase(rsNoteCallBacks!note_state) = "ca" Then
            RowStyle = TDBGrid1.Styles(11)
        Else
            RowStyle = TDBGrid1.Styles(10)
        End If
    End If
    
    Exit Sub

errHandle:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Fetch Row Style Error"
    End Select
    Exit Sub
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

On Error Resume Next

    MousePointer = vbHourglass
    If bSortOrd Then
        rsNoteCallBacks.Sort = "[" & rsNoteCallBacks.Fields(ColIndex).name & "]  ASC"
        bSortOrd = False
    Else
        rsNoteCallBacks.Sort = "[" & rsNoteCallBacks.Fields(ColIndex).name & "]  DESC"
        bSortOrd = True
    End If
    TDBGrid1.col = ColIndex
    rsNoteCallBacks.MoveFirst
    TDBGrid1.ReBind
    MousePointer = vbDefault

End Sub


Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)

On Error Resume Next
    bCheckAvailable = False
    'Debug.Print rsNoteCallBacks.Bookmark
    Text2.Text = TDBGrid1.Columns(2).Value
    Text3.Text = TDBGrid1.Columns(7).Value
    If Trim(TDBGrid1.Columns(6).Value) <> "" Then
        Check1.Value = TDBGrid1.Columns(6).Value
    Else
        Check1.Value = 0
    End If
    bCheckAvailable = True
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
            rsNoteCallBacks.MoveLast
            rsNoteCallBacks.MoveNext
        Else
            rsNoteCallBacks.MoveFirst
            rsNoteCallBacks.MovePrevious
        End If
        rsNoteCallBacks.Move offset
    Else
        rsNoteCallBacks.Move offset, StartLocation
    End If
        
    StartRow = rsNoteCallBacks.Bookmark
    Pos = rsNoteCallBacks.AbsolutePosition
    
    For Row = 0 To Rows
        If rsNoteCallBacks.BOF Or rsNoteCallBacks.EOF Then Exit For
        For col = 0 To cols
            If Trim(rsNoteCallBacks!note_msg) <> "" Then
                sOutPutMsg = Replace(Trim(rsNoteCallBacks!note_msg), "*##*", "'")
                sOutPutMsg = Replace(sOutPutMsg, "$++$", """")
            End If
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsNoteCallBacks!note_listid) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsNoteCallBacks!note_callback_time) & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsNoteCallBacks!note_company_name) & ""
                Case (3):   RowBuf.Value(Row, 3) = sOutPutMsg
                Case (4):   RowBuf.Value(Row, 4) = Trim(rsNoteCallBacks!note_datestamp) & ""
                Case (5):   RowBuf.Value(Row, 5) = Trim(rsNoteCallBacks!note_created_by) & ""
                Case (6):
                            If Trim(rsNoteCallBacks!note_callback_today) = 1 And Trim(rsNoteCallBacks!note_callback_today_date) = Trim(Text1.Text) Then
                                RowBuf.Value(Row, 6) = Trim(rsNoteCallBacks!note_callback_today) & ""
                            End If
                Case (7):   RowBuf.Value(Row, 7) = Trim(rsNoteCallBacks!note_index) & ""
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsNoteCallBacks.Bookmark
        RowsFetched = RowsFetched + 1
        rsNoteCallBacks.MoveNext
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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Command1.Value = True
    End If
    
End Sub
