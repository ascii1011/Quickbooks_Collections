VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmOldNotes 
   Caption         =   "Form2"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   ScaleHeight     =   4365
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Notes"
      Height          =   4095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8235
      Begin VB.TextBox Text6 
         Height          =   1695
         Left            =   3720
         MaxLength       =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1320
         Width           =   4395
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4860
         TabIndex        =   13
         Top             =   420
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Message"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         ToolTipText     =   "Save this message."
         Top             =   3600
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Left            =   5460
         TabIndex        =   11
         ToolTipText     =   "Clear current note, so i can create a new one."
         Top             =   3600
         Width           =   1275
      End
      Begin VB.TextBox Text11 
         Height          =   315
         Left            =   4860
         TabIndex        =   10
         ToolTipText     =   "Ex. '04/31/2004'"
         Top             =   900
         Width           =   1035
      End
      Begin VB.TextBox Text12 
         Height          =   315
         Left            =   6840
         TabIndex        =   9
         ToolTipText     =   "Ex. '4:00 PM'"
         Top             =   900
         Width           =   1275
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   7680
         Picture         =   "frmOldNotes.frx":0000
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   8
         ToolTipText     =   "Email an alert, concerning this customer."
         Top             =   3600
         Width           =   390
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   7440
         Picture         =   "frmOldNotes.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print a letter"
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6000
         Picture         =   "frmOldNotes.frx":05CC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   900
         Width           =   270
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   3720
         TabIndex        =   4
         Text            =   " "
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Insert"
         Height          =   315
         Left            =   6720
         TabIndex        =   3
         ToolTipText     =   "Clear current note, so i can create a new one."
         Top             =   3120
         Width           =   555
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Edit"
         Height          =   315
         Left            =   7500
         TabIndex        =   2
         ToolTipText     =   "Clear current note, so i can create a new one."
         Top             =   3120
         Width           =   555
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   7080
         Picture         =   "frmOldNotes.frx":0A0E
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   1
         ToolTipText     =   "Import Mail"
         Top             =   3600
         Width           =   390
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5741
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Created"
         Columns(0).DataField=   ""
         Columns(0).DataWidth=   1
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Message"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Callback"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "By"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1508"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1429"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1535"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1455"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1323"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1244"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=661"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=582"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=51,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=48,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=49,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=50,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
         _StyleDefs(73)  =   "Named:id=47:Open"
         _StyleDefs(74)  =   ":id=47,.parent=42,.fgcolor=&HFF&"
      End
      Begin VB.Label Label4 
         Caption         =   "Msg Date:"
         Height          =   195
         Left            =   3780
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Count:"
         Height          =   195
         Left            =   1860
         TabIndex        =   18
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   2460
         TabIndex        =   17
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Callback Date:"
         Height          =   195
         Left            =   3780
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Time:"
         Height          =   195
         Left            =   6420
         TabIndex        =   15
         Top             =   960
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmOldNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
