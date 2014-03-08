VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmInvoiceQry 
   Caption         =   "Collections"
   ClientHeight    =   11355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14760
   Icon            =   "frmInvoiceQry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11355
   ScaleWidth      =   14760
   Begin VB.Timer tTimeZone 
      Interval        =   60000
      Left            =   10980
      Top             =   9780
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   4260
      TabIndex        =   105
      Top             =   10620
      Width           =   2475
   End
   Begin VB.Frame Frame2 
      Caption         =   "Notes"
      Height          =   4095
      Left            =   4380
      TabIndex        =   65
      Top             =   4920
      Width           =   8295
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   6588
         _Version        =   393216
         TabHeight       =   423
         TabCaption(0)   =   "Messages:"
         TabPicture(0)   =   "frmInvoiceQry.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label12"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TDBGrid3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Create a Message/Callback"
         TabPicture(1)   =   "frmInvoiceQry.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command13"
         Tab(1).Control(1)=   "Check7"
         Tab(1).Control(2)=   "Picture6"
         Tab(1).Control(3)=   "Picture5"
         Tab(1).Control(4)=   "Command9"
         Tab(1).Control(5)=   "Command8"
         Tab(1).Control(6)=   "Combo7"
         Tab(1).Control(7)=   "Picture4"
         Tab(1).Control(8)=   "Command4"
         Tab(1).Control(9)=   "Picture3"
         Tab(1).Control(10)=   "Text12"
         Tab(1).Control(11)=   "Text11"
         Tab(1).Control(12)=   "Command3"
         Tab(1).Control(13)=   "Command2"
         Tab(1).Control(14)=   "Text5"
         Tab(1).Control(15)=   "Text6"
         Tab(1).Control(16)=   "Label31"
         Tab(1).Control(17)=   "Label26"
         Tab(1).Control(18)=   "Label25"
         Tab(1).Control(19)=   "Label17"
         Tab(1).Control(20)=   "Label16"
         Tab(1).Control(21)=   "Label4"
         Tab(1).ControlCount=   22
         TabCaption(2)   =   "Modify Note"
         TabPicture(2)   =   "frmInvoiceQry.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label27"
         Tab(2).Control(1)=   "Label28"
         Tab(2).Control(2)=   "Command11"
         Tab(2).Control(3)=   "Text14"
         Tab(2).Control(4)=   "Picture7"
         Tab(2).Control(5)=   "Picture8"
         Tab(2).Control(6)=   "Text16"
         Tab(2).Control(7)=   "Text15"
         Tab(2).Control(8)=   "Command12"
         Tab(2).ControlCount=   9
         Begin VB.CommandButton Command13 
            Caption         =   "Check"
            Height          =   315
            Left            =   -71220
            TabIndex        =   107
            ToolTipText     =   "Checks current length of message."
            Top             =   3300
            Width           =   675
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Priority Alert Completed."
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   -70320
            TabIndex        =   104
            Top             =   2220
            Width           =   3075
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Update Callback Date"
            Enabled         =   0   'False
            Height          =   315
            Left            =   -69000
            TabIndex        =   97
            ToolTipText     =   "Save this message."
            Top             =   1380
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            Enabled         =   0   'False
            Height          =   315
            Left            =   -70320
            TabIndex        =   94
            ToolTipText     =   "Ex. '04/31/2004'"
            Top             =   960
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox Text16 
            Enabled         =   0   'False
            Height          =   315
            Left            =   -68760
            TabIndex        =   93
            ToolTipText     =   "Ex. '4:00 PM'"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -69240
            Picture         =   "frmInvoiceQry.frx":091E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   92
            Top             =   960
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -67620
            Picture         =   "frmInvoiceQry.frx":0D60
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   91
            Top             =   960
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -67620
            Picture         =   "frmInvoiceQry.frx":0EEA
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   90
            Top             =   1140
            Width           =   390
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -69300
            Picture         =   "frmInvoiceQry.frx":1074
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   81
            ToolTipText     =   "Import Mail"
            Top             =   2580
            Width           =   390
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Edit"
            Height          =   315
            Left            =   -68460
            TabIndex        =   80
            ToolTipText     =   "Clear current note, so i can create a new one."
            Top             =   1800
            Width           =   555
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Insert"
            Height          =   315
            Left            =   -67800
            TabIndex        =   79
            ToolTipText     =   "Clear current note, so i can create a new one."
            Top             =   1800
            Width           =   555
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   -70320
            TabIndex        =   78
            Text            =   " "
            Top             =   1800
            Width           =   1695
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -69240
            Picture         =   "frmInvoiceQry.frx":11FE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   77
            Top             =   1140
            Width           =   270
         End
         Begin VB.CommandButton Command4 
            Height          =   495
            Left            =   -67740
            Picture         =   "frmInvoiceQry.frx":1640
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Print a letter"
            Top             =   420
            Width           =   495
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -68700
            Picture         =   "frmInvoiceQry.frx":1A82
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   75
            ToolTipText     =   "Email an alert, concerning this customer."
            Top             =   2580
            Width           =   390
         End
         Begin VB.TextBox Text12 
            Height          =   315
            Left            =   -68760
            TabIndex        =   74
            ToolTipText     =   "Ex. '4:00 PM'"
            Top             =   1140
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Height          =   315
            Left            =   -70320
            TabIndex        =   73
            ToolTipText     =   "Ex. '04/31/2004'"
            Top             =   1140
            Width           =   1035
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            Height          =   375
            Left            =   -68160
            TabIndex        =   72
            ToolTipText     =   "Clear current note, so i can create a new one."
            Top             =   3180
            Width           =   915
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save Message"
            Height          =   375
            Left            =   -70260
            TabIndex        =   71
            ToolTipText     =   "Save this message."
            Top             =   3180
            Width           =   1275
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   -70320
            TabIndex        =   70
            Top             =   540
            Width           =   2475
         End
         Begin VB.TextBox Text6 
            Height          =   2955
            Left            =   -74880
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            Top             =   300
            Width           =   4395
         End
         Begin VB.TextBox Text14 
            Height          =   3255
            Left            =   -74880
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            Top             =   360
            Width           =   4395
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Add message"
            Height          =   315
            Left            =   -70320
            TabIndex        =   67
            ToolTipText     =   "Save this message."
            Top             =   3240
            Width           =   1155
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid3 
            Height          =   2895
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   5106
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=6085"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=6006"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2884"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2805"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=1402"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1323"
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
         Begin VB.Label Label31 
            Caption         =   "Messages may be no longer then 4950 characters."
            Height          =   195
            Left            =   -74880
            TabIndex        =   106
            Top             =   3420
            Width           =   3615
         End
         Begin VB.Label Label28 
            Caption         =   "Callback Date:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -70320
            TabIndex        =   96
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Time:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -68760
            TabIndex        =   95
            Top             =   720
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label26 
            Caption         =   "Time:"
            Height          =   195
            Left            =   -68760
            TabIndex        =   89
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label25 
            Caption         =   "Callback Date:"
            Height          =   195
            Left            =   -70320
            TabIndex        =   88
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Msg Date:"
            Height          =   195
            Left            =   -70320
            TabIndex        =   87
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Mail Options:"
            Height          =   195
            Left            =   -70320
            TabIndex        =   86
            Top             =   2700
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Remarks:"
            Height          =   195
            Left            =   -70320
            TabIndex        =   85
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Height          =   195
            Left            =   7080
            TabIndex        =   84
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Count:"
            Height          =   195
            Left            =   6480
            TabIndex        =   83
            Top             =   3360
            Width           =   495
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   10020
      Top             =   9780
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5580
      TabIndex        =   55
      Top             =   9780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4140
      Picture         =   "frmInvoiceQry.frx":1C0C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   54
      ToolTipText     =   "Today's callbacks"
      Top             =   9900
      Width           =   270
   End
   Begin VB.CommandButton Command7 
      Caption         =   "X"
      Height          =   315
      Left            =   3120
      TabIndex        =   53
      Top             =   9900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      Caption         =   "User Filters"
      Height          =   9015
      Left            =   12780
      TabIndex        =   43
      Top             =   0
      Width           =   1875
      Begin VB.CommandButton Command10 
         Caption         =   "Stats"
         Height          =   375
         Left            =   360
         TabIndex        =   62
         ToolTipText     =   "Clear current note, so i can create a new one."
         Top             =   3480
         Width           =   1035
      End
      Begin VB.CheckBox Check6 
         Caption         =   "View Invoice Payments"
         Height          =   435
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   1635
      End
      Begin VB.CheckBox Check5 
         Caption         =   "View Customer Credits"
         Height          =   435
         Left            =   120
         TabIndex        =   48
         Top             =   1020
         Width           =   1635
      End
      Begin VB.CheckBox Check4 
         Caption         =   "View Invoice   Zero Balance"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   2700
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "View Media"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "View Open"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   660
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "View Customer Zero Balance"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contact Info"
      Height          =   9015
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin TrueDBGrid70.TDBGrid TDBGrid2 
         Height          =   3435
         Left            =   120
         TabIndex        =   40
         Top             =   4920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6059
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Date"
         Columns(0).DataField=   ""
         Columns(0).DataWidth=   100
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Invoice"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Balance"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1693"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1614"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1826"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1746"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
         _StyleDefs(69)  =   "Named:id=25:payment"
         _StyleDefs(70)  =   ":id=25,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(71)  =   "Named:id=26:Balance"
         _StyleDefs(72)  =   ":id=26,.parent=25,.fgcolor=&H0&,.borderColor=&H80000007&,.bold=-1,.fontsize=825"
         _StyleDefs(73)  =   ":id=26,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(74)  =   ":id=26,.fontname=MS Sans Serif"
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmInvoiceQry.frx":204E
         Left            =   1080
         List            =   "frmInvoiceQry.frx":2050
         TabIndex        =   61
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1080
         TabIndex        =   51
         Top             =   2220
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   50
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox Text13 
         Height          =   315
         Left            =   3000
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3480
         Picture         =   "frmInvoiceQry.frx":2052
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   31
         ToolTipText     =   "Email the Customer a Note."
         Top             =   4500
         Width           =   390
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   4500
         Width           =   2295
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Not Upfront"
         Height          =   195
         Left            =   1860
         TabIndex        =   28
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Upfront"
         Height          =   195
         Left            =   720
         TabIndex        =   27
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   1080
         TabIndex        =   25
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox Text8 
         Height          =   1335
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Top             =   2640
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInvoiceQry.frx":21DC
         Left            =   1200
         List            =   "frmInvoiceQry.frx":21DE
         TabIndex        =   9
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   2220
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   555
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1380
         Width           =   2775
      End
      Begin VB.Label Label24 
         Caption         =   "Web Status:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label23 
         Caption         =   "Count:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   8520
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   660
         TabIndex        =   41
         Top             =   8520
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Rep:"
         Height          =   195
         Left            =   2640
         TabIndex        =   37
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Email:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   4560
         Width           =   555
      End
      Begin VB.Label Label14 
         Caption         =   "Fax:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   4140
         Width           =   555
      End
      Begin VB.Label Label13 
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2700
         Width           =   675
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   2820
         TabIndex        =   22
         Top             =   8520
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Total Balance: $"
         Height          =   195
         Left            =   1620
         TabIndex        =   21
         Top             =   8520
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Importance:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Phone:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Contact:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Company:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search Options"
      Height          =   2355
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   3915
      Begin VB.CommandButton Command1 
         Caption         =   "Query"
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Callbacks"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Invoices"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picCal2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2580
         Picture         =   "frmInvoiceQry.frx":21E0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         Top             =   1860
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCal1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2580
         Picture         =   "frmInvoiceQry.frx":2622
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         Top             =   1500
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CheckBox IncludeLineItems 
         Caption         =   "Include Line Items"
         Height          =   255
         Left            =   540
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox ToTxnDate 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   1860
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox FromTxnDate 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1500
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "To (mm/dd/yyyy)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1860
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "From (mm/dd/yyyy)"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customers"
      Height          =   4935
      Left            =   4380
      TabIndex        =   7
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton Command14 
         Caption         =   "Map"
         Height          =   315
         Left            =   4260
         TabIndex        =   111
         Top             =   4500
         Width           =   615
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         ItemData        =   "frmInvoiceQry.frx":2A64
         Left            =   1080
         List            =   "frmInvoiceQry.frx":2A66
         TabIndex        =   108
         Top             =   4500
         Width           =   2175
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "frmInvoiceQry.frx":2A68
         Left            =   1080
         List            =   "frmInvoiceQry.frx":2A6A
         TabIndex        =   102
         Top             =   4140
         Width           =   2175
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Not Upfront"
         Height          =   195
         Left            =   2640
         TabIndex        =   101
         Top             =   3780
         Width           =   1155
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Upfront"
         Height          =   195
         Left            =   1680
         TabIndex        =   100
         Top             =   3780
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "All"
         Height          =   195
         Left            =   1080
         TabIndex        =   98
         Top             =   3780
         Width           =   855
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmInvoiceQry.frx":2A6C
         Left            =   2520
         List            =   "frmInvoiceQry.frx":2A79
         TabIndex        =   58
         Text            =   "Starts with"
         Top             =   3300
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmInvoiceQry.frx":2AA1
         Left            =   1020
         List            =   "frmInvoiceQry.frx":2AB7
         TabIndex        =   57
         Text            =   "Customers"
         Top             =   3300
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Height          =   555
         Left            =   7560
         Picture         =   "frmInvoiceQry.frx":2AF1
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Today's callbacks"
         Top             =   4200
         Width           =   555
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-->"
         Height          =   315
         Left            =   7620
         TabIndex        =   38
         ToolTipText     =   "View Options"
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   4020
         TabIndex        =   34
         Top             =   3300
         Width           =   2115
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   2955
         Left            =   180
         TabIndex        =   39
         Top             =   240
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   5212
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).DataWidth=   1
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Acc #"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Customers"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Balance"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Status"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Rep"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Importance"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "UpFront"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=159"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=79"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1508"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1429"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4471"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4392"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1535"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1455"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1402"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1323"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=847"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=767"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1614"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1535"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=1191"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1111"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
         _StyleDefs(90)  =   ":id=47,.parent=42,.fgcolor=&HFF&"
      End
      Begin VB.Label lbRegionTime 
         Height          =   195
         Left            =   3360
         TabIndex        =   110
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "Region:"
         Height          =   195
         Left            =   180
         TabIndex        =   109
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label30 
         Caption         =   "Importance:"
         Height          =   195
         Left            =   180
         TabIndex        =   103
         Top             =   4200
         Width           =   915
      End
      Begin VB.Label Label29 
         Caption         =   "Show only:"
         Height          =   195
         Left            =   180
         TabIndex        =   99
         Top             =   3780
         Width           =   795
      End
      Begin VB.Label Label20 
         Caption         =   "Search by:"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   3360
         Width           =   795
      End
      Begin VB.Label Label19 
         Caption         =   "Count:"
         Height          =   195
         Left            =   7020
         TabIndex        =   33
         Top             =   3300
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   7500
         TabIndex        =   32
         Top             =   3300
         Width           =   615
      End
   End
   Begin SHDocVwCtl.WebBrowser wb3 
      Height          =   555
      Left            =   8760
      TabIndex        =   60
      Top             =   9900
      Width           =   675
      ExtentX         =   1191
      ExtentY         =   979
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
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   180
      TabIndex        =   64
      Top             =   4560
      Width           =   3675
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   180
      TabIndex        =   63
      Top             =   4560
      Width           =   3675
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   0
      Top             =   2520
   End
   Begin VB.Label Label7 
      Caption         =   "ListID:"
      Height          =   195
      Left            =   4800
      TabIndex        =   56
      Top             =   9900
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmInvoiceQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents adoDataConn As ADODB.Connection
Attribute adoDataConn.VB_VarHelpID = -1
Private WithEvents rsRecordSet As ADODB.Recordset
Attribute rsRecordSet.VB_VarHelpID = -1
 
Dim col As TrueDBGrid70.Column
Dim cols As TrueDBGrid70.Columns

Dim strLineArray(200) As String
Public sOrderBy As String

Dim bSortOrd As Boolean
    
''''''''''''''''''''''''''''''grab customer info'''''''''''''''''''''''''''
Dim requestXML  As String
Dim responseXML As String

'''''''''''''''''''''''''''grab invoice info''''''''''''''''''''''''''''''''
Dim fromDate As String
Dim toDate As String


'These are the variables which will be set to do the invoice query
Dim strRefNumber As String
Dim strFromDateTime As String
Dim strToDateTime As String
Dim strDateQueryType As String
Dim strDateMacro As String
Dim strCustomerJob As String
Dim booCustomerWithChildren As Boolean
Dim strAccount As String
Dim strListID As String
Dim booAccountWithChildren As Boolean
Dim strFromRefNumberRange As String
Dim strToRefNumberRange As String
Dim strRefNumberPiece As String
Dim strRefNumberCriteria As String
Dim strPaidStatus As String

''''''''''''''''''''''''sql part
Public rsCallBacks          As New ADODB.Recordset
Public rsImportance         As New ADODB.Recordset
Public rsNotes              As New ADODB.Recordset
Public sNoteClickedOn       As String
Public lNoteClickedOnIndex  As Long

Public rsGrabCustomerInfo   As New ADODB.Recordset
Public rsGrabInvoiceInfo   As New ADODB.Recordset
Public rsGrabPaymentInfo   As New ADODB.Recordset
Public rsGrabInvoiceComp   As New ADODB.Recordset

Public rsGrabCustomerDtls   As New ADODB.Recordset
Public rsGrabPartRemarks   As New ADODB.Recordset

''''''''''''''gen
Dim sCustTxnID As String
Dim sCustRow As String

''''''''''alert variables
Dim Alert_MaxDollor             As Currency
Dim Alert_MaxInvoices           As Long
Dim Alert_StartLevel            As Long

Dim Alert_Table_MaxDollor       As Currency
Dim Alert_Table_MaxInvoices     As Long
Dim Alert_Table_StartLevel      As Long
Dim Alert_Table_Reason          As String

Dim Insert_Alert_Table_MaxDollor       As Currency
Dim Insert_Alert_Table_MaxInvoices     As Long
Dim Insert_Alert_Table_StartLevel      As Long
Dim Insert_Alert_Table_Reason          As String

Dim Current_StartLevel          As Long


Dim Character As IAgentCtlCharacterEx





Sub prcGrabNote()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    
On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    If rsNotes.State = 1 Then
        Set rsNotes = Nothing
    End If
    'List1.AddItem "Retrieving settings for:" & sUser & ":" & Date & ":" & Time
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_note_sp"
    
    'reg_list_user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(ListID) & "")
    cmdCommand.Parameters.Append parParameter
        
    Set rsNotes = cmdCommand.Execute
    
    Label12.Caption = rsNotes.RecordCount
    If rsNotes.RecordCount > 0 Then
        TDBGrid3.ApproxCount = rsNotes.RecordCount
    End If
    'If Not rsNotes.EOF Then
        'TDBGrid3.ReBind
    'Else
        'List1.AddItem "No saved settings found, user input required."
        TDBGrid3.ReBind
    'End If
    
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
    Set rsNotes = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prcUpdate 6, Check1
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prcUpdate 10, Check2
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prcUpdate 11, Check3
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prcUpdate 46, Check4
End Sub

Private Sub Check5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prcUpdate 44, Check5
End Sub

Private Sub Check6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prcUpdate 45, Check6
End Sub


Sub prcUpdate(sindex As String, cCheck As CheckBox)
    frmMain.prcUpdateOneAttr sindex, cCheck.Value, sUser
    frmMain.funGrabProfile
    prcUpdateFilterList
    frmInvoiceQry.prcRefresh
End Sub

Private Sub Combo4_Change()
    prcFind2 LCase(Trim(Text7.Text)), Trim(Combo4.Text), Trim(Combo5.Text)
End Sub

Private Sub Combo4_Click()
    prcFind2 LCase(Trim(Text7.Text)), Trim(Combo4.Text), Trim(Combo5.Text)
End Sub

Private Sub Combo5_Change()
    prcFind2 LCase(Trim(Text7.Text)), Trim(Combo4.Text), Trim(Combo5.Text)
End Sub

Private Sub Combo5_Click()
    prcFind2 LCase(Trim(Text7.Text)), Trim(Combo4.Text), Trim(Combo5.Text)
End Sub

Private Sub Combo8_Click()
    prcGrabCustomerInfo sOrderBy
End Sub


Private Sub Combo9_Click()
    prcGrabCustomerInfo sOrderBy
End Sub

Private Sub Command10_Click()
    frmMain.prcShowFrmQBInStats
End Sub

Private Sub Command11_Click()
    prcAddToMessage
End Sub

Sub prcAddToMessage()
    Dim sUpdatedMsg As String
    Dim iTempRowNumber As Integer
    
    If Trim(Text14.Text) <> "" Then
    
        sUpdatedMsg = Left(Trim(Text14.Text), 1990)
        sUpdatedMsg = Replace(sUpdatedMsg, "'", "*##*")
        sUpdatedMsg = Replace(sUpdatedMsg, """", "$++$")
        
        prcInsert_AddMsg sUpdatedMsg, lNoteClickedOnIndex
        'prcGrabNote
    End If

On Error Resume Next

    iTempRowNumber = sNoteClickedOn
    TDBGrid3.Bookmark = iTempRowNumber
    
    SSTab1.Tab = 1
    Text14.Text = ""
    prcExplodeMsg
    prcGrab_Addmsgs lNoteClickedOnIndex, Text6
End Sub

Sub prcGrab_Addmsgs(Ni As Long, txtBox As TextBox)
    Dim Addmsgs As Sql_Results_Struct
    
    Addmsgs.Query = " select * from qb_note_addmsg " & _
                    " where note_index = '" & Ni & "' " & _
                    " order by nadd_datestamp asc "
    SQL_Query_auto Addmsgs.Query, Addmsgs.Data
        
    If Not Addmsgs.Data.EOF Then
        Addmsgs.Data.MoveFirst
        While Not Addmsgs.Data.EOF
            txtBox.Text = txtBox.Text & vbNewLine & vbNewLine
            txtBox.Text = txtBox.Text & "****Updaded By " & Trim(Addmsgs.Data!nadd_created_by) & " at " & Trim(Addmsgs.Data!nadd_datestamp) & "****" & vbNewLine
            txtBox.Text = txtBox.Text & Trim(Addmsgs.Data!nadd_msg) & ""
            Addmsgs.Data.MoveNext
        Wend
    End If
    
    SQL_Close_Clear Addmsgs.Data
End Sub


Sub prcInsert_AddMsg(sMsg As String, lNoteIndex As Long)
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim sSQL                    As String
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    sSQL = " insert into qb_note_addmsg " & _
            " ( note_index, nadd_datestamp, nadd_msg, nadd_created_by ) " & _
            " values " & _
            " ( '" & lNoteIndex & "', '" & Now & "', '" & sMsg & "', '" & sUser & "' ) "
                
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & vbNewLine & _
                    "STMT addmsg:" & sSQL & vbNewLine & vbNewLine & "Try again?", vbExclamation + vbYesNo, "prcUpdateNote")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Sub prcAddToMessage_old()
    Dim sUpdatedMsg As String
    Dim iTempRowNumber As Integer
    
    If Trim(Text14.Text) <> "" Then
    
        sUpdatedMsg = Left(Trim(TDBGrid3.Columns(1).Value) & vbNewLine & vbNewLine & _
                            "****Updaded By " & sUser & " at " & Date & " " & Time & "****" & vbNewLine & _
                            Trim(Text14.Text), 4950)
                            
        sUpdatedMsg = Replace(sUpdatedMsg, "'", "*##*")
        sUpdatedMsg = Replace(sUpdatedMsg, """", "$++$")
        
        prcUpdateNote sUpdatedMsg, lNoteClickedOnIndex
        prcGrabNote
    End If

On Error Resume Next

    iTempRowNumber = sNoteClickedOn
    TDBGrid3.Bookmark = iTempRowNumber
    
    SSTab1.Tab = 1
    Text14.Text = ""
    prcExplodeMsg
End Sub



Sub prcUpdateNote(sMsg As String, lNoteIndex As Long)
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim sSQL                    As String
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    sSQL = " update qb_note " & _
            " set note_msg = '" & sMsg & "' " & _
            " where note_index = '" & lNoteIndex & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & vbNewLine & _
                    "STMT:" & sSQL & vbNewLine & vbNewLine & "Try again?", vbExclamation + vbYesNo, "prcUpdateNote")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub Command13_Click()
    MsgBox "length of current message is: " & Len(Trim(Text6.Text))
End Sub

Private Sub Command14_Click()
    frmTimeZoneMap.Show
End Sub

Private Sub Command5_Click()
    frmMain.prcShowFrmCallbacks
End Sub

Private Sub Command7_Click()
    Dim place
    Dim placetemp
    
    
    'MsgBox TDBGrid1.ApproxCount
    
    TDBGrid1.Bookmark = 1
    place = 1
    placetemp = place
        If place <= 12 Then
            TDBGrid1.Row = place
        Else
            While placetemp > 12
                TDBGrid1.Row = 12
                TDBGrid1.Bookmark = TDBGrid1.RowBookmark(TDBGrid1.Row)
                placetemp = placetemp - 12
                If place <= 12 Then
                    TDBGrid1.Bookmark = TDBGrid1.RowBookmark(place)
                    TDBGrid1.Row = placetemp
                End If
            Wend
        End If
        
        
    'MsgBox TDBGrid1.Bookmark
    'TDBGrid1.Scroll 0, 2
End Sub

Private Sub Command7_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim place
    Dim placetemp
    
    
    'MsgBox TDBGrid1.ApproxCount
    
    TDBGrid1.Bookmark = 1
    place = 1
    placetemp = place
        If place <= 12 Then
            TDBGrid1.Row = place
        Else
            While placetemp > 12
                TDBGrid1.Row = 12
                TDBGrid1.Bookmark = TDBGrid1.RowBookmark(TDBGrid1.Row)
                placetemp = placetemp - 12
                If place <= 12 Then
                    TDBGrid1.Bookmark = TDBGrid1.RowBookmark(place)
                    TDBGrid1.Row = placetemp
                End If
            Wend
        End If
        
        
    'MsgBox TDBGrid1.Bookmark
    'TDBGrid1.Scroll 0, 2
End Sub

Private Sub Command7_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim place
    Dim placetemp
    
    
    'MsgBox TDBGrid1.ApproxCount
    
    TDBGrid1.Bookmark = 1
    place = 1
    placetemp = place
        If place <= 12 Then
            TDBGrid1.Row = place
        Else
            While placetemp > 12
                TDBGrid1.Row = 12
                TDBGrid1.Bookmark = TDBGrid1.RowBookmark(TDBGrid1.Row)
                placetemp = placetemp - 12
                If place <= 12 Then
                    TDBGrid1.Bookmark = TDBGrid1.RowBookmark(place)
                    TDBGrid1.Row = placetemp
                End If
            Wend
        End If
        
        
    'MsgBox TDBGrid1.Bookmark
    'TDBGrid1.Scroll 0, 2
End Sub





Private Sub Command8_Click()
    Dim sRemarkId() As String
    
    If Combo7.Text <> "" Then
        sRemarkId = Split(Combo7.Text, ":")
        If UBound(sRemarkId) > 0 Then
            If Trim(sRemarkId(0)) = "Date" Then
                Text6.Text = Text6.Text & vbNewLine & sRemarkId(1) & ":" & sRemarkId(2) & ":" & sRemarkId(3)
            Else
                prcGrabAPartRemark Trim(sRemarkId(0))
            End If
        End If
    End If
End Sub

Sub prcGrabAPartRemark(sID As String)
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
    cmdCommand.CommandText = " select remark_msg from qbx_remarks where remark_index = '" & sID & "' "
            
    Set rsSingleRemark = cmdCommand.Execute
    
    If Not rsSingleRemark.EOF Then
        If rsSingleRemark.RecordCount = 1 Then
            rsSingleRemark.MoveFirst
            Text6.Text = Text6.Text & vbNewLine & Trim(rsSingleRemark!remark_msg) & ""
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

Private Sub Command9_Click()
    frmMain.prcShowFrmRemarks
End Sub

Private Sub Form_Activate()
    iInvoiceQryLostFocus = 0
    Debug.Print "Inv active"
End Sub

Private Sub Form_Deactivate()
    iInvoiceQryLostFocus = 1
    Debug.Print "Inv notactive"
End Sub

Private Sub Form_GotFocus()
    Debug.Print "Inv Got Focus"
End Sub

Private Sub Form_LostFocus()
    Debug.Print "Inv Lost Focus"
End Sub

Private Sub Label3_Click()
    prcLoadAgent
End Sub


Sub prcLoadAgent()
    Dim temp
On Error Resume Next
    Set Character = Nothing
    Agent1.Characters.Unload "genie"
        
    '-- Load the new character
On Error GoTo errhandler
    Agent1.Characters.Load "Genie", "genie.acs"
    
    Set Character = Agent1.Characters("Genie")
    
    '-- Set the character's language
    Character.LanguageID = &H409
    
    Character.Left = (Me.Left + 3450) / Screen.TwipsPerPixelX
    Character.top = (Me.top + 900) / Screen.TwipsPerPixelY

    Character.Show
    
    Character.Play "GestureUp"

    If Not Trim(Combo3.Text) = "" Then
        Character.Speak Trim(Combo3.Text)
    Else
        Character.Speak "No phone number available"
    End If
    
    Character.Play "Hide"
    Pause 8
    While Character.IdleOn = True
        Agent1.Characters.Unload "genie"
    Wend
        
    Exit Sub
    
errhandler:

        Set Character = Nothing
End Sub

Private Sub Label9_Click()
    If List2.Visible = True Then
        List2.Visible = False
    Else
        List2.Visible = True
    End If
End Sub

Private Sub List1_DblClick()
    List1.Visible = False
    TDBGrid2.Visible = True
End Sub

Private Sub Option5_Click()
    prcGrabCustomerInfo sOrderBy
End Sub

Private Sub Option6_Click()
    prcGrabCustomerInfo sOrderBy
End Sub

Private Sub Option7_Click()
    prcGrabCustomerInfo sOrderBy
End Sub

Private Sub Picture4_Click()
    iCalendarRequest = 2
    frmCalendar.Show
End Sub

Private Sub Picture5_Click()
    frmMain.prcShowemailimport
End Sub

Private Sub Picture6_Click()
    Text12.Text = Time
End Sub

Private Sub TDBGrid2_DblClick()
    'List1.Visible = True
    'TDBGrid2.Visible = False
End Sub

Private Sub TDBGrid3_DblClick()
    'Text6.Text = Trim(TDBGrid3.Columns(1).Value)
    
    If SQL_ReConnect_old(frmMain.cnMC) = False Then
        frmMain.StatusBar1.Panels.Item(6).Text = "Not Connected."
        Exit Sub
    End If
    frmMain.StatusBar1.Panels.Item(6).Text = "Connected."
    
    prcEnableModifyNotes
    prcExplodeMsg
    If lNoteClickedOnIndex <> 0 Then prcGrab_Addmsgs lNoteClickedOnIndex, Text6
End Sub

Sub prcExplodeMsg()
    Dim sDateIndex
    Dim sCreated
    Dim sTmpMsg As String
    
    Command2.Enabled = False
    Text6.Text = ""
    lNoteClickedOnIndex = 0
    
On Error GoTo errExp:

    sNoteClickedOn = TDBGrid3.Bookmark
    
    sDateIndex = Split(Trim(TDBGrid3.Columns(0).Value), "::")
    If UBound(sDateIndex) = 1 Then
        Text5.Text = Trim(sDateIndex(0))
        lNoteClickedOnIndex = Trim(sDateIndex(1))
    End If
    
    sTmpMsg = Replace(funGrabNoteByID(lNoteClickedOnIndex), "*##*", "'")
    sTmpMsg = Replace(sTmpMsg, "$++$", """")
    Text6.Text = sTmpMsg
    
    sCreated = Split(Trim(TDBGrid3.Columns(2).Value), " ")
    If UBound(sCreated) > 0 Then
        If UBound(sCreated) = 1 Then
            Text12.Text = sCreated(0) & " " & sCreated(1)
        ElseIf UBound(sCreated) = 2 Then
            Text11.Text = sCreated(0)
            Text12.Text = sCreated(1) & " " & sCreated(2)
        End If
    Else
        Text11.Text = ""
        Text12.Text = ""
    End If
    
    sCreated = Split(Trim(TDBGrid3.Columns(3).Value), "$$$")
    If UBound(sCreated) > 0 Then
        Text6.Text = Trim(Text6.Text) & vbNewLine & vbNewLine & "By: " & sCreated(0)
    End If
    Exit Sub
    
errExp:
    MsgBox Err.Number
End Sub

Function funGrabNoteByID(sID As Long) As String
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsNoteMsg  As New ADODB.Recordset

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
    
    If rsNoteMsg.State = 1 Then
        Set rsNoteMsg = Nothing
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select note_msg from qb_note where note_index = '" & sID & "' "
            
    Set rsNoteMsg = cmdCommand.Execute
    
    If Not rsNoteMsg.EOF Then
        If rsNoteMsg.RecordCount = 1 Then
            rsNoteMsg.MoveFirst
            funGrabNoteByID = Trim(rsNoteMsg!note_msg) & ""
        End If
    End If
    
    Set rsNoteMsg = Nothing
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
    Set rsNoteMsg = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function

Private Sub TDBGrid3_Error(ByVal DataError As Integer, Response As Integer)
On Error Resume Next

    Response = 0
End Sub

Private Sub TDBGrid3_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
Dim ColIndex As Integer, col As Integer
    Dim RowsFetched As Integer, Row As Long
    Dim StartRow As Variant
    Dim Response As Long
    Dim cols As Long
    Dim Rows As Long
    Dim Pos As Long
    Dim sOutPutMsg As String
    Dim strlen As Integer

On Error GoTo NoRead
    
    'If bCancelRead Then Exit Sub
    cols = RowBuf.ColumnCount - 1
    Rows = RowBuf.RowCount - 1
    RowsFetched = 0
    
    If IsNull(StartLocation) Then
        If offset < 0 Then
            rsNotes.MoveLast
            rsNotes.MoveNext
        Else
            rsNotes.MoveFirst
            rsNotes.MovePrevious
        End If
        rsNotes.Move offset
    Else
        rsNotes.Move offset, StartLocation
    End If
        
    StartRow = rsNotes.Bookmark
    Pos = rsNotes.AbsolutePosition
    
    For Row = 0 To Rows
        If rsNotes.BOF Or rsNotes.EOF Then Exit For
        For col = 0 To cols
            'strlen = Len(rsNotes!Status)
            If Trim(rsNotes!note_msg) <> "" Then
                sOutPutMsg = Replace(Trim(rsNotes!note_msg), "*##*", "'")
                sOutPutMsg = Replace(sOutPutMsg, "$++$", """")
            End If
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsNotes!note_datestamp) & "::" & Trim(rsNotes!note_index)
                Case (1):   RowBuf.Value(Row, 1) = Left(sOutPutMsg, 100)
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsNotes!note_callback_date) & " " & Trim(rsNotes!note_callback_time) & ""
                Case (3):   RowBuf.Value(Row, 3) = Trim(rsNotes!note_created_by) & "$$$" & Trim(rsNotes!note_datestamp)
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsNotes.Bookmark
        RowsFetched = RowsFetched + 1
        rsNotes.MoveNext
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

Sub prcInsertNote(custinfo() As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sSQL As String

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    'List1.AddItem "Retrieving settings for:" & sUser & ":" & Date & ":" & Time
    
    
    sSQL = " insert qb_note " & _
        " (note_listid, note_msg, note_callback_date, note_callback_time, " & _
        " note_company_name,  note_created_by, note_datestamp, note_msg_backup, " & _
        " note_state, note_company_amount, note_company_status) " & _
        " values " & _
        " ('" & custinfo(6) & "', '" & custinfo(5) & "', '" & custinfo(3) & "', '" & custinfo(4) & "', " & _
        " '" & custinfo(0) & "', '" & sUser & "', '" & Now & "', '" & custinfo(5) & "', " & _
        " '" & custinfo(7) & "', '" & custinfo(1) & "', '" & custinfo(2) & "') "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
            

    'Set cmdCommand.ActiveConnection = frmMain.cnMC
    'cmdCommand.CommandType = adCmdStoredProc
    'cmdCommand.CommandText = "insert_note_sp"
    ''listid
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, custinfo(6) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'smsg
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 3000, custinfo(5) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'note_callback_date
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 15, custinfo(3) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'note_callback_time
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 15, custinfo(4) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'note_company_name
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 200, custinfo(0) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'state
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, custinfo(7) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'user
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sUser) & "")
    'cmdCommand.Parameters.Append parParameter
    
    'amount
    'If custinfo(1) = "0.00" Then
    '    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, "0")
    '    cmdCommand.Parameters.Append parParameter
    'Else
    '    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, "")
    '    cmdCommand.Parameters.Append parParameter
    'End If
    
    'jobstatus
    'Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, Trim(custinfo(2)) & "")
    'cmdCommand.Parameters.Append parParameter
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            'prcLog Now & "--prcGrabRegList:connected(" & frmMonitor.frmmain.cnmc.State & "), err:" & Err.Description
            'List1.AddItem "unable to prcGrabRegList for " & sUser
            'prcLog "errlog.txt", Now & "--GrabRegList--" & Err.Description
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Inserting Note Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Function funGrabImportance(iWhere As Integer, sTempListID As String) As String
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    
    funGrabImportance = ""

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
    'List1.AddItem "Retrieving settings for:" & sUser & ":" & Date & ":" & Time
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_Importance_sp"
    
    'reg_list_user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sTempListID) & "")
    cmdCommand.Parameters.Append parParameter
        
    Set rsImportance = cmdCommand.Execute
    
    If rsImportance.RecordCount = 1 Then
        If Not rsImportance.EOF Then
            rsImportance.MoveLast
            rsImportance.MoveFirst
            
            If iWhere = 2 Then
                funGrabImportance = Trim(rsImportance!importance_name)
            ElseIf iWhere = 1 Then
                Combo1.Text = Trim(rsImportance!importance_name)
                If rsImportance!UpFront = "True" Then
                    Option3.Value = True
                    Option4.Value = False
                ElseIf rsImportance!UpFront = "False" Then
                    Option3.Value = False
                    Option4.Value = True
                End If
            End If
        End If
    Else
        prcInsertImportance 99, ""
    End If
    
    Set rsImportance = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            'prcLog Now & "--prcGrabRegList:connected(" & frmMonitor.frmmain.cnmc.State & "), err:" & Err.Description
            'List1.AddItem "unable to prcGrabRegList for " & sUser
            'prcLog "errlog.txt", Now & "--GrabRegList--" & Err.Description
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Importance Record Opening Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsImportance = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function



Sub prcInsertImportance(sType As String, sName As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    'List1.AddItem "Retrieving settings for:" & sUser & ":" & Date & ":" & Time
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "insert_Importance_sp"
    
    'listid
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(ListID) & "")
    cmdCommand.Parameters.Append parParameter
    
    'type
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 10, Trim(sType) & "")
    cmdCommand.Parameters.Append parParameter
    
    'name
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sName) & "")
    cmdCommand.Parameters.Append parParameter
    
    'user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sUser) & "")
    cmdCommand.Parameters.Append parParameter
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            'prcLog Now & "--prcGrabRegList:connected(" & frmMonitor.frmmain.cnmc.State & "), err:" & Err.Description
            'List1.AddItem "unable to prcGrabRegList for " & sUser
            'prcLog "errlog.txt", Now & "--GrabRegList--" & Err.Description
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Inserting Importance Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub




Sub prcUpdateImportance(sType As String, sName As String, sOption3 As String, sOption4 As String)
    Dim Response
    Dim sUpfront As String
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim i As Integer

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    sType = ""
    For i = 0 To UBound(aryGImportLvl) - 1
        If LCase(sName) = LCase(aryGImportLvl(i, 2)) Then
            sName = aryGImportLvl(i, 2)
            sType = aryGImportLvl(i, 1)
            i = UBound(aryGImportLvl)
        End If
    Next i
    If sType = "" Then
        Combo1.Text = ""
        Exit Sub
    End If
    'If LCase(sName) = "collections" Then
    '    sName = "Collections"
    '    sType = 1
    'ElseIf LCase(sName) = "high" Then
    '    sName = "High"
    '    sType = 2
    'ElseIf LCase(sName) = "medium" Then
    '    sName = "Medium"
    '    sType = 3
    'ElseIf LCase(sName) = "low" Then
    '    sName = "Low"
    '    sType = 4
    'ElseIf LCase(sName) = "sales" Then
    '    sName = "Sales"
    '    sType = 5
    'ElseIf LCase(sName) = "discuss" Then
    '    sName = "Discuss"
    '    sType = 6
    'ElseIf LCase(sName) = "temporary" Then
    '    sName = "Temporary"
    '    sType = 7
    'ElseIf LCase(sName) = "" Then
    '    sName = ""
    '    sType = 99
    'Else
    '    Combo1.Text = ""
    '    Exit Sub
    'End If
    
    sUpfront = ""
    If sOption3 = "True" Then
        sUpfront = "True"
    ElseIf sOption4 = "True" Then
        sUpfront = "False"
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "update_importance_sp"
    
    'listid
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(ListID) & "")
    cmdCommand.Parameters.Append parParameter
    
    'type
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 2000, Trim(sType) & "")
    cmdCommand.Parameters.Append parParameter
    
    'name
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sName) & "")
    cmdCommand.Parameters.Append parParameter
    
    'user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(sUser) & "")
    cmdCommand.Parameters.Append parParameter
    
    'upfront
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 10, Trim(sUpfront) & "")
    cmdCommand.Parameters.Append parParameter
    
    cmdCommand.Execute
        
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            'prcLog Now & "--prcInsertUserDrives:connected(" & frmMonitor.frmmain.cnmc.State & "), err:" & Err.Description
            'List1.AddItem "unable to prcInsertUserDrives for " & sUser
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "updating Importance Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Private Sub ClearQueryVariables()
  strRefNumber = ""
  strFromDateTime = ""
  strToDateTime = ""
  strDateQueryType = ""
  strDateMacro = ""
  strCustomerJob = ""
  booCustomerWithChildren = False
  strAccount = ""
  booAccountWithChildren = False
  strFromRefNumberRange = ""
  strToRefNumberRange = ""
  strRefNumberPiece = ""
  strRefNumberCriteria = ""
  strPaidStatus = ""
End Sub





Private Sub Combo1_Click()
    prcUpdateImportance 0, Trim(Combo1.Text), Trim(Option3.Value), Trim(Option4.Value)
    prcGrabCustomerInfo sOrderBy
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    'now check status for alerts
    ' grab alert settings
    prcGrabAlertSettings
    ' grab current importance type setting
    prcGetCurrentImportanceType
    '1) If this company is in the alert table then go further
    If funIsCompanyInAlerts = True Then
        If Alert_Table_StartLevel <> Current_StartLevel Then
            ''''''2) if one of the alert settings is still over threshhold then leave it alone
            If Current_StartLevel > Alert_StartLevel And Alert_Table_MaxDollor < Alert_MaxDollor And Alert_Table_MaxInvoices < Alert_MaxInvoices Then
                '3) else delete that record and update the table that is displayed if the form is active
                'prcRemove
                prcRemoveCustFromAlert
                
                If sFrmPriorityAlerts = 1 Then
                    frmPriorityAlerts.prcGrabPriorityAlerts "True", "Name Asc"
            frmPriorityAlerts.prcSearch "name", ""
                End If
            End If
        End If
    ElseIf Current_StartLevel <= Alert_StartLevel Then
        Insert_Alert_Table_MaxInvoices = funMoreThanOneInvoice
        prcAddCustToAlert
        
        If sFrmPriorityAlerts = 1 Then
            frmPriorityAlerts.prcGrabPriorityAlerts "True", "Name Asc"
            frmPriorityAlerts.prcSearch "name", Trim(CurrCust.CompanyFullName)
        End If
    End If
End Sub

Function funMoreThanOneInvoice() As Integer
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsGrabInv As New ADODB.Recordset
    
On Error GoTo errHandle:

    funMoreThanOneInvoice = 0
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select count(*) as invcount from qbx_inv where inv_customerref_listid = '" & Trim(strListID) & "' and inv_enabled = '1' and inv_balanceremaining <> '0.00' "
        
    Set rsGrabInv = cmdCommand.Execute
    
    If Not rsGrabInv.EOF Then
        rsGrabInv.MoveFirst
        funMoreThanOneInvoice = Trim(rsGrabInv!invcount)
    End If
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            'Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Checking Invoice Count Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
End Function

Sub prcAddCustToAlert()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sSQL As String

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If

    sSQL = " insert into qbx_alerts " & _
        "(alert_id, alert_importance, alert_importance_type, " & _
        " alert_total_balance, alert_total_invoices, alert_reason, alert_datetime, " & _
        " cust_isactive, cust_fullname) " & _
        " values " & _
        " ('" & strListID & "', '" & Trim(Combo1.Text) & "', '" & Current_StartLevel & "', " & _
        " CONVERT(money, '" & CurrCust.Balance & "'), '" & Insert_Alert_Table_MaxInvoices & "', ':Level too high','" & Now & "', " & _
        " 'True', '" & Trim(Text2.Text) & "') "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Is Company in Alerts?")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Sub prcGetCurrentImportanceType()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsCustAlerts As New ADODB.Recordset
    Dim sSQL As String

On Error GoTo errHandle:

    Current_StartLevel = 0
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If

    sSQL = " select * from qbx_cust where cust_listid = '" & strListID & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    Set rsCustAlerts = cmdCommand.Execute
    
    If Not rsCustAlerts.EOF Then
        rsCustAlerts.MoveFirst
        Current_StartLevel = Trim(rsCustAlerts!importance_type) & ""
    End If
    
    Set rsCustAlerts = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Is Company in Alerts?")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsCustAlerts = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcRemoveCustFromAlert()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sSQL As String

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If

    sSQL = " delete from qbx_alerts where alert_id = '" & strListID & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Is Company in Alerts?")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Function funIsCompanyInAlerts() As Boolean
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsCheckOnListIdInAlerts As New ADODB.Recordset
    Dim sSQL As String

On Error GoTo errHandle:

    funIsCompanyInAlerts = False
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Function
    End If
    
    sSQL = " select * from qbx_alerts where alert_id = '" & strListID & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    Set rsCheckOnListIdInAlerts = cmdCommand.Execute
    
    If Not rsCheckOnListIdInAlerts.EOF Then
        rsCheckOnListIdInAlerts.MoveFirst
        
        Alert_Table_Reason = Trim(rsCheckOnListIdInAlerts!Alert_Reason) & ""
        Alert_Table_MaxInvoices = Trim(rsCheckOnListIdInAlerts!Alert_total_invoices) & ""
        Alert_Table_MaxDollor = Trim(rsCheckOnListIdInAlerts!Alert_total_balance) & ""
        Alert_Table_StartLevel = Trim(rsCheckOnListIdInAlerts!Alert_importance_type) & ""
        
        funIsCompanyInAlerts = True
    End If
    
    Set rsCheckOnListIdInAlerts = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Is Company in Alerts?")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsCheckOnListIdInAlerts = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function



Sub prcGrabAlertSettings()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsAlertSettings As New ADODB.Recordset
    
On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select *  from qbx_alert_settings  "
        
    Set rsAlertSettings = cmdCommand.Execute
    
    If Not rsAlertSettings.EOF Then
        rsAlertSettings.MoveFirst
        
        Alert_MaxDollor = Trim(rsAlertSettings!alert_setting_max_dollar)
        Alert_MaxInvoices = Trim(rsAlertSettings!alert_setting_max_invoices)
        Alert_StartLevel = Trim(rsAlertSettings!alert_setting_start_at_level)
    End If
    
    Set rsAlertSettings = Nothing
    Set cmdCommand = Nothing
    Set parParameter = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            'Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Profile Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsAlertSettings = Nothing
    Set cmdCommand = Nothing
    Set parParameter = Nothing
End Sub

Private Sub Command2_Click()
        prcSaveMessage
End Sub

Sub prcSaveMessage()
    Dim tempDate As String
    Dim temptime As String
    Dim sTmpMsg  As String
    
    Dim iTrack As Integer
    
    tempDate = Trim(Text11.Text)
    temptime = Trim(Text12.Text)

    
On Error Resume Next

    sTmpMsg = Replace(Trim(Text6.Text), "'", "*##*")
    sTmpMsg = Replace(sTmpMsg, """", "$++$")
    
    iTrack = 1
    'keep record of on computer of this message
    frmMain.MessageLog "Tracking ID: (" & iTrack & "), Length = " & Len(sTmpMsg) & vbNewLine & "Msg: " & sTmpMsg
    
    If IsDate(tempDate) Or tempDate = "" Then
        iTrack = iTrack + 3
        If IsDate(temptime) Or Trim(temptime) = "" Then
            iTrack = iTrack + 5
            'If Trim(sCompany) <> "" Then
                'tempDate = CDate(tempDate)
                If temptime <> "" Then
                    temptime = CDate(temptime)
                End If
                'MsgBox Len(sTmpMsg)
                sTmpMsg = Left(sTmpMsg, 4950)
                Dim stemparray(8) As String
                stemparray(0) = Replace(CustomerMainInfo(1), "'", "")
                iTrack = iTrack + 7
                If CustomerMainInfo(6) = "0.00" Then
                    iTrack = iTrack + 11
                    stemparray(1) = 0
                Else
                    iTrack = iTrack + 17
                    stemparray(1) = CustomerMainInfo(6)
                End If
                
                stemparray(2) = CustomerMainInfo(9)
                stemparray(3) = tempDate
                stemparray(4) = temptime
                stemparray(5) = sTmpMsg
                stemparray(6) = ListID
                stemparray(7) = Billing(5)
                
                iTrack = iTrack + 31
                
                prcInsertNote stemparray
                
                iTrack = iTrack + 51
                
                If Trim(Text11.Text) <> "" Then
                    iTrack = iTrack + 83
                    prcPullLastNoteID
                    prcInsertCallbackRecord
                    iTrack = iTrack + 101
                End If
                
                iTrack = iTrack + 117
                
                prcGrabNote
                Command3.Value = True
            'Else
            '    MsgBox "I can't find a company name to associate with this callback note.  Please make sure that you have double clicked one of the companies above."
            'End If
        Else
            MsgBox "The Callback time is not set correctly.  Format the time as: '4:00 PM' and try again."
        End If
    Else
        MsgBox "Please provide a date in the following format: '04/31/2004' and try again."
    End If
    
    
    frmMain.MessageLog "Tracking ID: (" & iTrack & "), Length = " & Len(sTmpMsg) & vbNewLine & "Msg: " & sTmpMsg
End Sub

Sub prcInsertCallbackRecord()
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
    cmdCommand.CommandText = " insert qb_callback_dates " & _
                            " (note_index, callback_date, callback_time, callback_created_by, callback_active) " & _
                            " values " & _
                            " ('" & lNoteClickedOnIndex & "', '" & Trim(Text11.Text) & "', '" & Trim(Text12.Text) & "', '" & sUser & "', '1') "
            
    cmdCommand.Execute
    
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

Sub prcPullLastNoteID()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsGrabLastUserNoteId As New ADODB.Recordset

    lNoteClickedOnIndex = 0
On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    If rsGrabLastUserNoteId.State = 1 Then
        Set rsGrabLastUserNoteId = Nothing
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select top 1 note_index from qb_note where note_created_by = '" & sUser & "' order by note_datestamp desc "
            
    Set rsGrabLastUserNoteId = cmdCommand.Execute
    
    If Not rsGrabLastUserNoteId.EOF Then
        rsGrabLastUserNoteId.MoveFirst
        lNoteClickedOnIndex = Trim(rsGrabLastUserNoteId!note_index) & ""
    End If
    
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
    Set rsGrabLastUserNoteId = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Private Sub Command3_Click()
    Text5.Text = Now
    Text6.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Command2.Enabled = True
    Text6.SetFocus
End Sub

Private Sub Command4_Click()
    frmMain.prcShowFrmPrintPage
End Sub



Private Sub Command6_Click()
    prcUpdateFilterList
    If Command6.Caption = "-->" Then
        Me.Width = 14880
        Command6.Caption = "<--"
    ElseIf Command6.Caption = "<--" Then
        Me.Width = 12870
        Command6.Caption = "-->"
    End If
End Sub


Sub prcDisableForm()
    frmInvoiceQry.Enabled = False
End Sub


Sub prcEnableForm()
    frmInvoiceQry.Enabled = True
End Sub


Sub prcClearForm()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    'Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text13.Text = ""
    Combo1.Text = ""
    Option3.Value = False
    Option4.Value = False
    Label10.Caption = "0.00"
    Label12.Caption = 0
    Frame1.Enabled = False
    Frame2.Enabled = False
    Picture2.Enabled = False
End Sub

Sub prcClearMessageForm()
    Text5.Text = Now
    Text6.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Check7.Visible = False
End Sub




Private Sub Form_Load()
    Dim iMgrResponse As Integer
    
    'frmSupport.Show
    
    tTimeZone.Interval = 60000
    tTimeZone.Enabled = True
    
    TDBGrid2.Visible = True
    List1.Visible = False
    List2.Visible = False
    TDBGrid1.FetchRowStyle = True
    TDBGrid2.FetchRowStyle = True
    Option5.Value = True
    Check7.Visible = False
    
    sOrderBy = " cust_fullname asc "
    sFrmInvoiceQry = 1
    
    Command2.Caption = "Save Message"
    
    frmMain.prcUpdateFormSec 0
    
    If iMgrResponse = 0 Then

        Me.Width = 12870
        Me.Height = 9615
        'Me.Height = 11730
        
        prcGrabPartRemarks Combo7
        
    On Error GoTo ErEnd:
    
        Frame3.Enabled = False
        Frame2.Enabled = False
        Frame1.Enabled = False
        
        Dim i As Integer
        
        Combo1.Clear
        For i = 0 To UBound(aryGImportLvl) - 1
            Combo1.AddItem aryGImportLvl(i, 2)
        Next i
        
        Combo8.Clear
        For i = 0 To UBound(aryGImportLvl) - 1
            Combo8.AddItem aryGImportLvl(i, 2)
        Next i
        
        Combo9.Clear
        Combo9.Text = "All"
        For i = 0 To UBound(aryGRegions) - 1
            Combo9.AddItem aryGRegions(i, 0)
        Next i
        
        prcGrabCustomerInfo sOrderBy
    
    Else
        GoTo ErEnd:
    End If
    
    
    Exit Sub
    
ErEnd:
    If iMgrResponse = 1 Then
        MsgBox "connection error!!"
    End If
    Unload Me
End Sub
Sub prcGrabPartRemarks(Output As ComboBox)
    Dim Remarks As Sql_Results_Struct
    
    Remarks.Query = " select * from qbx_remarks where remark_delete = '0' order by remark_created_date desc "
    SQL_Query_auto Remarks.Query, Remarks.Data
        
    If Not Remarks.Data.EOF Then
        Remarks.Data.MoveFirst
        Output.Text = Trim(Remarks.Data!remark_index) & ":" & Trim(Left(Remarks.Data!remark_msg, 15)) & ""
        Output.AddItem ""
        Output.AddItem "Date:" & Now
        While Not Remarks.Data.EOF
            Output.AddItem Trim(Remarks.Data!remark_index) & ":" & Trim(Left(Remarks.Data!remark_msg, 15)) & ""
            Remarks.Data.MoveNext
        Wend
    End If
    
    SQL_Close_Clear Remarks.Data
End Sub

Sub prcGrabPartRemarks_old()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    Combo7.Text = ""
    Combo7.Clear
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    If rsGrabPartRemarks.State = 1 Then
        Set rsGrabPartRemarks = Nothing
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_remarks where remark_delete = '0' order by remark_created_date desc "
            
    Set rsGrabPartRemarks = cmdCommand.Execute
    
    If Not rsGrabPartRemarks.EOF Then
        rsGrabPartRemarks.MoveFirst
        Combo7.Text = Trim(rsGrabPartRemarks!remark_index) & ":" & Trim(Left(rsGrabPartRemarks!remark_msg, 15)) & ""
        Combo7.AddItem ""
        Combo7.AddItem "Date:" & Now
        While Not rsGrabPartRemarks.EOF
            Combo7.AddItem Trim(rsGrabPartRemarks!remark_index) & ":" & Trim(Left(rsGrabPartRemarks!remark_msg, 15)) & ""
            rsGrabPartRemarks.MoveNext
        Wend
    End If
    
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
    Combo7.Text = ""
    Combo7.Clear
    Set rsGrabPartRemarks = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub



Sub prcLoadImportanceCmb()
    Combo1.AddItem "Balance"
    Combo1.AddItem "Paid"
    Combo1.AddItem "Problem"
End Sub


Private Sub Form_Resize()
On Error Resume Next

    With Me
        If .WindowState <> vbMinimized Then
                    
            If sFrmQueryFilter = 1 Then
                frmQueryFilter.top = Me.top
                frmQueryFilter.Left = Me.Left + Me.Width
            End If
                
        Else
            If sFrmQueryFilter = 1 Then
                Unload frmQueryFilter
            End If
        End If
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmInvoiceQry = 0
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        Option4.Value = False
        prcUpdateUpFront "True"
    Else
        Option4.Value = True
    End If
End Sub

Private Sub Option4_Click()
    If Option4.Value = True Then
        Option3.Value = False
        prcUpdateUpFront "False"
    Else
        Option3.Value = True
    End If
End Sub

Sub prcUpdateUpFront(sUpfront As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "update_upfront_sp"
    
    'listid
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 100, Trim(ListID) & "")
    cmdCommand.Parameters.Append parParameter
    
    'supfront
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 10, Trim(sUpfront) & "")
    cmdCommand.Parameters.Append parParameter
    
    cmdCommand.Execute
        
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            'prcLog Now & "--prcInsertUserDrives:connected(" & frmMonitor.frmmain.cnmc.State & "), err:" & Err.Description
            'List1.AddItem "unable to prcInsertUserDrives for " & sUser
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "updating Importance Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    
End Sub

Private Sub picCal1_Click()
    iCalendarRequest = 1
    frmCalendar.Show
End Sub

Private Sub picCal2_Click()
    iCalendarRequest = 2
    frmCalendar.Show
End Sub





Private Sub Picture1_Click()
    frmMain.prcShowFrmCallbacks
End Sub

Private Sub Picture2_Click()
    sEmailAttachmentMessage = ""
    frmMain.prcShowFrmEmailCustomer
End Sub

Private Sub Picture3_Click()
    Dim Response
    Response = MsgBox("Are you sure, you would like to send an email alert about customer: " & Trim(Text2.Text) & "?", vbExclamation + vbYesNo, "Note Record Opening Error")
    If Response = vbYes Then
        prcCustomerAlert
    Else
        Exit Sub
    End If
End Sub

'alert
Sub prcCustomerAlert()
    Dim sMsg As String
    
    sMsg = "This is a collection's alert about '" & Trim(Text2.Text) & "'. <br><br>" & _
        "Customer Info...<br>" & "<br>" & _
        "Name: " & Trim(Text2.Text) & "<br>" & _
        "Contact: " & Trim(Text3.Text) & "<br>" & _
        "Phone: " & Trim(Text4.Text) & "<br>" & _
        "Fax: " & Trim(Text9.Text) & "<br>" & _
        "Email: " & Trim(Text10.Text) & "<br>" & _
        "Address: " & Trim(Text8.Text) & "<br><br><br>" & _
        "This message is from: " & sUser & "<br>" & _
        "Date: " & Now
        
    SendOCEmail "mfishman@modernconsumer.com; sleavy@modernconsumer.com; icedeno@modernconsumer.com; njaureguy@modernconsumer.com", sUser, sMsg, "", "RE: Collection Alert - " & Trim(Text2.Text) & ""
End Sub



Sub prcRefresh()
    If bFocus = False Then
        prcClearCustomerDtls
        TDBGrid2.ReBind
        TDBGrid3.ReBind
        prcGrabCustomerInfo sOrderBy
        prcCheck4ListID
        prcCallCustomerDtls
        prcUpdateFilterList
        prcGrabPartRemarks Combo7
        'prcClearMessageForm
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        Text6.SetFocus
    End If
    
End Sub

Sub prcUpdateFilterList()
    Check1.Value = sProfileAttrDtlsAry(1, 6)
    Check2.Value = sProfileAttrDtlsAry(1, 10)
    Check3.Value = sProfileAttrDtlsAry(1, 11)
    
    Check4.Value = sProfileAttrDtlsAry(1, 46)
    Check5.Value = sProfileAttrDtlsAry(1, 44)
    Check6.Value = sProfileAttrDtlsAry(1, 45)
End Sub

Sub prcCheck4ListID()
    Dim bFound As Boolean
    Dim sTempListID As String
    Dim iCheckListID As Integer
    bFound = False
    
    
    If rsGrabCustomerInfo.RecordCount > 0 Then
        rsGrabCustomerInfo.MoveFirst
        sTempListID = rsGrabCustomerInfo!cust_listid
        While Not rsGrabCustomerInfo.EOF
            If Trim(rsGrabCustomerInfo!cust_listid) = ListID Then
            'Debug.Print rsGrabCustomerInfo!cust_fullname
                bFound = True
                
                rsGrabCustomerInfo.MoveLast
            End If
            rsGrabCustomerInfo.MoveNext
        Wend
    End If
    
    If bFound = False And ListID = "" Then
        ListID = sTempListID
        TDBGrid1.MoveFirst
    End If
    
End Sub


Sub prcOpenRecordSet()
    Dim strSQL As String
    Dim sMax As String
    
    prcEmptyInvRecords
    
    ' Open Employees Table with a cursor that allows updates
    Set rsGrabInvoiceComp = New ADODB.Recordset
    strSQL = "qbx_inv_temp_record"
    
    sMax = 1000
    SQL_ReConnect_old frmMain.cnMC
    While rsGrabInvoiceComp.State <> 1 Or sMax = 999
        rsGrabInvoiceComp.Open strSQL, frmMain.cnMC, adOpenKeyset, adLockOptimistic, adCmdTable
        sMax = sMax + 1
    Wend
    
    'CreateNetworkFile "z:\public\", sUser & "-log.txt", "sMax: " & sMax & ", rsGrabInvoiceComp.State: (" & rsGrabInvoiceComp.State & ")"
    
    prcGrabInvoiceInfoBackup ListID
    If rsGrabInvoiceComp.State = 1 Then
        prcGrabInvoiceInfo ListID
    End If
    
    prcGrabCustomerDtls ListID
    
    
    'rsGrabInvoiceComp.Close
    'Set rsGrabInvoiceComp = Nothing
End Sub

Sub prcGrabCustomerDtls(sListID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_cust where cust_listid = '" & sListID & "' "
        
    Set rsGrabCustomerDtls = cmdCommand.Execute
    
    If rsGrabCustomerDtls.RecordCount = 1 Then
        rsGrabCustomerDtls.MoveFirst
        prcFillCustomerDtls
    End If
    
    Set rsGrabCustomerDtls = Nothing
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




Function funGetUserWebStatus(sID As String) As String
    Dim lRow As Long, lCol As Long
    Dim sQuery As String, WebStatus As String
    
    Dim i As Integer
            
    sQuery = " SELECT * FROM dealers WHERE ID='" & sID & "' "
            
    'sQuery = " SELECT * FROM leads.dloans_20" & sQuery & " WHERE export = '" & stID & "' " & _
        '  " and datestamp >= '" & Format(dtFrom, "yyyy-mm-dd") & "' and datestamp <= '" & Format(dtTo, "yyyy-mm-dd") & "' "
            
On Error GoTo errs:
    
    Dim strConnect As String
    
    strConnect = "driver={MySQL ODBC 3.51 Driver};server=comet.bluegravity.com;uid=root;pwd=ato-1224f;database=autoloan;"
    Set adoDataConn = New ADODB.Connection
    
    adoDataConn.CursorLocation = adUseClient
    adoDataConn.Open strConnect

    Set rsRecordSet = New ADODB.Recordset
    rsRecordSet.CursorType = adOpenStatic
    rsRecordSet.CursorLocation = adUseClient
    rsRecordSet.LockType = adLockPessimistic
        
    rsRecordSet.source = sQuery
    
    rsRecordSet.ActiveConnection = adoDataConn
    rsRecordSet.Open

    If Not rsRecordSet.EOF Then
        Dim sStatus As String
        rsRecordSet.MoveFirst
        'funConvertStatus (stemp2)
        WebStatus = funConvertStatus(Trim(rsRecordSet!DSTATUS) & "")
        Combo6.Text = WebStatus
        
        While Not rsRecordSet.EOF
            Combo6.AddItem funConvertStatus(Trim(rsRecordSet!DSTATUS)) & ""
            rsRecordSet.MoveNext
        Wend
        funGetUserWebStatus = WebStatus
    Else
        funGetUserWebStatus = ""
    End If
    
    If WebStatus <> "" And strListID <> "" Then
        prcInsertWebStatusIntoCustomer WebStatus
    End If
    
    adoDataConn.Close
    Set adoDataConn = Nothing
    
    Exit Function
    
errs:
    MsgBox "funGetUserWebStatus Error: " & Err.Number & vbNewLine & Err.Description
    
    Resume Next
    
End Function

Sub prcFillCustomerDtls()
    Dim sFilename As String
    
    Combo2.Clear
    Combo2.Text = ""
    Combo3.Clear
    Combo3.Text = ""
    Billing(0) = ""
    Billing(1) = ""
    Billing(2) = ""
    Billing(3) = ""
    Billing(4) = ""
    Billing(5) = ""
    Billing(6) = ""
    Billing(7) = ""
    
    Shipping(0) = ""
    Shipping(1) = ""
    Shipping(2) = ""
    Shipping(3) = ""
    Shipping(4) = ""
    Shipping(5) = ""
    Shipping(6) = ""
    Shipping(7) = ""
    
    sShipping = ""
    sBilling = ""
    
    Option3.Value = False
    Option4.Value = False
        
    'accountnumber
    sAccountNumber = ""
    sAccountNumber = Trim(rsGrabCustomerDtls!cust_accountnumber)
    sAccountNumber = Replace(sAccountNumber, "/", "-")
    
    strListID = ""
    strListID = Trim(rsGrabCustomerDtls!cust_listid)
    
    If sAccountNumber <> "" Then
        funGetUserWebStatus sAccountNumber
        'sFilename = sGLink_Cmc & sAccountNumber
        'sFilename = "http://local/test.com/cmc.php?id=" & sAccountNumber
        'wb3.navigate sFilename
    Else
        sFilename = sGHtml_Dealer_Status
        'sFilename = "z:\qb\collections\html\dealerstatusempty.html"
        wb3.navigate sFilename
    End If
    
    'listid
    CustomerMainInfo(0) = Trim(rsGrabCustomerDtls!cust_listid)
    Text1.Text = CustomerMainInfo(0)
    
    Current_StartLevel = Trim(rsGrabCustomerDtls!importance_type)
    
    'Importance_name
    If Not IsNull(Trim(rsGrabCustomerDtls!importance_name)) Then
        Combo1.Text = Trim(rsGrabCustomerDtls!importance_name)
            
        If Not IsNull(Trim(rsGrabCustomerDtls!importance_upfront)) Then
            If Trim(rsGrabCustomerDtls!importance_upfront) = "True" Then
                Option3.Value = True
                Option4.Value = False
            ElseIf Trim(rsGrabCustomerDtls!importance_upfront) = "False" Then
                Option3.Value = False
                Option4.Value = True
            End If
        End If
    Else
        prcUpdateImportance 0, "", "", ""
    End If
    
    'fullname
    CustomerMainInfo(1) = Trim(rsGrabCustomerDtls!cust_fullname)
    Text2.Text = CustomerMainInfo(1)
    
    CurrCust.CompanyFullName = Trim(rsGrabCustomerDtls!cust_fullname)
    
    'contact
    CustomerMainInfo(2) = Trim(rsGrabCustomerDtls!cust_salutation) & " " & Trim(rsGrabCustomerDtls!cust_firstname) & " " & Trim(rsGrabCustomerDtls!cust_lastname)
    If Trim(CustomerMainInfo(2)) = "" Then
        CustomerMainInfo(2) = Trim(rsGrabCustomerDtls!cust_contact)
    End If
    If Trim(rsGrabCustomerDtls!cust_altcontact) <> "" Then
        Combo2.Text = Trim(rsGrabCustomerDtls!cust_altcontact)
        Combo2.AddItem Trim(rsGrabCustomerDtls!cust_altcontact)
    Else
        Combo2.Text = Trim(CustomerMainInfo(2))
    End If
    Text3.Text = Trim(CustomerMainInfo(2))
    Combo2.AddItem Trim(CustomerMainInfo(2))
    
    'phone
    CustomerMainInfo(3) = Trim(rsGrabCustomerDtls!cust_phone1)
    Text4.Text = CustomerMainInfo(3)
    If Trim(rsGrabCustomerDtls!cust_phone2) <> "" Then
        Combo3.Text = Trim(rsGrabCustomerDtls!cust_phone2)
        Combo3.AddItem Trim(rsGrabCustomerDtls!cust_phone2)
    Else
        Combo3.Text = Trim(CustomerMainInfo(3))
    End If
    Combo3.AddItem Trim(CustomerMainInfo(3))
    
    'fax
    CustomerMainInfo(4) = Trim(rsGrabCustomerDtls!cust_fax1)
    Text9.Text = CustomerMainInfo(4)
    
    'email
    CustomerMainInfo(5) = Trim(rsGrabCustomerDtls!cust_email1)
    Text10.Text = CustomerMainInfo(5)
    
                
    CustomerMainInfo(6) = Trim(rsGrabCustomerDtls!cust_Balance)
    'frmInvoiceQry.Label10.Caption = CustomerMainInfo(6)
    CurrCust.Balance = Trim(rsGrabCustomerDtls!cust_Balance)
                
    CurrCust.TotalBalance = Trim(rsGrabCustomerDtls!cust_totalbalance_money)
                
    CustomerMainInfo(9) = Trim(rsGrabCustomerDtls!cust_jobstatus)
    
    'rep
    Text13.Text = Trim(rsGrabCustomerDtls!cust_salesrepref_fullname)
    
    If Trim(rsGrabCustomerDtls!cust_billaddress_add1) <> "" Then
        Billing(0) = Trim(rsGrabCustomerDtls!cust_billaddress_add1)
        sBilling = Billing(0)
    End If
    If Trim(rsGrabCustomerDtls!cust_billaddress_add2) <> "" Then
        Billing(1) = Trim(rsGrabCustomerDtls!cust_billaddress_add2)
        sBilling = sBilling & vbNewLine & Billing(1)
    End If
    If Trim(rsGrabCustomerDtls!cust_billaddress_add3) <> "" Then
        Billing(2) = Trim(rsGrabCustomerDtls!cust_billaddress_add3)
        sBilling = sBilling & vbNewLine & Billing(2)
    End If
    If Trim(rsGrabCustomerDtls!cust_billaddress_add4) <> "" Then
        Billing(3) = Trim(rsGrabCustomerDtls!cust_billaddress_add4)
        sBilling = sBilling & vbNewLine & Billing(3)
    End If
    If Trim(rsGrabCustomerDtls!cust_billcity) <> "" Then
        Billing(4) = Trim(rsGrabCustomerDtls!cust_billcity)
        sBilling = sBilling & vbNewLine & Billing(4)
    End If
    If Trim(rsGrabCustomerDtls!cust_billstate) <> "" Then
        Billing(5) = Trim(rsGrabCustomerDtls!cust_billstate)
        sBilling = sBilling & ", " & Billing(5)
    End If
    If Trim(rsGrabCustomerDtls!cust_billpostalcode) <> "" Then
        Billing(6) = Trim(rsGrabCustomerDtls!cust_billpostalcode)
        sBilling = sBilling & " " & Billing(6)
    End If
    If Trim(rsGrabCustomerDtls!cust_billcountry) <> "" Then
        Billing(7) = Trim(rsGrabCustomerDtls!cust_billcountry)
        sBilling = sBilling & vbNewLine & Billing(7)
    End If
    
    Text8.Text = Trim(sBilling)

    
    If Trim(rsGrabCustomerDtls!cust_shipaddress_add1) <> "" Then
        Shipping(0) = Trim(rsGrabCustomerDtls!cust_shipaddress_add1)
        sShipping = Shipping(0)
    End If
    If Trim(rsGrabCustomerDtls!cust_shipaddress_add2) <> "" Then
        Shipping(1) = Trim(rsGrabCustomerDtls!cust_shipaddress_add2)
        sShipping = sShipping & vbNewLine & Shipping(1)
    End If
    If Trim(rsGrabCustomerDtls!cust_shipaddress_add3) <> "" Then
        Shipping(2) = Trim(rsGrabCustomerDtls!cust_shipaddress_add3)
        sShipping = sShipping & vbNewLine & Shipping(2)
    End If
    If Trim(rsGrabCustomerDtls!cust_shipaddress_add4) <> "" Then
        Shipping(3) = Trim(rsGrabCustomerDtls!cust_shipaddress_add4)
        sShipping = sShipping & vbNewLine & Shipping(3)
    End If
    If Trim(rsGrabCustomerDtls!cust_shipcity) <> "" Then
        Shipping(4) = Trim(rsGrabCustomerDtls!cust_shipcity)
        sShipping = sShipping & vbNewLine & Shipping(4)
    End If
    If Trim(rsGrabCustomerDtls!cust_shipstate) <> "" Then
        Shipping(5) = Trim(rsGrabCustomerDtls!cust_shipstate)
        sShipping = sShipping & ", " & Shipping(5)
    End If
    If Trim(rsGrabCustomerDtls!cust_shippostalcode) <> "" Then
        Shipping(6) = Trim(rsGrabCustomerDtls!cust_shippostalcode)
        sShipping = sShipping & " " & Shipping(6)
    End If
    If Trim(rsGrabCustomerDtls!cust_shipcountry) <> "" Then
        Shipping(7) = Trim(rsGrabCustomerDtls!cust_shipcountry)
        sShipping = sShipping & vbNewLine & Shipping(7)
    End If
    
End Sub



Sub prcClearCustomerDtls()
    Set rsGrabInvoiceComp = Nothing
    Set rsNotes = Nothing
    Combo1.Text = ""
    Option3.Value = False
    Option4.Value = False
    
    sEmailAttachmentMessage = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    'Text5.Text = ""
    'Text6.Text = ""
    Text9.Text = ""
    Text8.Text = ""
    Text10.Text = ""
    'Text11.Text = ""
    'Text12.Text = ""
    Text13.Text = ""
End Sub


Sub prcEmptyInvRecords()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " truncate table qbx_inv_temp_record "
        
    cmdCommand.Execute
    
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


Sub prcConvertGuiVals_Test()

    If Option6.Value = True Then
        CustList.UpFront = True
    ElseIf Option7.Value = True Then
        CustList.UpFront = False
    End If
    
    CustList.Importance = Trim(Combo8.Text)
    
    CustList.Region = funZoneTime(Trim(Combo9.Text))

End Sub
Sub prcGrabCust_List_Test()
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    'convert gui to vars
    prcConvertGuiVals_Test
    'create customer query structure
    Dim CList As Sql_Results_Struct
    'build query
    CList.Query = QryBuildCustList
    'execute query
    SQL_Query_auto CList.Query, CList.Data
    
    
On Error GoTo errHandle:

    'regui-fi records
    TDBGrid1.ReBind
    Label18.Caption = CList.Data.RecordCount
    Frame3.Enabled = True
    
    Exit Sub
    
errHandle:
    MsgBox Err.Number & " " & Err.source & vbNewLine & Err.Description
End Sub

Sub prcGrabCustomerInfo(sOrderBy As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String
    
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    'bSortOrd = True
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    If SQL_ReConnect_old(frmMain.cnMC) = False Then Exit Sub
    'If frmMain.cnMC.State <> 1 Then
        'Exit Sub
    'End If
    Set rsGrabCustomerInfo = Nothing
    
    If sProfileAttrDtlsAry(1, 6) = 0 Then
        strXtraSQL = " where cust_jobstatus <> 'awarded' and cust_accountnumber <> '' "
    End If
    
    If sProfileAttrDtlsAry(1, 10) = 0 Then
        If sProfileAttrDtlsAry(1, 11) = 0 Then
            If strXtraSQL = "" Then
                strXtraSQL = " where CONVERT(int, cust_totalbalance_money) = '0' and  CONVERT(int, cust_totalbalance_money) <> '0' "
            Else
                strXtraSQL = strXtraSQL & " and CONVERT(int, cust_totalbalance_money) = '0' and  CONVERT(int, cust_totalbalance_money) <> '0' "
            End If
        Else
            If strXtraSQL = "" Then
                strXtraSQL = " where CONVERT(int, cust_totalbalance_money) = '0' "
            Else
                strXtraSQL = strXtraSQL & " and CONVERT(int, cust_totalbalance_money) = '0' "
            End If
        End If
    Else
        If sProfileAttrDtlsAry(1, 11) = 0 Then
            If strXtraSQL = "" Then
                strXtraSQL = " where CONVERT(int, cust_totalbalance_money) <> '0' "
            Else
                strXtraSQL = strXtraSQL & " and CONVERT(int, cust_totalbalance_money) <> '0' "
            End If
        End If
    End If
    
    If sProfileAttrDtlsAry(1, 44) = 0 Then
        If strXtraSQL = "" Then
            strXtraSQL = " where sign(cust_totalbalance_money) != CONVERT(money, '-1') "
        Else
            strXtraSQL = strXtraSQL & " and sign(cust_totalbalance_money) != CONVERT(money, '-1') "
        End If
    End If
    
    If Option6.Value = True Then
        If strXtraSQL = "" Then
            strXtraSQL = " where importance_upfront = 'True' "
        Else
            strXtraSQL = strXtraSQL & " and importance_upfront = 'True' "
        End If
    ElseIf Option7.Value = True Then
        If strXtraSQL = "" Then
            strXtraSQL = " where importance_upfront = 'False' "
        Else
            strXtraSQL = strXtraSQL & " and importance_upfront = 'False' "
        End If
    End If
    
    If Trim(Combo8.Text) <> "" Then
        If strXtraSQL = "" Then
            strXtraSQL = " where importance_name = '" & Trim(Combo8.Text) & "' "
        Else
            strXtraSQL = strXtraSQL & " and importance_name = '" & Trim(Combo8.Text) & "' "
        End If
    End If
    
    If Trim(Combo9.Text) <> "All" And Trim(Combo9.Text) <> "" Then
        If strXtraSQL = "" Then
            strXtraSQL = " where " & funStates(Trim(Combo9.Text))
        Else
            strXtraSQL = strXtraSQL & " and " & funStates(Trim(Combo9.Text))
        End If
    End If
    
    funZoneTime Trim(Combo9.Text)
    
    
    If sOrderBy = "" Then
        sOrderBy = " cust_fullname asc "
    End If
    
    If strXtraSQL <> "" Then
        strXtraSQL = strXtraSQL & " and "
    Else
        strXtraSQL = " where "
    End If
    
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    
    cmdCommand.CommandText = " select cust_listid, cust_accountnumber, " & _
                " cust_fullname, cust_totalbalance_money, cust_jobstatus, " & _
                " cust_salesrepref_fullname, importance_name, importance_upfront, " & _
                " cust_accountnumber_numeric " & _
                " from qbx_cust " & _
                strXtraSQL & " cust_fullname <> 'Customer' order by " & sOrderBy & " "
    
    If Trim(Text7.Text) = "" Then
        Text17.Text = cmdCommand.CommandText
    Else
        Text17.Text = Trim(Text17.Text) & vbNewLine & vbNewLine & cmdCommand.CommandText
    End If
    
    'frmSupport.Text6.Text = cmdCommand.CommandText
        
    Set rsGrabCustomerInfo = cmdCommand.Execute
    
    
    If rsGrabCustomerInfo.RecordCount <> 0 Then
    
        
        With rsGrabCustomerInfo
            If Not .BOF Or Not .EOF Then
                TDBGrid1.ApproxCount = .RecordCount
                
                TDBGrid1.Refresh
                'Text3.Text = .RecordCount
            End If
        End With
        
        'Frame2.Enabled = True
        'Frame3.Enabled = True
    End If
    TDBGrid1.ReBind
    Label18.Caption = rsGrabCustomerInfo.RecordCount
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Frame3.Enabled = True
    
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

Function funStates(sZone As String) As String

On Error GoTo ZoneErr:
    Select Case sZone
        Case "Western Time":    funStates = " ( cust_billstate = 'CA' or cust_billstate = 'WA' " & _
                                    " or cust_billstate = 'NV' or cust_billstate = 'OR' ) "
                                    
        Case "Mountain Time":    funStates = " ( cust_billstate = 'MT' or cust_billstate = 'ID' or cust_billstate = 'WY' " & _
                                    " or cust_billstate = 'UT' or cust_billstate = 'CO' " & _
                                    " or cust_billstate = 'AZ' or cust_billstate = 'NM' ) "
                                    
                                    
        Case "Central Time":    funStates = " ( cust_billstate = 'ND' or cust_billstate = 'SD' or cust_billstate = 'NE' " & _
                                    " or cust_billstate = 'KS' or cust_billstate = 'OK' " & _
                                    " or cust_billstate = 'TX' or cust_billstate = 'MN' " & _
                                    " or cust_billstate = 'IA' or cust_billstate = 'MO' " & _
                                    " or cust_billstate = 'AR' or cust_billstate = 'LA' " & _
                                    " or cust_billstate = 'WI' or cust_billstate = 'IL' " & _
                                    " or cust_billstate = 'TN' or cust_billstate = 'MS' " & _
                                    " or cust_billstate = 'AL' ) "
        
        Case "Eastern Time":    funStates = " ( cust_billstate = 'MI' or cust_billstate = 'IN' or cust_billstate = 'KY' " & _
                                    " or cust_billstate = 'OH' or cust_billstate = 'WY' or cust_billstate = 'ME' " & _
                                    " or cust_billstate = 'NH' or cust_billstate = 'VT' or cust_billstate = 'MA' " & _
                                    " or cust_billstate = 'RI' or cust_billstate = 'CT' or cust_billstate = 'NY' " & _
                                    " or cust_billstate = 'PA' or cust_billstate = 'NJ' or cust_billstate = 'DE' " & _
                                    " or cust_billstate = 'MD' or cust_billstate = 'VA' or cust_billstate = 'NC' " & _
                                    " or cust_billstate = 'SC' or cust_billstate = 'GA' or cust_billstate = 'FL' ) "
    End Select
    
    Exit Function
    
ZoneErr:
    Resume Next
    
End Function

Function funZoneTime(sZone As String) As String

    Select Case sZone
        Case "Western Time":    lbRegionTime.Caption = Hour(Time) - 3 & ":" & Minute(Time) & Format(Time, " AM/PM")
        Case "Mountain Time":    lbRegionTime.Caption = Hour(Time) - 2 & ":" & Minute(Time) & Format(Time, " AM/PM")
        Case "Central Time":    lbRegionTime.Caption = Hour(Time) - 1 & ":" & Minute(Time) & Format(Time, " AM/PM")
        Case "Eastern Time":    lbRegionTime.Caption = Format(Time, "hh:mm AM/PM")
        Case "All":    lbRegionTime.Caption = ""
        Case "":    lbRegionTime.Caption = ""
    End Select
    
    funZoneTime = lbRegionTime.Caption
    
End Function

Private Sub TDBGrid1_DblClick()
    
    If SQL_ReConnect_old(frmMain.cnMC) = False Then
        frmMain.StatusBar1.Panels.Item(6).Text = "Not Connected."
        Exit Sub
    End If
    frmMain.StatusBar1.Panels.Item(6).Text = "Connected."
    
    TDBGrid2.Visible = True
    List1.Visible = False
    List2.Visible = False
    Label10.Caption = "0.00"
    
    ListID = Trim(TDBGrid1.Columns(0).Value)
    
    prcDisableModifyNotes
    prcCallCustomerDtls
    prcClearMessageForm
    Frame1.Enabled = True
    Frame2.Enabled = True
    If Trim(Combo3.Text) <> "" Then
        Label3.ForeColor = vbRed
        Label3.ToolTipText = "Click me"
    Else
        Label3.ForeColor = vbBlack
        Label3.ToolTipText = ""
    End If
    
End Sub

Sub prcDisableModifyNotes()

    SSTab1.Tab = 0
    Command11.Enabled = False
    Command12.Enabled = False
    Text15.Enabled = False
    Text16.Enabled = False
    
End Sub

Sub prcEnableModifyNotes()

    SSTab1.Tab = 1
    Command11.Enabled = True
    Command12.Enabled = True
    Text15.Enabled = True
    Text16.Enabled = True
    
End Sub

Sub prcCallCustomerDtls()

    'pulls customer contact details
    prcClearCustomerDtls
    
    'pulls invoices/payments
    prcOpenRecordSet
    
    Label10.Caption = CustomerMainInfo(6)
    'pull sql importance
    'funGrabImportance 1, ListID
                    
                    
    'call back info
    prcGrabNote
End Sub

Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
On Error Resume Next

    Response = 0
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim sTempOrderBy As String
On Error Resume Next

    MousePointer = vbHourglass
    If bSortOrd Then
        CustList.OrderBy = " asc "
        If rsGrabCustomerInfo.Fields(ColIndex).name = "cust_accountnumber" Then
            sTempOrderBy = "cust_accountnumber_numeric ASC "
            CustList.OrderName = " cust_accountnumber "
            'prcGrabCustomerInfo "cust_accountnumber_numeric ASC "
        Else
            sTempOrderBy = rsGrabCustomerInfo.Fields(ColIndex).name & " ASC "
            CustList.OrderName = rsGrabCustomerInfo.Fields(ColIndex).name
            'prcGrabCustomerInfo rsGrabCustomerInfo.Fields(ColIndex).Name & " ASC "
        End If
        'rsGrabCustomerInfo.Sort = "[" & rsGrabCustomerInfo.Fields(ColIndex).Name & "]  ASC"
        bSortOrd = False
    Else
        CustList.OrderBy = " desc "
        If rsGrabCustomerInfo.Fields(ColIndex).name = "cust_accountnumber" Then
            sTempOrderBy = "cust_accountnumber_numeric DESC "
            CustList.OrderName = " cust_accountnumber "
            'prcGrabCustomerInfo "cust_accountnumber_numeric DESC "
        Else
            sTempOrderBy = rsGrabCustomerInfo.Fields(ColIndex).name & " DESC "
            CustList.OrderName = rsGrabCustomerInfo.Fields(ColIndex).name
            'prcGrabCustomerInfo rsGrabCustomerInfo.Fields(ColIndex).Name & " DESC "
        End If
        'rsGrabCustomerInfo.Sort = "[" & rsGrabCustomerInfo.Fields(ColIndex).Name & "]  DESC"
        bSortOrd = True
    End If
    
    prcGrabCustomerInfo sTempOrderBy
    sOrderBy = " " & sTempOrderBy
    
    TDBGrid2.col = ColIndex
    rsGrabCustomerInfo.MoveFirst
    TDBGrid1.ReBind
    MousePointer = vbDefault

End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    
    On Error GoTo errHandle
       
    rsGrabCustomerInfo.Bookmark = Bookmark
    If Left(Trim(rsGrabCustomerInfo!cust_totalbalance_money), 1) <> "-" And Trim(rsGrabCustomerInfo!cust_totalbalance_money) <> "0" Then
        RowStyle = TDBGrid1.Styles(10)
    End If
    
    Exit Sub

errHandle:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Fetch Row Style Error"
    End Select
    Exit Sub
    
    
End Sub


Private Sub TDBGrid1_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    Dim ColIndex As Integer, col As Integer
    Dim RowsFetched As Integer, Row As Long
    Dim StartRow As Variant
    Dim Response As Long
    Dim cols As Long
    Dim Rows As Long
    Dim Pos As Long
    Dim strlen As Integer
    Dim stemp As String

On Error GoTo NoRead
    
    'If bCancelRead Then Exit Sub
    cols = RowBuf.ColumnCount - 1
    Rows = RowBuf.RowCount - 1
    RowsFetched = 0
    
    If IsNull(StartLocation) Then
        If offset < 0 Then
            rsGrabCustomerInfo.MoveLast
            rsGrabCustomerInfo.MoveNext
        Else
            rsGrabCustomerInfo.MoveFirst
            rsGrabCustomerInfo.MovePrevious
        End If
        rsGrabCustomerInfo.Move offset
    Else
        rsGrabCustomerInfo.Move offset, StartLocation
    End If
        
    StartRow = rsGrabCustomerInfo.Bookmark
    Pos = rsGrabCustomerInfo.AbsolutePosition
    
    For Row = 0 To Rows
        If rsGrabCustomerInfo.BOF Or rsGrabCustomerInfo.EOF Then Exit For
        For col = 0 To cols
            stemp = Trim(rsGrabCustomerInfo!importance_upfront) & ""
            If stemp = "" Then
                stemp = "N/A"
            ElseIf LCase(stemp) = "true" Then
                stemp = "Yes"
            ElseIf LCase(stemp) = "false" Then
                stemp = "No"
            End If
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsGrabCustomerInfo!cust_listid) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsGrabCustomerInfo!cust_accountnumber) & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsGrabCustomerInfo!cust_fullname) & ""
                Case (3):   RowBuf.Value(Row, 3) = funFormatDecimal(Trim(rsGrabCustomerInfo!cust_totalbalance_money) & "")
                Case (4):   RowBuf.Value(Row, 4) = Trim(rsGrabCustomerInfo!cust_jobstatus) & ""
                Case (5):   RowBuf.Value(Row, 5) = Trim(rsGrabCustomerInfo!cust_salesrepref_fullname) & ""
                Case (6):   RowBuf.Value(Row, 6) = Trim(rsGrabCustomerInfo!importance_name) & ""
                Case (7):   RowBuf.Value(Row, 7) = stemp
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsGrabCustomerInfo.Bookmark
        RowsFetched = RowsFetched + 1
        rsGrabCustomerInfo.MoveNext
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

Sub prcGrabInvoiceInfo(sListID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    'Dim i As Integer
    Dim strXtraSQL      As String
    Dim dInvDate        As Date
    Dim dDayz           As Double
    Dim sInterval       As String
    Dim sTmpInvoiceNumber As String
    Dim iInvCurrCount   As Integer

On Error GoTo errHandle:

    iInvCurrCount = 0
    
    MousePointer = vbHourglass
    If SQL_ReConnect_old(frmMain.cnMC) = False Then Exit Sub
    'If frmMain.cnMC.State <> 1 Then
        'Exit Sub
    'End If
    
    If rsGrabInvoiceInfo.State = 1 Then
        Set rsGrabInvoiceInfo = Nothing
    End If
    
    'no zero balance
    If sProfileAttrDtlsAry(1, 46) = 0 Then
        strXtraSQL = " and inv_balanceremaining <> '0.00' "
    End If
        
    sInterval = ""
    dDayz = sProfileAttrDtlsAry(1, 12)
    If sProfileAttrDtlsAry(1, 13) = "days" Then
        sInterval = "d"
    ElseIf sProfileAttrDtlsAry(1, 13) = "months" Then
        sInterval = "m"
    ElseIf sProfileAttrDtlsAry(1, 13) = "years" Then
        sInterval = "yyyy"
    End If
    dInvDate = DateAdd(sInterval, dDayz, Date)
    strXtraSQL = strXtraSQL & " and inv_txndate > '" & dInvDate & "' "
    
    'CreateNetworkFile "z:\public\", "veronica-log.txt", "inv-info-5"
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select inv_unique, inv_txnid, inv_txndate, inv_refnumber, inv_subtotal, inv_balanceremaining " & _
                " from qbx_inv " & _
                " where inv_enabled = '1' and " & _
                " inv_customerref_listid = '" & sListID & "' " & strXtraSQL & _
                " order by inv_unique desc "
                                
    'MsgBox cmdCommand.CommandText
            
    Set rsGrabInvoiceInfo = cmdCommand.Execute
    
    'CreateNetworkFile "z:\public\", "veronica-log.txt", "rsGrabInvoiceInfo.EOF: " & rsGrabInvoiceInfo.EOF
    If Not rsGrabInvoiceInfo.EOF Then
    
        Frame1.Enabled = True
        rsGrabInvoiceInfo.MoveFirst
        
        If rsGrabInvoiceInfo.RecordCount > 1 Then
            iInvCurrCount = 1
        End If
        
        sTmpInvoiceNumber = ""
        
        While Not rsGrabInvoiceInfo.EOF
            
                If sTmpInvoiceNumber = "" Then
                    sTmpInvoiceNumber = Trim(rsGrabInvoiceInfo!inv_refnumber)
                        
                    rsGrabInvoiceComp.AddNew
                    rsGrabInvoiceComp!qbx_inv_rs_txndate = Trim(rsGrabInvoiceInfo!inv_txndate)
                    rsGrabInvoiceComp!qbx_inv_rs_ref_num = Trim(rsGrabInvoiceInfo!inv_refnumber)
                    rsGrabInvoiceComp!qbx_inv_rs_amount = Trim(rsGrabInvoiceInfo!inv_subtotal)
                    rsGrabInvoiceComp.Update
                        
                    prcGrabInvoicePay Trim(rsGrabInvoiceInfo!inv_txnid), Trim(rsGrabInvoiceInfo!inv_refnumber)
                                                
                    rsGrabInvoiceComp.AddNew
                    rsGrabInvoiceComp!qbx_inv_rs_txndate = "Balance"
                    rsGrabInvoiceComp!qbx_inv_rs_ref_num = Trim(rsGrabInvoiceInfo!inv_refnumber)
                    rsGrabInvoiceComp!qbx_inv_rs_amount = Trim(rsGrabInvoiceInfo!inv_balanceremaining)
                    rsGrabInvoiceComp.Update
                    
                Else
                    If sTmpInvoiceNumber <> Trim(rsGrabInvoiceInfo!inv_refnumber) Then
                    
                        'Debug.Print rsGrabInvoiceInfo!inv_unique
                        rsGrabInvoiceComp.AddNew
                        rsGrabInvoiceComp!qbx_inv_rs_txndate = Trim(rsGrabInvoiceInfo!inv_txndate)
                        rsGrabInvoiceComp!qbx_inv_rs_ref_num = Trim(rsGrabInvoiceInfo!inv_refnumber)
                        rsGrabInvoiceComp!qbx_inv_rs_amount = Trim(rsGrabInvoiceInfo!inv_subtotal)
                        rsGrabInvoiceComp.Update
                        
                        prcGrabInvoicePay Trim(rsGrabInvoiceInfo!inv_txnid), Trim(rsGrabInvoiceInfo!inv_refnumber)
                                                
                        rsGrabInvoiceComp.AddNew
                        rsGrabInvoiceComp!qbx_inv_rs_txndate = "Balance"
                        rsGrabInvoiceComp!qbx_inv_rs_ref_num = Trim(rsGrabInvoiceInfo!inv_refnumber)
                        rsGrabInvoiceComp!qbx_inv_rs_amount = Trim(rsGrabInvoiceInfo!inv_balanceremaining)
                        rsGrabInvoiceComp.Update
                        
                    End If
                End If
                                
            rsGrabInvoiceInfo.MoveNext
        Wend
        
        ''''''''''catch all payments not referenced by invoices ''''''
        'prcGrabOtherPayments Trim(rsGrabInvoiceInfo!inv_txnid)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Command4.Enabled = True
        'Label10.Caption = Trim(TDBGrid1.Columns(3).Value) & ""
    Else
        Command4.Enabled = False
        Label10.Caption = "0.00"
    End If
    
    'CreateNetworkFile "z:\public\", "veronica-log.txt", "RecordCount" & rsGrabInvoiceInfo.RecordCount
    Label21.Caption = rsGrabInvoiceInfo.RecordCount
    TDBGrid2.ApproxCount = rsGrabInvoiceComp.RecordCount
    TDBGrid2.ReBind
    
    Set rsGrabInvoiceInfo = Nothing
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
    
    Set rsGrabInvoiceInfo = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Sub prcGrabOtherPayments(sInv_txnid As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsOtherPayments As New ADODB.Recordset
    Dim strXtraSQL      As String
    Dim dInvPayDate        As Date
    Dim dDayz           As Double
    Dim sInterval       As String
    Dim l As Long
    Dim bOtherFound As Boolean
    
On Error GoTo errHandle:

    bOtherFound = False
    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    If sProfileAttrDtlsAry(1, 45) = 0 Then
        strXtraSQL = " and sign(inv_pay_amount) != '-1' "
    End If
    
    sInterval = ""
    dDayz = sProfileAttrDtlsAry(1, 15)
    If sProfileAttrDtlsAry(1, 16) = "days" Then
        sInterval = "d"
    ElseIf sProfileAttrDtlsAry(1, 16) = "months" Then
        sInterval = "m"
    ElseIf sProfileAttrDtlsAry(1, 16) = "years" Then
        sInterval = "yyyy"
    End If
    dInvPayDate = DateAdd(sInterval, dDayz, Date)
    strXtraSQL = strXtraSQL & " and inv_pay_txndate > '" & dInvPayDate & "' "
    
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select inv_pay_txndate, inv_pay_txndate, inv_pay_amount, inv_pay_refnumber " & _
                "from qbx_inv_payments where inv_txnid_link = '" & sInv_txnid & "' " & strXtraSQL
                
    'MsgBox cmdCommand.CommandText
    
    Set rsOtherPayments = cmdCommand.Execute
    
    If Not rsOtherPayments.EOF Then
        rsOtherPayments.MoveFirst
        While Not rsOtherPayments.EOF
            For l = 0 To UBound(PaymentsByInvoiceOnly) - 1
                If Trim(rsOtherPayments!inv_pay_refnumber) <> PaymentsByInvoiceOnly(l).inv_pay_refnumber Then
                    PaymentsByInvoiceOnly(l).display_this = True
                    bOtherFound = True
                End If
            Next l
            rsOtherPayments.MoveNext
        Wend
    End If
    
    'display other payments
    If bOtherFound = True Then
        rsGrabInvoiceComp.AddNew
        rsGrabInvoiceComp!qbx_inv_rs_txndate = ""
        rsGrabInvoiceComp!qbx_inv_rs_ref_num = ""
        rsGrabInvoiceComp!qbx_inv_rs_amount = ""
        rsGrabInvoiceComp.Update
        rsGrabInvoiceComp.AddNew
        rsGrabInvoiceComp!qbx_inv_rs_txndate = "---"
        rsGrabInvoiceComp!qbx_inv_rs_ref_num = "Other Payments"
        rsGrabInvoiceComp!qbx_inv_rs_amount = "---"
        rsGrabInvoiceComp.Update
    End If
    For l = 0 To UBound(PaymentsByInvoiceOnly) - 1
        If PaymentsByInvoiceOnly(l).display_this = True Then
            rsGrabInvoiceComp.AddNew
            rsGrabInvoiceComp!qbx_inv_rs_txndate = PaymentsByInvoiceOnly(l).inv_txndate
            rsGrabInvoiceComp!qbx_inv_rs_ref_num = "CK:" & PaymentsByInvoiceOnly(l).inv_pay_refnumber
            rsGrabInvoiceComp!qbx_inv_rs_amount = PaymentsByInvoiceOnly(l).inv_pay_amount
            rsGrabInvoiceComp.Update
        End If
    Next l
    
    Set rsOtherPayments = Nothing
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
    Set rsOtherPayments = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub


Sub CreateNetworkFile(sPath As String, sFilename As String, sMsg As String)
    Dim fs, f, A
    Dim strTemp As String

On Error GoTo errhandler
        
     List2.AddItem sMsg
    'f = FreeFile
    'Open sPath & sFilename For Append As #f
    'Print #f, Now & sMsg
    'f.Close
    
    Exit Sub
    
errhandler:
'f = FreeFile
    'Open sPath & "CreateFileError.txt" For Append As #f
    'Print #f, Now & " Err = " & Err.Description & " " & Err.number
    'f.Close
    
End Sub

Sub prcGrabInvoiceInfoBackup(sListID As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    'Dim i As Integer
    Dim strXtraSQL      As String
    Dim dInvDate        As Date
    Dim dDayz           As Double
    Dim sInterval       As String
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    If rsGrabInvoiceInfo.State = 1 Then
        Set rsGrabInvoiceInfo = Nothing
    End If
    
    'no zero balance
    If sProfileAttrDtlsAry(1, 46) = 0 Then
        strXtraSQL = " and inv_balanceremaining <> '0.00' "
    End If
    
    
    sInterval = ""
    dDayz = sProfileAttrDtlsAry(1, 12)
    If sProfileAttrDtlsAry(1, 13) = "days" Then
        sInterval = "d"
    ElseIf sProfileAttrDtlsAry(1, 13) = "months" Then
        sInterval = "m"
    ElseIf sProfileAttrDtlsAry(1, 13) = "years" Then
        sInterval = "yyyy"
    End If
    dInvDate = DateAdd(sInterval, dDayz, Date)
    strXtraSQL = strXtraSQL & " and inv_txndate > '" & dInvDate & "' "
    
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select inv_txnid, inv_txndate, inv_refnumber, inv_subtotal, inv_balanceremaining " & _
                " from qbx_inv where inv_enabled = '1' and inv_customerref_listid = '" & sListID & "' " & strXtraSQL
            
    List1.Clear
    List1.AddItem "empty"
            
    Set rsGrabInvoiceInfo = cmdCommand.Execute
    
    
    If Not rsGrabInvoiceInfo.EOF Then
        List1.Clear
        List1.AddItem "Date | Account | Totals"
        Frame1.Enabled = True
        rsGrabInvoiceInfo.MoveFirst
        
        While Not rsGrabInvoiceInfo.EOF
        
            List1.AddItem Trim(rsGrabInvoiceInfo!inv_txndate) & " | " & _
                        Trim(rsGrabInvoiceInfo!inv_refnumber) & " | " & _
                        Trim(rsGrabInvoiceInfo!inv_subtotal)
            
            prcGrabInvoicePayBackup Trim(rsGrabInvoiceInfo!inv_txnid), Trim(rsGrabInvoiceInfo!inv_refnumber)
                                    
        
            List1.AddItem Trim(rsGrabInvoiceInfo!inv_txndate) & " | " & _
                        Trim(rsGrabInvoiceInfo!inv_refnumber) & " | " & _
                        Trim(rsGrabInvoiceInfo!inv_balanceremaining)
            
            rsGrabInvoiceInfo.MoveNext
        Wend
        
    End If
    
    Label21.Caption = rsGrabInvoiceInfo.RecordCount
    TDBGrid2.ApproxCount = rsGrabInvoiceComp.RecordCount
    TDBGrid2.ReBind
    
    Set rsGrabInvoiceInfo = Nothing
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
    
    Set rsGrabInvoiceInfo = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Sub prcGrabInvoicePay(sTxnID As String, sInvNum As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String
    Dim dInvPayDate        As Date
    Dim dDayz           As Double
    Dim sInterval       As String
    Dim l As Long
    
On Error GoTo errHandle:

    ReDim PaymentsByInvoiceOnly(1)
    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    If sProfileAttrDtlsAry(1, 45) = 0 Then
        strXtraSQL = " and sign(inv_pay_amount) != '-1' "
    End If
    
    
    sInterval = ""
    dDayz = sProfileAttrDtlsAry(1, 15)
    If sProfileAttrDtlsAry(1, 16) = "days" Then
        sInterval = "d"
    ElseIf sProfileAttrDtlsAry(1, 16) = "months" Then
        sInterval = "m"
    ElseIf sProfileAttrDtlsAry(1, 16) = "years" Then
        sInterval = "yyyy"
    End If
    dInvPayDate = DateAdd(sInterval, dDayz, Date)
    strXtraSQL = strXtraSQL & " and inv_pay_txndate > '" & dInvPayDate & "' "
    
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select inv_pay_txndate, inv_pay_amount, inv_pay_refnumber " & _
                "from qbx_inv_payments where inv_txnid_link = '" & sTxnID & "' " & strXtraSQL
    'MsgBox cmdCommand.CommandText
    Set rsGrabPaymentInfo = cmdCommand.Execute
    
    If Not rsGrabPaymentInfo.EOF Then
    
        ReDim PaymentsByInvoiceOnly(rsGrabPaymentInfo.RecordCount)
    
        Frame1.Enabled = True
        rsGrabPaymentInfo.MoveFirst
        
        While Not rsGrabPaymentInfo.EOF
        
            PaymentsByInvoiceOnly(l).inv_txnid_link = sTxnID
            PaymentsByInvoiceOnly(l).inv_txndate = Trim(rsGrabPaymentInfo!inv_pay_txndate)
            PaymentsByInvoiceOnly(l).inv_pay_amount = Trim(rsGrabPaymentInfo!inv_pay_amount)
            PaymentsByInvoiceOnly(l).inv_pay_refnumber = Trim(rsGrabPaymentInfo!inv_pay_refnumber)
            PaymentsByInvoiceOnly(l).display_this = False
            
            rsGrabInvoiceComp.AddNew
            rsGrabInvoiceComp!qbx_inv_rs_txndate = Trim(rsGrabPaymentInfo!inv_pay_txndate)
            'rsGrabInvoiceComp!qbx_inv_rs_ref_num = Trim(sInvNum) & "(" & Trim(rsGrabPaymentInfo!inv_pay_refnumber) & ")"
            rsGrabInvoiceComp!qbx_inv_rs_ref_num = "CK:" & Trim(rsGrabPaymentInfo!inv_pay_refnumber) & ""
            rsGrabInvoiceComp!qbx_inv_rs_amount = Trim(rsGrabPaymentInfo!inv_pay_amount)
            rsGrabInvoiceComp.Update
            
            rsGrabPaymentInfo.MoveNext
            l = l + 1
            
        Wend
        
    End If
    
    Set rsGrabPaymentInfo = Nothing
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
    Set rsGrabPaymentInfo = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Sub prcGrabInvoicePayBackup(sTxnID As String, sInvNum As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String
    Dim dInvPayDate        As Date
    Dim dDayz           As Double
    Dim sInterval       As String
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    If sProfileAttrDtlsAry(1, 45) = 0 Then
        strXtraSQL = " and sign(inv_pay_amount) != '-1' "
    End If
    
    
    sInterval = ""
    dDayz = sProfileAttrDtlsAry(1, 15)
    
    If sProfileAttrDtlsAry(1, 16) = "days" Then
        sInterval = "d"
    ElseIf sProfileAttrDtlsAry(1, 16) = "months" Then
        sInterval = "m"
    ElseIf sProfileAttrDtlsAry(1, 16) = "years" Then
        sInterval = "yyyy"
    End If
    dInvPayDate = DateAdd(sInterval, dDayz, Date)
    strXtraSQL = strXtraSQL & " and inv_pay_txndate > '" & dInvPayDate & "' "
    
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select inv_pay_txndate, inv_pay_amount, inv_pay_refnumber " & _
                "from qbx_inv_payments where inv_txnid_link = '" & sTxnID & "' " & strXtraSQL
            
    Set rsGrabPaymentInfo = cmdCommand.Execute
    
    If Not rsGrabPaymentInfo.EOF Then
    
        Frame1.Enabled = True
        rsGrabPaymentInfo.MoveFirst
        
        While Not rsGrabPaymentInfo.EOF
        
            List1.AddItem Trim(rsGrabPaymentInfo!inv_pay_txndate) & " | " & _
                        Trim(sInvNum) & " | " & _
                        Trim(rsGrabPaymentInfo!inv_pay_amount) & _
                        Trim(rsGrabPaymentInfo!inv_pay_refnumber)
            
            rsGrabPaymentInfo.MoveNext
            
        Wend
        
    End If
    
    Set rsGrabPaymentInfo = Nothing
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
    Set rsGrabPaymentInfo = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub TDBGrid2_Error(ByVal DataError As Integer, Response As Integer)
On Error Resume Next

    Response = 0
End Sub

Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    
    On Error GoTo errHandle
       
    rsGrabInvoiceComp.Bookmark = Bookmark
    If Left(Trim(rsGrabInvoiceComp!qbx_inv_rs_amount), 1) = "-" Then
        RowStyle = TDBGrid2.Styles(10)
    End If
    If Trim(rsGrabInvoiceComp!qbx_inv_rs_txndate) = "Balance" Then
        RowStyle = TDBGrid2.Styles(11)
    End If
    
    Exit Sub

errHandle:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Invoice Grid Fetch Row Style Error"
    End Select
    Exit Sub
    
    
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
            rsGrabInvoiceComp.MoveLast
            rsGrabInvoiceComp.MoveNext
        Else
            rsGrabInvoiceComp.MoveFirst
            rsGrabInvoiceComp.MovePrevious
        End If
        rsGrabInvoiceComp.Move offset
    Else
        rsGrabInvoiceComp.Move offset, StartLocation
    End If
        
    StartRow = rsGrabInvoiceComp.Bookmark
    Pos = rsGrabInvoiceComp.AbsolutePosition
    
    For Row = 0 To Rows
        If rsGrabInvoiceComp.BOF Or rsGrabInvoiceComp.EOF Then Exit For
        For col = 0 To cols
            'strlen = Len(rsGrabInvoiceComp!Status)
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsGrabInvoiceComp!qbx_inv_rs_txndate) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsGrabInvoiceComp!qbx_inv_rs_ref_num) & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsGrabInvoiceComp!qbx_inv_rs_amount) & ""
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsGrabInvoiceComp.Bookmark
        RowsFetched = RowsFetched + 1
        rsGrabInvoiceComp.MoveNext
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)


    If KeyCode = 13 Then
        
        If SQL_ReConnect_old(frmMain.cnMC) = False Then
            frmMain.StatusBar1.Panels.Item(6).Text = "Not Connected."
            Exit Sub
        End If
        frmMain.StatusBar1.Panels.Item(6).Text = "Connected."
        
        prcProcessFindingCustomer
    Else
        prcFind2 LCase(Trim(Text7.Text)), Trim(Combo4.Text), Trim(Combo5.Text)
    End If
    
    'ListID = Trim(TDBGrid1.Columns(0).Value)
    'prcSearch KeyCode
End Sub

Sub prcProcessFindingCustomer()
        If IsNull(Trim(TDBGrid1.Columns(0).Value)) = True Then Exit Sub
        ListID = Trim(TDBGrid1.Columns(0).Value)
        prcDisableModifyNotes
        prcCallCustomerDtls
        prcClearMessageForm
        Frame1.Enabled = True
        Frame2.Enabled = True
        Exit Sub
End Sub

Sub prcFind2(sSearchFor As String, sWho As String, sSearchHow As String)
    
On Error Resume Next
    
    If rsGrabCustomerInfo.State = 1 Then
    
        Dim sSearchWho As String
        
        If sWho = "Account" Then
            sSearchWho = "cust_accountnumber"
        ElseIf sWho = "Customers" Then
            sSearchWho = "cust_fullname"
        ElseIf sWho = "Balance" Then
            sSearchWho = "cust_totalbalance"
        ElseIf sWho = "Status" Then
            sSearchWho = "cust_jobstatus"
        ElseIf sWho = "Rep" Then
            sSearchWho = "cust_salesrepref_fullname"
        ElseIf sWho = "Importance" Then
            sSearchWho = "importance_name"
        Else
            sSearchWho = ""
        End If
        
        If sSearchWho <> "" And sSearchHow <> "" Then
            rsGrabCustomerInfo.MoveFirst
            
            If sSearchFor <> "" Then
                If sSearchHow = "Contains" Then
                    rsGrabCustomerInfo.Find sSearchWho & " like '%" & sSearchFor & "%'", , adSearchForward
                ElseIf sSearchHow = "Exact match" Then
                    rsGrabCustomerInfo.Find sSearchWho & " = '" & sSearchFor & "'", , adSearchForward
                ElseIf sSearchHow = "Starts with" Then
                    rsGrabCustomerInfo.Find sSearchWho & " like '" & sSearchFor & "%'", , adSearchForward
                Else
                    Exit Sub
                End If
            End If
            If Not rsGrabCustomerInfo.EOF And Not rsGrabCustomerInfo.BOF Then
                TDBGrid1.Bookmark = rsGrabCustomerInfo.Bookmark
                TDBGrid1.ReBind
            End If
        End If
        
    End If
    
End Sub

Sub prcSearch(KeyCode As Integer)
    'Dim sSearchFor As String 'from user input
    'Dim vRow
   '
   ' sSearchFor = LCase(Trim(Text7.Text))
   '
   ' vRow = rsGrabCustomerInfo.Find("[" & rsGrabCustomerInfo.fields(TDBGrid1.col).Name & "] LIKE '%" & sSearchFor & "%'")
   '
   ' If vRow >= 0 Then TDBGrid1.Bookmark = vRow
   ' TDBGrid1.col = 1
   ' TDBGrid1.SetFocus
End Sub

Sub prcFind(sSearchFor As String, col As Long)
    Dim j As Integer
    Dim k
    Dim ilength As Integer
    Dim sSearchAgainst As String 'from collection
    Dim place
    Dim placetemp
    Dim placeCountUp
    
On Error GoTo errSkip:
    
    If sSearchFor = "" Then
        Exit Sub
    End If
    
    If sSearchFor <> "" Then
        ilength = Len(sSearchFor)
        
        
        If TDBGrid1.ApproxCount > 0 Then
            TDBGrid1.MoveFirst
        End If
        
            For j = 0 To TDBGrid1.ApproxCount - 1
        
    
            sSearchAgainst = LCase(Left(Trim(TDBGrid1.Columns(col).Value), ilength))
            
            If sSearchAgainst = sSearchFor Then
                ''''''''''''''''''''''''''''''''
                TDBGrid1.Bookmark = 1
                place = j
                placetemp = place
                
                If place <= 12 Then
                    TDBGrid1.Row = place
                Else
                    While placetemp > 12
                        TDBGrid1.Row = 12 'visible position
                        placeCountUp = placeCountUp + 12
                        placetemp = placetemp - 12
                        'TDBGrid1.Bookmark = placeCountUp + 1 'actualy position
                        If placetemp <= 12 Then
                            placeCountUp = placeCountUp + placetemp
                            'TDBGrid1.Bookmark = TDBGrid1.RowBookmark(place)
                            TDBGrid1.Row = placetemp
                        End If
                    Wend
                End If
                
                
                'MsgBox Trim(TDBGrid1.Columns(1).value)
                
                '''''''''''''''''''''''''''
                
                Exit Sub
                
            End If
            
            TDBGrid1.MoveNext
            
        Next j
            
    End If
    
    Exit Sub
    
errSkip:
    TDBGrid1.MoveFirst
    Exit Sub
End Sub

Private Sub Timer1_Timer()
    'if character.Wait(
End Sub

Private Sub tTimeZone_Timer()
    funZoneTime Trim(Combo9.Text)
End Sub

Private Sub wb3_DownloadComplete()
    Dim sSplitTest
    Dim sSplit1
    Dim sSplit2
    Dim sSplit3
    Dim stemp As String
    Dim stemp2 As String
    Dim i As Integer
    Dim WebStatus As String
    Dim sTempConvert As String
    
On Error Resume Next
    'Me.Caption = wb3.LocationName
    
    Text1.Text = Me.wb3.document.documentElement.innerHTML
    sSplit1 = Split(Trim(Text1.Text), "<BR>")
    
    Combo6.Clear
    Combo6.Text = ""
    
    If UBound(sSplit1) > 0 Then
    
        stemp = Trim(sSplit1(1))
        sSplitTest = Split(stemp, "-")
        
        If UBound(sSplitTest) = 0 Then
            sSplitTest = Split(stemp, ":")
            If UBound(sSplitTest) > 0 Then
                stemp2 = sSplitTest(1)
                Combo6.Text = sSplitTest(0) & ":" & funConvertStatus(stemp2)
                WebStatus = Trim(Combo6.Text)
            Else
                Combo6.Text = ""
            End If
        Else
            sSplit2 = Split(stemp, "-")
            
            If UBound(sSplit2) > 0 Then
                For i = 0 To UBound(sSplit2)
                    sSplit3 = Split(Trim(sSplit2(i)), ":")
                    stemp = Trim(sSplit3(0))
                    stemp2 = Trim(sSplit3(1))
                    sTempConvert = stemp & ":" & funConvertStatus(stemp2)
                    If i = 0 Then
                        Combo6.Text = sTempConvert
                        Combo6.AddItem sTempConvert
                        WebStatus = WebStatus & sTempConvert
                    Else
                        Combo6.AddItem sTempConvert
                        WebStatus = WebStatus & ", " & sTempConvert
                    End If
                Next i
            Else
                Combo6.Text = "Error (0x03)"
                WebStatus = Trim(Combo6.Text)
            End If
        End If
    Else
        Combo6.Text = "Error (0x04)"
        WebStatus = Trim(Combo6.Text)
    End If
    
    If WebStatus <> "" And strListID <> "" Then
        prcInsertWebStatusIntoCustomer WebStatus
    End If
End Sub

Sub prcInsertWebStatusIntoCustomer(WebStatus As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qbx_cust " & _
                " set cust_webstatus = '" & WebStatus & "' " & _
                " where cust_listid = '" & strListID & "' "
        
    cmdCommand.Execute
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "update webstatus for customer")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Function funConvertStatus(sTempStatus As String) As String
        Dim sStatus As String
        
        If sTempStatus = "A" Then
            sStatus = "ACTIVE"
        ElseIf sTempStatus = "C" Then
            sStatus = "DailyCap"
        ElseIf sTempStatus = "M" Then
            sStatus = "MonthlyCap"
        ElseIf sTempStatus = "P" Then
            sStatus = "MaxxedOut"
        ElseIf sTempStatus = "S" Then
            sStatus = "Suspended"
        ElseIf sTempStatus = "N" Then
            sStatus = "NoPayment"
        ElseIf sTempStatus = "I" Then
            sStatus = "Inactive"
        ElseIf sTempStatus = "X" Then
            sStatus = "Cancelled"
        End If
        
        funConvertStatus = sStatus
End Function
