VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11415
   Begin TabDlg.SSTab SSTab1 
      Height          =   8835
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   15584
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Alerts"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Importance"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Combo8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame3 
         Caption         =   "Modify Importance Level"
         Height          =   2715
         Left            =   2460
         TabIndex        =   28
         Top             =   1860
         Width           =   7455
         Begin VB.CheckBox Check1 
            Caption         =   "Importance Active"
            Height          =   195
            Left            =   4440
            TabIndex        =   36
            Top             =   1680
            Width           =   1635
         End
         Begin VB.CommandButton Command6 
            Appearance      =   0  'Flat
            Caption         =   "Update"
            Height          =   315
            Left            =   6300
            TabIndex        =   33
            Top             =   2160
            Width           =   915
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   4440
            TabIndex        =   32
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            Caption         =   "Remove"
            Height          =   315
            Left            =   4740
            TabIndex        =   31
            Top             =   2160
            Width           =   915
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   4440
            TabIndex        =   30
            Top             =   540
            Width           =   2535
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid2 
            Height          =   2175
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   3836
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Level ID"
            Columns(0).DataField=   ""
            Columns(0).DataWidth=   100
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Active"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Name"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0).DividerColor=   12307669
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=953"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=873"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=3678"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3598"
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
         Begin VB.Label Label16 
            Caption         =   "Importance Title:"
            Height          =   195
            Left            =   4440
            TabIndex        =   35
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Importance Level ID:"
            Height          =   195
            Left            =   4440
            TabIndex        =   34
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   6195
         Left            =   300
         TabIndex        =   40
         Top             =   960
         Width           =   5355
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   240
            TabIndex        =   44
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   240
            TabIndex        =   42
            Top             =   540
            Width           =   3375
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid1 
            Height          =   2175
            Left            =   120
            TabIndex        =   41
            Top             =   1920
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   3836
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Level ID"
            Columns(0).DataField=   ""
            Columns(0).DataWidth=   100
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Active"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Name"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0).DividerColor=   12307669
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=953"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=873"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=3678"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3598"
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
         Begin VB.Label Label18 
            Caption         =   "Create Importance"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Insert this new importance before "
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "frmOptions.frx":0038
         Left            =   1380
         List            =   "frmOptions.frx":0045
         TabIndex        =   37
         Top             =   480
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Caption         =   "Create Importance Title:"
         Height          =   2715
         Left            =   5280
         TabIndex        =   18
         Top             =   720
         Width           =   7455
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            Caption         =   "Create"
            Height          =   315
            Left            =   5640
            TabIndex        =   19
            Top             =   1080
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Remove Importance Level"
         Height          =   3735
         Left            =   4380
         TabIndex        =   13
         Top             =   4260
         Width           =   8715
         Begin VB.ListBox List2 
            Height          =   1815
            Left            =   5700
            TabIndex        =   25
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Re-assign"
            Height          =   315
            Left            =   6300
            TabIndex        =   24
            Top             =   3240
            Width           =   915
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   3180
            TabIndex        =   21
            Top             =   1560
            Width           =   2295
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   3180
            TabIndex        =   20
            Top             =   2580
            Width           =   2295
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Left            =   660
            TabIndex        =   15
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Remove"
            Height          =   315
            Left            =   1260
            TabIndex        =   14
            Top             =   3240
            Width           =   915
         End
         Begin VB.Label Label7 
            Caption         =   "Importance Title:"
            Height          =   195
            Left            =   6240
            TabIndex        =   26
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Re-assign From:"
            Height          =   195
            Left            =   3180
            TabIndex        =   23
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Re-assignTo:"
            Height          =   195
            Left            =   3180
            TabIndex        =   22
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   $"frmOptions.frx":0061
            Height          =   435
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   7575
         End
         Begin VB.Label Label10 
            Caption         =   "Importance Title:"
            Height          =   195
            Left            =   1200
            TabIndex        =   16
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4515
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   8775
         Begin VB.CommandButton Command1 
            Caption         =   "Update"
            Height          =   315
            Left            =   4560
            TabIndex        =   11
            Top             =   1560
            Width           =   975
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "frmOptions.frx":010E
            Left            =   3540
            List            =   "frmOptions.frx":0130
            TabIndex        =   5
            Text            =   "00"
            Top             =   1560
            Width           =   675
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "frmOptions.frx":015C
            Left            =   2400
            List            =   "frmOptions.frx":016C
            TabIndex        =   4
            Text            =   "1000"
            Top             =   1560
            Width           =   915
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2400
            TabIndex        =   3
            Top             =   1140
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmOptions.frx":0188
            Left            =   2400
            List            =   "frmOptions.frx":01B9
            TabIndex        =   2
            Text            =   "1"
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label13 
            Caption         =   "Anything at or above these settings will be marked with an alert status."
            Height          =   195
            Left            =   300
            TabIndex        =   27
            Top             =   240
            Width           =   5355
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   300
            TabIndex        =   12
            Top             =   2100
            Width           =   5235
         End
         Begin VB.Label Label5 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   10
            Top             =   1560
            Width           =   195
         End
         Begin VB.Label Label4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3360
            TabIndex        =   9
            Top             =   1560
            Width           =   75
         End
         Begin VB.Label Label3 
            Caption         =   "Total Open Balance:"
            Height          =   195
            Left            =   300
            TabIndex        =   8
            Top             =   1620
            Width           =   2115
         End
         Begin VB.Label Label2 
            Caption         =   "Importance Level:"
            Height          =   195
            Left            =   300
            TabIndex        =   7
            Top             =   1200
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Active Invoices:"
            Height          =   195
            Left            =   300
            TabIndex        =   6
            Top             =   780
            Width           =   1875
         End
      End
      Begin VB.Label Label14 
         Caption         =   "an Importance Level."
         Height          =   195
         Left            =   3780
         TabIndex        =   39
         Top             =   540
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "I would like to "
         Height          =   195
         Left            =   300
         TabIndex        =   38
         Top             =   540
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    prcUpdateAlertSettings
    prcGrabAllSettings
End Sub

Private Sub Form_Load()
    prcGrabAllSettings
End Sub

Sub prcInitSettingVars()
    prcClearForm
    prcFillImportanceDisplay
End Sub

Sub prcClearForm()
    Text1.Text = ""
    Combo1.Text = ""
    Combo2.Clear
    Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    Combo5.Clear
    Combo5.Text = ""
    Combo6.Clear
    Combo4.Text = ""
    List1.Clear
    List2.Clear
End Sub


Sub prcFillImportanceDisplay()
    Dim i As Integer
    
    List1.Clear
    Combo2.Clear
    
    For i = 0 To UBound(aryGImportLvl) - 1
        'fill importance management
        List1.AddItem aryGImportLvl(i, 0) & ":" & aryGImportLvl(i, 1) & ":" & aryGImportLvl(i, 2)
        'fill alert management
        If aryGImportLvl(i, 0) = 1 Then
            Combo2.AddItem aryGImportLvl(i, 1) & ":" & aryGImportLvl(i, 2)
        End If
    Next i
End Sub

Sub prcGrabAllSettings()
    prcInitSettingVars
    prcGrabbingAlertSettings
End Sub

'''''''''''''''''''''''''''''Alert Settings'''''''''''''''''''''''''''''''
Sub prcGrabbingAlertSettings()
    
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsAlertSettings As New ADODB.Recordset
    Dim sMoney As String
    Dim i As Integer
    Dim lvl As Integer
    Dim sSpltMoney

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_alert_settings "
            
    Set rsAlertSettings = cmdCommand.Execute
    
    If Not rsAlertSettings.EOF Then
        rsAlertSettings.MoveFirst
        Combo1.Text = Trim(rsAlertSettings!alert_setting_max_invoices) & ""
        For i = 0 To UBound(aryGImportLvl) - 1
            lvl = Trim(rsAlertSettings!alert_setting_start_at_level)
            If aryGImportLvl(i, 1) = lvl Then
                Combo2.Text = aryGImportLvl(i, 1) & ":" & aryGImportLvl(i, 2)
            End If
        Next i
        
        sMoney = Trim(rsAlertSettings!alert_setting_max_dollar) & ""
        If sMoney <> "" Then
            sSpltMoney = Split(sMoney, ".")
            If UBound(sSpltMoney) > 0 Then
                Combo3.Text = sSpltMoney(0)
                Combo4.Text = sSpltMoney(1)
            Else
                Combo3.Text = sMoney
                Combo4.Text = "00"
            End If
        End If
    End If
    
    Set rsAlertSettings = Nothing
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
    Set rsAlertSettings = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub


Sub prcUpdateAlertSettings()
    
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sMoney As String
    Dim sSpltLvl
    Dim iLvl As Integer
    Dim iInvoices As Integer
    Dim iCheckGood As Integer

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    iCheckGood = 0
    'checking dollar values
    If IsNumeric(Trim(Combo3.Text)) = True And IsNumeric(Trim(Combo4.Text)) Then
        sMoney = Trim(Combo3.Text) & "." & Trim(Combo4.Text)
        iCheckGood = 1
    End If
    'checking level number
    sSpltLvl = Split(Trim(Combo2.Text), ":")
    If UBound(sSpltLvl) > 0 Then
        If IsNumeric(Trim(sSpltLvl(0))) = True Then
            iLvl = Trim(sSpltLvl(0))
            iCheckGood = iCheckGood + 2
        End If
    End If
    'checking max invoices
    If IsNumeric(Trim(Combo1.Text)) = True Then
        iInvoices = Trim(Combo1.Text)
        iCheckGood = iCheckGood + 4
    End If
    
    If iCheckGood <> 7 Then
        If iCheckGood = 1 Then
            Label6.Caption = "Error with Importance level and Invoices"
        ElseIf iCheckGood = 2 Then
            Label6.Caption = "Error with Total Balance and Invoices"
        ElseIf iCheckGood = 3 Then
            Label6.Caption = "Error with Invoices"
        ElseIf iCheckGood = 4 Then
            Label6.Caption = "Error with Total Balance and Importance level"
        ElseIf iCheckGood = 5 Then
            Label6.Caption = "Error with Importance level"
        ElseIf iCheckGood = 6 Then
            Label6.Caption = "Error with Total Balance"
        ElseIf iCheckGood = 0 Then
            Label6.Caption = "Error with all settings"
        Else
            Label6.Caption = "Error unknown"
        End If
        Exit Sub
    End If
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qbx_alert_settings " & _
                        " set alert_setting_max_dollar = CONVERT(MONEY, '" & sMoney & "'), " & _
                        " alert_setting_start_at_level = '" & iLvl & "', " & _
                        " alert_setting_max_invoices = '" & iInvoices & "' "
            
    cmdCommand.Execute
    
    Label6.Caption = "Alert Settings updated."
    
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
