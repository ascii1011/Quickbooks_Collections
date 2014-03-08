VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmQuickBooksFaxes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quickbooks Faxes"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   Icon            =   "frmQuickBooksFaxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   12015
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4215
      Left            =   7560
      TabIndex        =   16
      Top             =   0
      Width           =   4275
      Begin VB.TextBox txtBatch_Item 
         Height          =   315
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtInvoice 
         Height          =   315
         Left            =   2820
         TabIndex        =   24
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   960
         TabIndex        =   20
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   960
         TabIndex        =   19
         Top             =   2340
         Width           =   675
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   2340
         Width           =   255
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   2820
         TabIndex        =   17
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Batch:"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Invoice:"
         Height          =   195
         Left            =   2160
         TabIndex        =   33
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "To:"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Subject:"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Attach:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Date:"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Sent:"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Retries:"
         Height          =   195
         Left            =   1740
         TabIndex        =   27
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Status:"
         Height          =   195
         Left            =   2160
         TabIndex        =   26
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command10 
         Caption         =   "Preview"
         Height          =   315
         Left            =   6240
         TabIndex        =   15
         ToolTipText     =   "Print Priview"
         Top             =   3000
         Width           =   795
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmQuickBooksFaxes.frx":0442
         Left            =   480
         List            =   "frmQuickBooksFaxes.frx":0444
         TabIndex        =   6
         Top             =   3000
         Width           =   2595
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmQuickBooksFaxes.frx":0446
         Left            =   3660
         List            =   "frmQuickBooksFaxes.frx":0453
         TabIndex        =   5
         Text            =   "All"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmQuickBooksFaxes.frx":046C
         Left            =   3660
         List            =   "frmQuickBooksFaxes.frx":047C
         TabIndex        =   4
         Text            =   "All"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   315
         Left            =   480
         TabIndex        =   3
         Text            =   "0"
         Top             =   3720
         Width           =   555
      End
      Begin VB.TextBox Text18 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Text            =   "100"
         Top             =   3720
         Width           =   555
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   2355
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4154
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Batch"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Date"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Invoice"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "To"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Received"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Retries"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1191"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1111"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=4286"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4207"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1826"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1746"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1058"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=979"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=56,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=53,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=54,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=55,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=48,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=45,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=46,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=47,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=44,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=27,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=28,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=43,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(82)  =   ":id=25,.parent=33,.fgcolor=&HFF&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(83)  =   ":id=25,.strikethrough=0,.charset=0"
         _StyleDefs(84)  =   ":id=25,.fontname=MS Sans Serif"
         _StyleDefs(85)  =   "Named:id=26:Balance"
         _StyleDefs(86)  =   ":id=26,.parent=25,.fgcolor=&HC0C0C0&,.borderColor=&H80000007&,.bold=-1"
         _StyleDefs(87)  =   ":id=26,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(88)  =   ":id=26,.fontname=MS Sans Serif"
      End
      Begin VB.Label Label15 
         Caption         =   "Batch:"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   6540
         TabIndex        =   13
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Count:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5940
         TabIndex        =   12
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label26 
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   2760
         Width           =   435
      End
      Begin VB.Label Label27 
         Caption         =   "Sent:"
         Height          =   195
         Left            =   3660
         TabIndex        =   10
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Received:"
         Height          =   195
         Left            =   3660
         TabIndex        =   9
         Top             =   3480
         Width           =   915
      End
      Begin VB.Label Label31 
         Caption         =   "To"
         Height          =   195
         Left            =   1140
         TabIndex        =   8
         Top             =   3780
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "Limit between Batch #"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   3480
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmQuickBooksFaxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim rsBatchResults              As New ADODB.Recordset
Dim bSortOrd As Boolean


Private Sub Combo2_Click()
    prcParseBatchIndex
End Sub

Private Sub Combo3_Click()
    prcParseBatchIndex
End Sub

Private Sub Combo4_Click()
    prcParseBatchIndex
End Sub

Private Sub Command10_Click()
    With TDBGrid1.PrintInfo
        ' Set the page header
        .PageHeaderFont.Italic = True
        .PageHeader = "Composers table"
        
        ' Column headers will be on every page
        .RepeatColumnHeaders = True
        
        ' Display page numbers (centered)
        .PageFooter = "\tPage: \p"
        ' Invoke Print Preview
        .PrintPreview
    End With
End Sub



Private Sub Form_Load()
    Me.Width = 7590
    Me.Height = 4800
    sfrmQuickBooksFaxes = 1
    Text17.Enabled = False
    Text18.Enabled = False
    prcGetAllBatch
End Sub




Sub prcParseBatchIndex()
    Dim ssplit() As String
    
    ssplit = Split(Trim(Combo2.Text), ":")
    
    If UBound(ssplit) > 0 Then
        prcGetCurrentBatch Trim(ssplit(0))
    Else
        If Trim(Combo2.Text) = "All" Then
            Text17.Enabled = True
            Text18.Enabled = True
            prcGetCurrentBatch 0
        ElseIf IsNumeric(Trim(Combo2.Text)) Then
            Text17.Enabled = False
            Text18.Enabled = False
            prcGetCurrentBatch Trim(Combo2.Text)
        End If
    End If
End Sub



Sub prcGetCurrentBatch(lBatchNumber As Long)
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim sSQL As String, sSQL2 As String, sSQL3 As String
    Dim sTmp                    As String
    Dim sFromInv   As String
    Dim sToInv As String
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    
    Set rsBatchResults = Nothing
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    sFromInv = Trim(Text17.Text)
    sToInv = Trim(Text18.Text)
    
    If lBatchNumber = 0 Then
    
        sSQL = " select * from mc_send2fax_mailitems "
                    
        If Trim(Combo3.Text) <> "All" Then
            sSQL2 = sSQL2 & " where item_sent = '" & funReturnSentResult & "' "
            If Trim(Combo4.Text) <> "All" Then
                sSQL2 = sSQL2 & " and item_recieved = '" & funReturnReceivedResult & "' "
            End If
        Else
            If Trim(Combo4.Text) <> "All" Then
                sSQL2 = sSQL2 & " where item_recieved = '" & funReturnReceivedResult & "' "
            End If
        End If
        
        If sFromInv <> 0 Then
            If sSQL2 = "" Then
                sSQL3 = " where item_batch >= '" & sFromInv & "' and item_batch <= '" & sToInv & "' "
            Else
                sSQL3 = " and item_batch >= '" & sFromInv & "' and item_batch <= '" & sToInv & "' "
            End If
        End If
        sSQL = sSQL & sSQL2 & sSQL3
        'MsgBox sSQL
    Else
    
        sSQL = " select * from mc_send2fax_mailitems " & _
            " where item_batch = '" & lBatchNumber & "' "
                
        If Trim(Combo3.Text) <> "All" Then
            sSQL = sSQL & " and item_sent = '" & funReturnSentResult & "' "
        End If
        If Trim(Combo4.Text) <> "All" Then
            sSQL = sSQL & " and item_recieved = '" & funReturnReceivedResult & "' "
        End If
        
    End If
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    Set rsBatchResults = cmdCommand.Execute
    
    Label16.Caption = rsBatchResults.RecordCount
    
    TDBGrid1.ReBind
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    MsgBox Err.Number & vbNewLine & Err.Description
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

Function funReturnSentResult() As String
    Dim sTmp    As String
    
    sTmp = Trim(Combo3.Text)
    If sTmp = "Sent" Then
        funReturnSentResult = "True"
    Else
        funReturnSentResult = ""
    End If
End Function


Function funReturnReceivedResult() As String
    funReturnReceivedResult = Trim(Combo4.Text)
End Function


Sub prcGetAllBatch()
    Dim Response
    Dim cmdCommand              As New ADODB.Command
    Dim parParameter            As New ADODB.Parameter
    Dim rsBatchList             As New ADODB.Recordset
    Dim sSQL                    As String
    Dim i As Integer
    Dim lLastBatchIndex         As Long
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    
    Set rsBatchResults = Nothing
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
        
    sSQL = " select * from mc_send2fax_batch "
        
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sSQL
        
    Set rsBatchList = cmdCommand.Execute
    
    If Not rsBatchList.EOF Then
        Label26.Caption = "(" & rsBatchList.RecordCount & ")"
        rsBatchList.MoveFirst
        
        Combo2.AddItem "All"
        
        While Not rsBatchList.EOF
            i = i + 1
            If i = 1 Then
                Combo2.Text = Trim(rsBatchList!batch_index) & " : [" & Format(Trim(rsBatchList!Batch_datetime), "MM-DD-YY") & "]"
                'If rsBatchList!batch_finished = "True" Then
                '    Combo2.Text = Trim(rsBatchList!batch_index) & ":" & Format(Trim(rsBatchList!Batch_datetime), "MM-DD-YY")
                'Else
                '    Combo2.Text = Trim(rsBatchList!batch_index) & ":" & Format(Trim(rsBatchList!Batch_datetime), "MM-DD-YY") & " Done:" & Trim(rsBatchList!batch_finished)
                'End If
            End If
            Combo2.AddItem Trim(rsBatchList!batch_index) & " : [" & Format(Trim(rsBatchList!Batch_datetime), "MM-DD-YY") & "]"
            'If rsBatchList!batch_finished = "True" Then
            '    Combo2.AddItem Trim(rsBatchList!batch_index) & ":" & Format(Trim(rsBatchList!Batch_datetime), "MM-DD-YY")
            'Else
            '    Combo2.AddItem Trim(rsBatchList!batch_index) & ":" & Format(Trim(rsBatchList!Batch_datetime), "MM-DD-YY") & " Done:" & Trim(rsBatchList!batch_finished)
            'End If
            
            lLastBatchIndex = Trim(rsBatchList!batch_index)
            
            rsBatchList.MoveNext
        Wend
        
        prcGetCurrentBatch lLastBatchIndex
    Else
        prcGetCurrentBatch 0
    End If
        
    Set rsBatchList = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    MsgBox Err.Number & vbNewLine & Err.Description
    Set rsBatchList = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    sfrmQuickBooksFaxes = 0
End Sub

Private Sub TDBGrid1_Click()
   TDBGrid1.PostMsg 1
End Sub


Private Sub TDBGrid1_DblClick()
    prcGrabInfo
End Sub


Sub prcGrabInfo()
    Dim InvoiceID As String
    

On Error Resume Next

    InvoiceID = Trim(TDBGrid1.Columns(2).Value)
    
    ListID = funGrabCustIDByInvoice(InvoiceID)
    'frmInvoiceQry.prcFind Trim(TDBGrid1.Columns(0).value), 0
    frmInvoiceQry.prcCallCustomerDtls
    
    frmInvoiceQry.SSTab1.Tab = 0
    frmInvoiceQry.Command11.Enabled = True
    frmInvoiceQry.Frame2.Enabled = True
    
    'frmInvoiceQry.Text7.Text = Trim(TDBGrid_Search.Columns(2).Value)
    'frmInvoiceQry.prcFind2 LCase(Trim(frmInvoiceQry.Text7.Text)), Trim(frmInvoiceQry.Combo4.Text), Trim(frmInvoiceQry.Combo5.Text)
    'frmInvoiceQry.prcProcessFindingCustomer
    
    'Unload Me
End Sub


Function funGrabCustIDByInvoice(sinvoice As String) As String
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsReturnCustID  As New ADODB.Recordset

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then Exit Function
    funGrabCustIDByInvoice = ""
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select " & _
                " cust.cust_listid " & _
                " from qbx_inv inv " & _
                " left join qbx_cust cust on cust.cust_listid = inv.inv_customerref_listid " & _
                " where inv.inv_refnumber  =  '" & sinvoice & "'  "
            
    Set rsReturnCustID = cmdCommand.Execute
    
    If Not rsReturnCustID.EOF Then
        rsReturnCustID.MoveFirst
        funGrabCustIDByInvoice = Trim(rsReturnCustID!cust_listid) & ""
    End If
    
    Set rsReturnCustID = Nothing
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
    Set rsReturnCustID = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function


'''''''''''''''''''  Grid   '''''''''''''''''''''''''''''''''''

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
On Error Resume Next

    MousePointer = vbHourglass
    If bSortOrd Then
        rsBatchResults.Sort = "[" & rsBatchResults.Fields(ColIndex).name & "]  ASC"
        bSortOrd = False
    Else
        rsBatchResults.Sort = "[" & rsBatchResults.Fields(ColIndex).name & "]  DESC"
        bSortOrd = True
    End If
    TDBGrid1.col = ColIndex
    rsBatchResults.MoveFirst
    TDBGrid1.ReBind
    MousePointer = vbDefault

End Sub



Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)

On Error Resume Next

    'If TDBGrid1.Columns(2) = -1 Then
    '    sUpdateSelectValueForCurrentMailItem_batchNumber = Trim(TDBGrid1.Columns(0).Value)
    '    sUpdateSelectValueForCurrentMailItem_ItemNumber = Trim(TDBGrid1.Columns(1).Value)
    'Else
    '    MsgBox TDBGrid1.Columns(6)
    'End If
    
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
    Dim vtemp

On Error GoTo NoRead
    
    'If bCancelRead Then Exit Sub
    cols = RowBuf.ColumnCount - 1
    Rows = RowBuf.RowCount - 1
    RowsFetched = 0
    
    If IsNull(StartLocation) Then
        If offset < 0 Then
            rsBatchResults.MoveLast
            rsBatchResults.MoveNext
        Else
            rsBatchResults.MoveFirst
            rsBatchResults.MovePrevious
        End If
        rsBatchResults.Move offset
    Else
        rsBatchResults.Move offset, StartLocation
    End If
        
    StartRow = rsBatchResults.Bookmark
    Pos = rsBatchResults.AbsolutePosition
    
    For Row = 0 To Rows
        If rsBatchResults.BOF Or rsBatchResults.EOF Then Exit For
        For col = 0 To cols
            
            If Val(Trim(rsBatchResults!item_retries)) <> 0 Then
                vtemp = (Val(Trim(rsBatchResults!item_retries)) + 1) * 4
            Else
                vtemp = 0
            End If
            
            Select Case (col)
                'Case (0):   RowBuf.Value(Row, 0) = Trim(rsBatchResults!item_batch) & ""
                'Case (1):   RowBuf.Value(Row, 1) = Trim(rsBatchResults!item_number) & ""
                'Case (2):   RowBuf.Value(Row, 2) = Trim(rsBatchResults!item_invoice) & ""
                'Case (3):   RowBuf.Value(Row, 3) = Trim(rsBatchResults!item_to) & ""
                'Case (4):   RowBuf.Value(Row, 4) = Trim(rsBatchResults!item_sent) & ""
                'Case (5):   RowBuf.Value(Row, 5) = Trim(rsBatchResults!item_retries) & ""
                'Case (6):   RowBuf.Value(Row, 6) = Trim(rsBatchResults!item_recieved) & ""
                'Case (7):   RowBuf.Value(Row, 7) = Format(Trim(rsBatchResults!item_datetime), "dd/mm/yyyy") & ""
                
                
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsBatchResults!item_batch) & ""
                Case (1):   RowBuf.Value(Row, 1) = Format(Trim(rsBatchResults!item_datetime), "dd/mm/yyyy") & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsBatchResults!item_invoice) & ""
                Case (3):   RowBuf.Value(Row, 3) = Trim(rsBatchResults!item_to) & ""
                Case (4):   RowBuf.Value(Row, 4) = Trim(rsBatchResults!item_recieved) & ""
                Case (5):   RowBuf.Value(Row, 5) = vtemp
                
                
                'Case (8):   RowBuf.Value(Row, 8) = Format(Trim(rsBatchResults!item_datetime), "dd/mm/yyyy") & ""
                'Case (7):   RowBuf.Value(Row, 7) = Trim(rsBatchResults!item_selected) & ""
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsBatchResults.Bookmark
        RowsFetched = RowsFetched + 1
        rsBatchResults.MoveNext
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

Private Sub Text17_Change()
    prcParseBatchIndex
End Sub

Private Sub Text18_Change()
    prcParseBatchIndex
End Sub
