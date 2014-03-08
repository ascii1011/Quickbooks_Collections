VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmPriorityAlerts 
   Caption         =   "Priority Alerts"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "frmPriorityAlerts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   8685
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7860
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriorityAlerts.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriorityAlerts.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriorityAlerts.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "All"
            Object.ToolTipText     =   "All"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Active"
            Object.ToolTipText     =   "Active"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inactive"
            Object.ToolTipText     =   "Inactive"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2790
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3598
            MinWidth        =   3598
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3836
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
      Columns(1).Caption=   "Customer"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Importance"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Balance"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Invoices"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Active?"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12307669
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=423"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=344"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=6006"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5927"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2672"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2593"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1746"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1667"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1402"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1323"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1296"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1217"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=51,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=48,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=49,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=50,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=59,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=52,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=53,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=54,.parent=17"
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
      _StyleDefs(81)  =   "Named:id=47:Open"
      _StyleDefs(82)  =   ":id=47,.parent=42,.fgcolor=&H808000&,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(83)  =   ":id=47,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(84)  =   ":id=47,.fontname=MS Sans Serif"
      _StyleDefs(85)  =   "Named:id=68:CA green"
      _StyleDefs(86)  =   ":id=68,.parent=47,.fgcolor=&HFF&"
   End
End
Attribute VB_Name = "frmPriorityAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public rsPriorityAlerts          As New ADODB.Recordset
Dim bSortOrd As Boolean
Dim bHeadValue As String


Private Sub Form_Load()
    Me.Width = 8805
    Me.Height = 3465
    TDBGrid1.FetchRowStyle = True
    sFrmPriorityAlerts = 1
    Me.StatusBar1.Panels.Item(1).Text = "Records: "
    Me.Show
    'TDBGrid1.FetchRowStyle = True
        prcGrabPriorityAlerts "True", "Asc"
End Sub

Private Sub Form_Resize()
On Error Resume Next

    If Me.WindowState = 1 Then
        Me.WindowState = 0
    End If
    TDBGrid1.Width = Me.Width - 400
    TDBGrid1.Height = Me.Height - 1500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If sKillPriorityAlert = False Then
        Cancel = 1
    Else
        sFrmPriorityAlerts = 0
        Cancel = 0
    End If
End Sub

Sub prcSearch(ByWhat As String, SearchFor As String)

On Error Resume Next

    If ByWhat = "name" Then ByWhat = "cust_fullname"
    
    If SearchFor <> "" Then
            
        If Not rsPriorityAlerts.RecordCount = 0 Then
            If Not rsPriorityAlerts.EOF Then
                rsPriorityAlerts.MoveFirst
                rsPriorityAlerts.Find ByWhat & " = '" & Trim(SearchFor) & "'", , adSearchForward
                TDBGrid1.Bookmark = rsPriorityAlerts.Bookmark
                TDBGrid1.ReBind
            End If
        End If
    
    Else
        TDBGrid1.Bookmark = 1
    End If
      
End Sub

Sub prcGrabPriorityAlerts(sDisplay_active As String, sOrder As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim strXtraSQL      As String
    Dim ExtraWhere As String

On Error GoTo errHandle:
    If rsPriorityAlerts.State = 1 Then
        Set rsPriorityAlerts = Nothing
    End If
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then
        Exit Sub
    End If
    
    strXtraSQL = " select a.*, b.* from qbx_alerts a, qbx_cust b "
    
    If sDisplay_active <> "" Then
        ExtraWhere = " a.cust_isactive = '" & sDisplay_active & "' "
    End If
    
    strXtraSQL = strXtraSQL & funReturnPriviledgeRestrictions(ExtraWhere)
        
    If sOrder = "Asc" Or sOrder = "Desc" Then
        strXtraSQL = strXtraSQL & " order by a.alert_total_balance " & sOrder & " "
    ElseIf sOrder = "Name Asc" Then
        strXtraSQL = strXtraSQL & " order by a.cust_fullname asc "
    ElseIf sOrder = "Name Desc" Then
        strXtraSQL = strXtraSQL & " order by a.cust_fullname desc "
    End If
    
    'MsgBox strXtraSQL
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = strXtraSQL
        
    Set rsPriorityAlerts = cmdCommand.Execute
    
    'Label21.Caption = rsPriorityAlerts.RecordCount
    strXtraSQL = " Records: " & rsPriorityAlerts.RecordCount
    Me.StatusBar1.Panels.Item(1).Text = strXtraSQL
    
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





Private Sub TDBGrid1_DblClick()
    prcGrabPriorityInfo
End Sub



Sub prcGrabPriorityInfo()

On Error Resume Next

    ListID = Trim(TDBGrid1.Columns(0).Value)
    'frmInvoiceQry.prcFind Trim(TDBGrid1.Columns(0).value), 0
    frmInvoiceQry.prcCallCustomerDtls
    
    frmInvoiceQry.SSTab1.Tab = 0
    frmInvoiceQry.Command11.Enabled = True
    frmInvoiceQry.Frame2.Enabled = True
    'If Trim(Text1.Text) <> "" Then
    
        'msg
        'If Trim(TDBGrid1.Columns(3).Value) <> "" Then
            'frmInvoiceQry.Text6.Text = Trim(TDBGrid1.Columns(3).Value) & vbNewLine & vbNewLine
        'Else
            'frmInvoiceQry.Text6.Text = "<Empty>" & vbNewLine & vbNewLine
        'End If
        
        'frmInvoiceQry.Text6.Text = frmInvoiceQry.Text6.Text & "By: " & Trim(TDBGrid1.Columns(5).Value)
        
        'created datestamp
        'frmInvoiceQry.Text5.Text = Trim(TDBGrid1.Columns(4).Value)
        
        'callback date
        'frmInvoiceQry.Text11.Text = (Text1.Text)
        
        'callback time
        'frmInvoiceQry.Text12.Text = Trim(TDBGrid1.Columns(1).Value)
    'End If
    
End Sub


Private Sub TDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
On Error Resume Next

    Response = 0
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)

On Error GoTo errHandle
       
    rsPriorityAlerts.Bookmark = Bookmark
    If Left(Trim(rsPriorityAlerts!Alert_total_balance), 1) = "-" Then
        RowStyle = TDBGrid1.Styles(10)
    Else
        RowStyle = TDBGrid1.Styles(11)
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
        If ColIndex = 0 Then
            rsPriorityAlerts.Sort = "[alert_id]  ASC"
        ElseIf ColIndex = 1 Then
            rsPriorityAlerts.Sort = "[cust_fullname]  ASC"
        ElseIf ColIndex = 2 Then
            rsPriorityAlerts.Sort = "[alert_importance_type]  ASC"
        ElseIf ColIndex = 3 Then
            rsPriorityAlerts.Sort = "[alert_total_balance]  ASC"
        ElseIf ColIndex = 4 Then
            rsPriorityAlerts.Sort = "[alert_total_invoices]  ASC"
        ElseIf ColIndex = 5 Then
            rsPriorityAlerts.Sort = "[cust_isactive]  ASC"
        End If
        bSortOrd = False
    Else
        If ColIndex = 0 Then
            rsPriorityAlerts.Sort = "[alert_id]  DESC"
        ElseIf ColIndex = 1 Then
            rsPriorityAlerts.Sort = "[cust_fullname]  DESC"
        ElseIf ColIndex = 2 Then
            rsPriorityAlerts.Sort = "[alert_importance_type]  DESC"
        ElseIf ColIndex = 3 Then
            rsPriorityAlerts.Sort = "[alert_total_balance]  DESC"
        ElseIf ColIndex = 4 Then
            rsPriorityAlerts.Sort = "[alert_total_invoices]  DESC"
        ElseIf ColIndex = 5 Then
            rsPriorityAlerts.Sort = "[cust_isactive]  DESC"
        End If
            
        'rsPriorityAlerts.Sort = "[" & rsPriorityAlerts.Fields(ColIndex).name & "]  DESC"
        
        bSortOrd = True
    End If
    TDBGrid1.col = ColIndex
    rsPriorityAlerts.MoveFirst
    TDBGrid1.Bookmark = rsPriorityAlerts.Bookmark
    TDBGrid1.ReBind
    MousePointer = vbDefault

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
            rsPriorityAlerts.MoveLast
            rsPriorityAlerts.MoveNext
        Else
            rsPriorityAlerts.MoveFirst
            rsPriorityAlerts.MovePrevious
        End If
        rsPriorityAlerts.Move offset
    Else
        rsPriorityAlerts.Move offset, StartLocation
    End If
        
    StartRow = rsPriorityAlerts.Bookmark
    Pos = rsPriorityAlerts.AbsolutePosition
    
    For Row = 0 To Rows
        If rsPriorityAlerts.BOF Or rsPriorityAlerts.EOF Then Exit For
        For col = 0 To cols
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsPriorityAlerts!alert_id) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsPriorityAlerts!cust_fullname) & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsPriorityAlerts!alert_importance) & ""
                Case (3):   RowBuf.Value(Row, 3) = "$" & funFormatDecimal(Trim(rsPriorityAlerts!Alert_total_balance))
                Case (4):   RowBuf.Value(Row, 4) = Trim(rsPriorityAlerts!Alert_total_invoices) & ""
                Case (5):   RowBuf.Value(Row, 5) = Trim(rsPriorityAlerts!cust_isactive) & ""
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsPriorityAlerts.Bookmark
        RowsFetched = RowsFetched + 1
        rsPriorityAlerts.MoveNext
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim Response

On Error GoTo errHandle
    
    Select Case (Button.Key)
        Case ("All"): prcGrabPriorityAlerts "", "Asc"
        Case ("Active"):   prcGrabPriorityAlerts "True", "Asc"
        Case ("Inactive"):   prcGrabPriorityAlerts "False", "Asc"
    End Select
    
    Exit Sub

errHandle:
    'Debug.Print err.number
    Select Case (Err.Number)
        Case (91):
            Resume Next
        Case (364):
            Resume Next 'error closing qb connection while unloading
        Case (438):
            MsgBox Err.Number & vbNewLine & Err.Description
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Run Time Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub
