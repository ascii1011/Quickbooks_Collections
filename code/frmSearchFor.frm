VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmSearchFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search By ..."
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7500
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   7395
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSearchFor.frx":0000
         Left            =   120
         List            =   "frmSearchFor.frx":0002
         TabIndex        =   4
         Text            =   "Invoice"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmSearchFor.frx":0004
         Left            =   2400
         List            =   "frmSearchFor.frx":0014
         TabIndex        =   3
         Text            =   "Equal to"
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   4860
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid_Search 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   4048
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
         Columns(2).Caption=   "Customer"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12307669
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=132"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=53"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2090"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2011"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=5133"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=5054"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=4180"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4101"
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
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
      Begin VB.Label Label1 
         Caption         =   "Count:"
         Height          =   195
         Left            =   5820
         TabIndex        =   7
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   6360
         TabIndex        =   6
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1995
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSearchFor.frx":0046
      Top             =   3540
      Width           =   7275
   End
End
Attribute VB_Name = "frmSearchFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsSearch        As New ADODB.Recordset


'''''''''''''''
Public sVar_Field_Text As String    'value retrieved from interface
Public sVar_Field As String         'value to be used for query

Public sVar_EvalWith_Text As String 'value retrieved from interface
Public sVar_EvalWith_Part1 As String      'value to be used for query
Public sVar_EvalWith_Part2 As String      'value to be used for query

Public sVar_Value_Text As String    'value retrieved from interface
Public sVar_Value As String         'value to be used for query
'''''''''''''''

Public sVal_type As String  'tells what the formatting is for the string to be evaluated, ex. string, int, date, money

Public sWhichQuery As String 'tells what query to use

Public sOrder As String
Public sOrderBy As String

Public sValue_Formatting As String

Public sQuery As String

Dim bInternalOperation As Boolean


Private Sub Combo1_Change()
    If bInternalOperation = False Then
        Combo1.Text = ""
    End If
End Sub

Private Sub Combo1_Click()

    
    prcEvalField
    prcSetNextField
    
    Text1.SetFocus
End Sub


Sub prcEvalField()
    'field name
    
    sVar_Field_Text = Trim(Combo1.Text)
    
    Select Case (sVar_Field_Text)
        Case "Invoice":  sVar_Field = "inv_refnumber "   'qbx_inv
                        sWhichQuery = " qbx_inv "
                        sVal_type = "string"
                        sValue_Formatting = "string"
        Case "Invoice Balance":  sVar_Field = "inv_balanceremaining "   'qbx_inv
                        sWhichQuery = " qbx_inv "
                        sVal_type = "string"
                        sValue_Formatting = "money"
                        bInternalOperation = True
                        Text1.Text = 0
                        bInternalOperation = False
        Case "Invoice Applied Amount":  sVar_Field = "inv_appliedamout "   'qbx_inv
                        sWhichQuery = " qbx_inv "
                        sVal_type = "string"
                        sValue_Formatting = "money"
                        bInternalOperation = True
                        Text1.Text = 0
                        bInternalOperation = False
        Case "Invoice Subtotal":  sVar_Field = "inv_subtotal "   'qbx_inv
                        sWhichQuery = " qbx_inv "
                        sVal_type = "string"
                        sValue_Formatting = "money"
                        bInternalOperation = True
                        Text1.Text = 0
                        bInternalOperation = False
        Case "Invoice Date":  sVar_Field = "inv_txndate "   'qbx_inv
                        sWhichQuery = " qbx_inv "
                        sVal_type = "date"
                        sValue_Formatting = "date"
        
        Case "Total Balance":  sVar_Field = "cust_totalbalance_money "   'qbx_cust
                        sWhichQuery = " qbx_cust "
                        sVal_type = "money"
                        sValue_Formatting = "money"
                        bInternalOperation = True
                        Text1.Text = 0
                        bInternalOperation = False
        Case "Phone":  sVar_Field = "cust_phone1 "   'qbx_cust
                        sWhichQuery = " qbx_cust "
                        sVal_type = "string"
                        sValue_Formatting = "string"
        Case "Fax":  sVar_Field = "cust_fax1 "   'qbx_cust
                        sWhichQuery = " qbx_cust "
                        sVal_type = "string"
                        sValue_Formatting = "string"
        Case "Email":  sVar_Field = "cust_email1 "   'qbx_cust
                        sWhichQuery = " qbx_cust "
                        sVal_type = "string"
                        sValue_Formatting = "string"
        Case "Contact":  sVar_Field = "cust_contact "   'qbx_cust
                        sWhichQuery = " qbx_cust "
                        sVal_type = "string"
                        sValue_Formatting = "string"
        Case "Alt Contact":  sVar_Field = "cust_altcontact "   'qbx_cust
                        sWhichQuery = " qbx_cust "
                        sVal_type = "string"
                        sValue_Formatting = "string"
        
        Case Else
                sVar_Field = ""
                sWhichQuery = ""
    End Select
    
    
End Sub



Private Sub Combo2_Change()
    If bInternalOperation = False Then
        Combo2.Text = ""
    End If
End Sub

Private Sub Combo2_Click()
    sVar_EvalWith_Text = Trim(Combo2.Text)
    prcEvalField
    prcEvalWith
    
    Text1.SetFocus
End Sub

Sub prcSetNextField()

    
                bInternalOperation = True
    
    Combo2.Text = ""
    Combo2.Clear
    
    If sValue_Formatting = "string" Then
        Combo2.Text = "Equal to"
        Combo2.AddItem "Equal to"
        Combo2.AddItem "Similar to"
        Combo2.AddItem "Starts with"
        Combo2.AddItem "Ends with"
    ElseIf sValue_Formatting = "money" Then
        Combo2.Text = "Equal to"
        Combo2.AddItem "Equal to"
        Combo2.AddItem "Greater than"
        Combo2.AddItem "Equal or Greater than"
        Combo2.AddItem "Less than"
        Combo2.AddItem "Equal or Less than"
    ElseIf sValue_Formatting = "date" Then
        Combo2.Text = "Equal to"
        Combo2.AddItem "Equal to"
        Combo2.AddItem "Greater than"
        Combo2.AddItem "Equal or Greater than"
        Combo2.AddItem "Less than"
        Combo2.AddItem "Equal or Less than"
    End If
    
                bInternalOperation = False
    
    prcEvalWith
    
End Sub

Sub prcEvalWith()
    sVar_EvalWith_Text = Trim(Combo2.Text)
    
    'with evaluation
    Select Case (sVar_EvalWith_Text)
    
        Case "Equal to":                sVar_EvalWith_Part1 = " = "
                                        sVar_EvalWith_Part2 = " 'value' "
        
        Case "Similar to":              sVar_EvalWith_Part1 = " like "
                                        sVar_EvalWith_Part2 = " '%value%' "
        
        Case "Starts with":             sVar_EvalWith_Part1 = " like "
                                        sVar_EvalWith_Part2 = " 'value%' "
        
        Case "Ends with":               sVar_EvalWith_Part1 = " like "
                                        sVar_EvalWith_Part2 = " '%value' "
        
        Case "Greater than":            sVar_EvalWith_Part1 = " > "
                                        sVar_EvalWith_Part2 = " 'value' "
        
        Case "Equal or Greater than":   sVar_EvalWith_Part1 = " >= "
                                        sVar_EvalWith_Part2 = " 'value' "
        
        Case "Less than":               sVar_EvalWith_Part1 = " < "
                                        sVar_EvalWith_Part2 = " 'value' "
        
        Case "Equal or Less than":      sVar_EvalWith_Part1 = " <= "
                                        sVar_EvalWith_Part2 = " 'value' "
        
        Case Else
                sVar_EvalWith_Part1 = ""
                sVar_EvalWith_Part2 = ""
    End Select
    
    prcBuildQuery
End Sub



Private Sub Form_Load()
    prcInitVars
End Sub

Sub prcInitVars()

    sfrmSearchfor = 1
    Me.Width = 7680
    'Me.Height = 4110
    Me.Height = 6150
    
                bInternalOperation = False
    
    sVar_Field = "Invoice"
    sVar_EvalWith_Part1 = " = "
    sVar_EvalWith_Part2 = "'value' "
    sWhichQuery = " qbx_inv "
    
    
                bInternalOperation = True
                
    Combo1.Text = "Invoice"
    Combo1.AddItem "Invoice"
    Combo1.AddItem "Invoice Balance"
    'Combo1.AddItem "Invoice Date"
    Combo1.AddItem "Total Balance"
    Combo1.AddItem "Phone"
    Combo1.AddItem "Fax"
    Combo1.AddItem "Email"
    Combo1.AddItem "Contact"
    Combo1.AddItem "Alt Contact"
    
    
    
                bInternalOperation = False
                
End Sub

Sub prcEvaluateRequest()
    prcSubmitQuery
End Sub

Sub prcSubmitQuery()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter

On Error GoTo errHandle:

    If rsSearch.State = 1 Then Set rsSearch = Nothing
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then Exit Sub
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = sQuery
        
    Set rsSearch = cmdCommand.Execute
    
        Label2.Caption = rsSearch.RecordCount
        
        If Not rsSearch.EOF Then
            TDBGrid_Search.Columns(3).Caption = sVar_Field_Text
        Else
            TDBGrid_Search.Columns(3).Caption = ""
        End If
    
    TDBGrid_Search.ReBind
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Search Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Sub prcBuildQuery()
    Dim sInterfaceText As String
    Dim sXtraSQL As String
    
    sVar_Value = Trim(Text1.Text)
    
    If sValue_Formatting = "money" Then
        sVar_Value = funFormatDecimal(sVar_Value)
    End If

    sVar_EvalWith_Part2 = Replace(sVar_EvalWith_Part2, "value", sVar_Value)
    
    If sValue_Formatting = "money" Then
        sVar_EvalWith_Part2 = sVar_EvalWith_Part1 & "CONVERT(money, " & sVar_EvalWith_Part2 & ")"
    Else
        sVar_EvalWith_Part2 = sVar_EvalWith_Part1 & sVar_EvalWith_Part2
    End If
    
    
        
    
    If sWhichQuery = " qbx_inv " Then
    
        ''''''''''''''''''''restrictions for qbx_cust''''''''''''''''''
        If sProfileAttrDtlsAry(1, 6) = 0 Then
            sXtraSQL = " and cust.cust_jobstatus <> 'awarded' and cust.cust_accountnumber <> '' "
        End If
        
        If sProfileAttrDtlsAry(1, 10) = 0 Then
            If sProfileAttrDtlsAry(1, 11) = 0 Then
                sXtraSQL = sXtraSQL & " and CONVERT(int, cust.cust_totalbalance_money) = '0' and  CONVERT(int, cust_totalbalance_money) <> '0' "
            Else
                sXtraSQL = sXtraSQL & " and CONVERT(int, cust.cust_totalbalance_money) = '0' "
            End If
        Else
            If sProfileAttrDtlsAry(1, 11) = 0 Then
                sXtraSQL = sXtraSQL & " and CONVERT(int, cust.cust_totalbalance_money) <> '0' "
            End If
        End If
        
        If sProfileAttrDtlsAry(1, 44) = 0 Then
            sXtraSQL = sXtraSQL & " and sign(cust.cust_totalbalance_money) != CONVERT(money, '-1') "
        End If
        ''''''''''''''''''''''end restrictions for qbx_cust''''''''''''''''''''
    
    
        If sVal_type = "string" And sValue_Formatting = "money" Then
    
            sQuery = " select cust.cust_listid, cust.cust_accountnumber, cust.cust_name, inv.inv_refnumber, inv." & sVar_Field & " as fvalue " & _
                    " from qbx_inv inv " & _
                    " left join qbx_cust cust on cust.cust_listid = inv.inv_customerref_listid " & _
                    " where CONVERT(money, " & " inv." & sVar_Field & ")" & sVar_EvalWith_Part2 & sXtraSQL
        
            sInterfaceText = " select cust.cust_listid, cust.cust_accountnumber, cust.cust_name, inv.inv_refnumber, inv." & sVar_Field & " as fvalue " & vbNewLine & _
                    " from qbx_inv inv " & vbNewLine & _
                    " left join qbx_cust cust on cust.cust_listid = inv.inv_customerref_listid " & vbNewLine & _
                    " where CONVERT(money, " & " inv." & sVar_Field & ")" & sVar_EvalWith_Part2 & vbNewLine
                    
        Else
    
            sQuery = " select cust.cust_listid, cust.cust_accountnumber, cust.cust_name, inv.inv_refnumber, inv." & sVar_Field & " as fvalue " & _
                    " from qbx_inv inv " & _
                    " left join qbx_cust cust on cust.cust_listid = inv.inv_customerref_listid " & _
                    " where inv." & sVar_Field & sVar_EvalWith_Part2 & sXtraSQL
        
            sInterfaceText = " select cust.cust_listid, cust.cust_accountnumber, cust.cust_name, inv.inv_refnumber, inv." & sVar_Field & " as fvalue " & vbNewLine & _
                    " from qbx_inv inv " & vbNewLine & _
                    " left join qbx_cust cust on cust.cust_listid = inv.inv_customerref_listid " & vbNewLine & _
                    " where inv." & sVar_Field & sVar_EvalWith_Part2 & vbNewLine
        
        End If
                
    ElseIf sWhichQuery = " qbx_cust " Then
    
        ''''''''''''''''''''restrictions for qbx_cust''''''''''''''''''
        If sProfileAttrDtlsAry(1, 6) = 0 Then
            sXtraSQL = " and cust_jobstatus <> 'awarded' and cust_accountnumber <> '' "
        End If
        
        If sProfileAttrDtlsAry(1, 10) = 0 Then
            If sProfileAttrDtlsAry(1, 11) = 0 Then
                sXtraSQL = sXtraSQL & " and CONVERT(int, cust_totalbalance_money) = '0' and  CONVERT(int, cust_totalbalance_money) <> '0' "
            Else
                sXtraSQL = sXtraSQL & " and CONVERT(int, cust_totalbalance_money) = '0' "
            End If
        Else
            If sProfileAttrDtlsAry(1, 11) = 0 Then
                sXtraSQL = sXtraSQL & " and CONVERT(int, cust_totalbalance_money) <> '0' "
            End If
        End If
        
        If sProfileAttrDtlsAry(1, 44) = 0 Then
            sXtraSQL = sXtraSQL & " and sign(cust_totalbalance_money) != CONVERT(money, '-1') "
        End If
        ''''''''''''''''''''''end restrictions for qbx_cust''''''''''''''''''''
        
        
        sQuery = " select cust_listid, cust_accountnumber, cust_name, cust_index as inv_refnumber, " & sVar_Field & " as fvalue " & _
                " From qbx_cust " & _
                " where " & sVar_Field & sVar_EvalWith_Part2 & sXtraSQL
        
        sInterfaceText = " select cust_listid, cust_accountnumber, cust_name, cust_index as inv_refnumber, " & sVar_Field & " as fvalue " & vbNewLine & _
                " From qbx_cust " & vbNewLine & _
                " where " & sVar_Field & sVar_EvalWith_Part2 & vbNewLine
                
    End If
    
    Text2.Text = sInterfaceText & vbNewLine & vbNewLine & sQuery
    
    prcSubmitQuery
    
End Sub







Private Sub Form_Unload(Cancel As Integer)
    sfrmSearchfor = 0
End Sub

Private Sub TDBGrid_Search_DblClick()
    If TDBGrid_Search.Columns(0).Value <> "" Then
        prcGrabInfo
    End If
End Sub

Sub prcGrabInfo()

On Error Resume Next

    ListID = Trim(TDBGrid_Search.Columns(0).Value)
    'frmInvoiceQry.prcFind Trim(TDBGrid1.Columns(0).value), 0
    frmInvoiceQry.prcCallCustomerDtls
    
    frmInvoiceQry.SSTab1.Tab = 0
    frmInvoiceQry.Command11.Enabled = True
    frmInvoiceQry.Frame2.Enabled = True
    
    frmInvoiceQry.Text7.Text = Trim(TDBGrid_Search.Columns(2).Value)
    frmInvoiceQry.prcFind2 LCase(Trim(frmInvoiceQry.Text7.Text)), Trim(frmInvoiceQry.Combo4.Text), Trim(frmInvoiceQry.Combo5.Text)
    frmInvoiceQry.prcProcessFindingCustomer
    
    Unload Me
End Sub

Private Sub TDBGrid_Search_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
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
            rsSearch.MoveLast
            rsSearch.MoveNext
        Else
            rsSearch.MoveFirst
            rsSearch.MovePrevious
        End If
        rsSearch.Move offset
    Else
        rsSearch.Move offset, StartLocation
    End If
        
    StartRow = rsSearch.Bookmark
    Pos = rsSearch.AbsolutePosition
    
    For Row = 0 To Rows
        If rsSearch.BOF Or rsSearch.EOF Then Exit For
        For col = 0 To cols
            stemp = Trim(rsSearch!fvalue) & ""
            If sValue_Formatting = "money" Then
                If Trim(rsSearch!inv_refnumber) <> "" Then
                    stemp = "Invoice:" & Trim(rsSearch!inv_refnumber) & ", $" & funFormatDecimal(stemp)
                Else
                    stemp = "$" & funFormatDecimal(stemp)
                End If
            End If
            Select Case (col)
                Case (0):   RowBuf.Value(Row, 0) = Trim(rsSearch!cust_listid) & ""
                Case (1):   RowBuf.Value(Row, 1) = Trim(rsSearch!cust_accountnumber) & ""
                Case (2):   RowBuf.Value(Row, 2) = Trim(rsSearch!cust_name) & ""
                Case (3):   RowBuf.Value(Row, 3) = stemp
            End Select
        Next col
        RowBuf.Bookmark(Row) = rsSearch.Bookmark
        RowsFetched = RowsFetched + 1
        rsSearch.MoveNext
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
    If bInternalOperation = False Then
        prcEvalField
        prcEvalWith
    End If
End Sub
