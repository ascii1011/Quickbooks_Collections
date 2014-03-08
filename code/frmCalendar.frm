VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4365
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483638
      Year            =   2004
      Month           =   11
      Day             =   1
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
    With Calendar1
        If iCalendarRequest = 1 Then
            frmCallbacks.Text1.Text = .Value
            frmCallbacks.Command1.Value = True
            Unload Me
        ElseIf iCalendarRequest = 2 Then
            frmInvoiceQry.Text11.Text = .Value
            Unload Me
        ElseIf iCalendarRequest = 3 Then
            frmPrintPage.Text3.Text = Trim(frmPrintPage.Text3.Text) & vbNewLine & .Value
            Unload Me
        End If
    End With
End Sub

Private Sub Calendar1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
    If KeyCode = 120 And Shift = 1 Then
        frmLogin.Show
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 4455
    Me.Height = 3315
    
    Calendar1.ShowDays = True
    Calendar1.DayFont.Size = 8.25
    Calendar1.DayFontColor = &H0&
    
    Calendar1.DayFont.Size = 8.25
    Calendar1.DayFont.Bold = True
    
    Calendar1.GridFont.Size = 8.25
    Calendar1.GridFont.Bold = False
    
    Calendar1.TitleFont.Size = 12
    Calendar1.TitleFont.Bold = True
    
    With Calendar1
        If iCalendarRequest = 1 Then
            .Value = Trim(frmCallbacks.Text1.Text)
        ElseIf iCalendarRequest = 2 Then
            .Value = Now
        ElseIf iCalendarRequest = 3 Then
            .Value = Trim(frmPrintPage.Text9.Text)
        ElseIf iCalendarRequest = 4 Then
            .Value = Trim(frmPrintPage.Text10.Text)
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sFrmCalendar = 0
End Sub

Private Sub Form_Initialize()
    sFrmCalendar = 1
End Sub
