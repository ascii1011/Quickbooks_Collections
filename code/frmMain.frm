VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Modern Consumer - Collections Application"
   ClientHeight    =   8505
   ClientLeft      =   1815
   ClientTop       =   2040
   ClientWidth     =   13215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer PriorityAlertTimer 
      Interval        =   60000
      Left            =   300
      Top             =   3360
   End
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13155
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   13215
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   19440
         Left            =   1560
         Picture         =   "frmMain.frx":37D2
         ScaleHeight     =   19440
         ScaleWidth      =   25920
         TabIndex        =   4
         Top             =   600
         Width           =   25920
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2040
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   300
         Width           =   4095
      End
   End
   Begin VB.Timer Timer2 
      Left            =   300
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   300
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   240
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8781
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF33
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B385
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB67
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F441
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68647
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69033
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69145
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6929F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":693F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6950B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6961D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D697
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D7A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71823
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71935
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":759AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75AC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79B3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79C4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79D5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AFE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F4AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F8FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Collections"
            Object.ToolTipText     =   "Collections"
            Object.Tag             =   "mnuEdit_COLLECTIONS"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   "mnuEdit_REFRESH"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reporting"
            Object.ToolTipText     =   "Reporting"
            Object.Tag             =   "mnu_Reporting"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Priority Alerts"
            Object.ToolTipText     =   "Priority Alerts"
            Object.Tag             =   "mnu_PriorityAlerts"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Profile Manager"
            Object.ToolTipText     =   "Profile Manager"
            Object.Tag             =   "mnu_ProfileManager"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            Object.Tag             =   "mnuEdit_Search"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Launch help file"
            Object.Tag             =   "mnuHELP_START"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8250
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "12/13/2006"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:39 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "TSRTime"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7126
            MinWidth        =   7126
            Object.Tag             =   "mgr"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5362
            MinWidth        =   5362
            Object.Tag             =   "refresh"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnu_collections 
         Caption         =   "Collections"
      End
      Begin VB.Menu mnu_callbacks 
         Caption         =   "Callbacks"
      End
      Begin VB.Menu mnu_calendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_Search 
         Caption         =   "Search For"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEDIT_REQUEST_UPDATE 
         Caption         =   "Request Update"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit_REFRESH 
         Caption         =   "Ref&resh"
      End
      Begin VB.Menu mnuEdit_UpdateCollectionsApp 
         Caption         =   "Update Collections Application"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnu_manage_users 
         Caption         =   "&Profile Manager"
      End
      Begin VB.Menu mnuWindowspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_tools_options 
         Caption         =   "Options"
         Begin VB.Menu mnu_tools_options_alert_settings 
            Caption         =   "Alert Settings"
         End
         Begin VB.Menu mnu_tools_options_importance_settings 
            Caption         =   "Importance Settings"
         End
         Begin VB.Menu frmTools_Options_QBFaxes 
            Caption         =   "QuickBooks Faxes"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reporting"
      Enabled         =   0   'False
      Begin VB.Menu mnu_reporting 
         Caption         =   "Re&ports"
      End
   End
   Begin VB.Menu mnuLST_Refresh 
      Caption         =   "Refresh"
      Begin VB.Menu mnu_enable_auto_refresh 
         Caption         =   "Enable Auto Refresh"
      End
      Begin VB.Menu mnu_disable_auto_refresh 
         Caption         =   "Disable Auto Refresh"
      End
   End
   Begin VB.Menu mnuAdmin_Tools 
      Caption         =   "Admin Tools"
      Begin VB.Menu mnuAdmin_reset_view 
         Caption         =   "reset"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnu_ClearWindows 
         Caption         =   "&Clear Windows"
      End
      Begin VB.Menu mnuArrangeWindows 
         Caption         =   "&Arrange Windows"
      End
      Begin VB.Menu mnuWindowspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public cnMC                     As New ADODB.Connection
Attribute cnMC.VB_VarHelpID = -1
Dim f                           'var for file access
Dim lSecurityUpdate As Long     'minutes for security settings to be reloaded
Dim lDatabaseUpdate As Long     'minutes for Database to be updated

'versions
Dim sVersion_Current        As String
Dim sVersion_Current_Name   As String
Dim sVersion_Current_DateTime   As String

Dim sVersion_Latest         As String
Dim sVersion_Latest_Name    As String
Dim sVersion_Latest_DateTime    As String

Dim bUpdateTime As Boolean

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type


Private Sub frmTools_Options_QBFaxes_Click()
    prcQuickBooksFaxes
End Sub

Private Sub MDIForm_Activate()
    iMainLostFocus = 0
    Debug.Print "main active"
End Sub

Private Sub MDIForm_Deactivate()
    iMainLostFocus = 1
    Debug.Print "main deactive"
End Sub

Private Sub MDIForm_Initialize()
    Debug.Print "main init"
End Sub

Private Sub MDIForm_LinkClose()
    Debug.Print "main link close"
End Sub

Private Sub MDIForm_Load()
    
    prcDisabled
    Timer1.Enabled = False
    Timer2.Enabled = False
    frmMain.StatusBar1.Panels.Item(5).Text = "Auto Refresh - Off"
    iUpdateTries = 0
    lSecurityUpdate = 3
    lDatabaseUpdate = 15
    iInvoiceQryLostFocus = 1
    iMainLostFocus = 0
    sKillPriorityAlert = False
        
    frmMain.Caption = "Modern Consumer - Collections Application - Version " & App.Major & "." & App.Minor & "." & App.Revision
        
    prcGetSysInfo
    
    'sUser = "vdelgado"
    'sUser = "njaureguy"
    'sUser = "icedeno"
    'sUser = "jreyes"
    
    prcSecurityAndUpdates
    If bUpdateTime = True Then Unload Me
    
    'frmQuickBooksFaxes.Show
End Sub

Sub prcSecurityAndUpdates()
    Dim strIP As String
    Dim strName As String
    Dim Response
    'strIP = Winsock1.LocalIP         'Captures IP Address and stores it
    'strName = Winsock1.LocalHostName 'Captures Host Name and stores
    Dim sVersionResponse As String
    
    bUpdateTime = False
    
    prcDoRegSet
    prcMCConn
    
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State = 1 Then
    
        prcGetUserInfo
        
        'check if certain DLLs are installed, if not.. install and register
        prcDepends
        
        'check if updater needs to be updated
        prcCheckForUpdatedUpdater
        
        'check version and upgrade if needed
        'If funGrabCurrentRegVersions = 0 Then
            sVersionResponse = funGrabProductUpdatesVersions
            If sVersionResponse = 0 Then
                If sVersion_Latest <> sVersion_Current Then
                    ''''frmProcessing.Text1.Text = "Updating 'Update Manager' for Collections."
                    ''''frmProcessing.Show
                    'Dim Response
                    'Response = MsgBox("There is an update available for" & vbNewLine & _
                    '        "Modcon Collections Application." & vbNewLine & vbNewLine & _
                    '        "Your current is version:" & sVersion_Current & vbNewLine & _
                    '        "The update version is:" & sVersion_Latest & vbNewLine & _
                    '        "Please make sure that the 'update version' is higher then your 'current version'." & vbNewLine & vbNewLine & _
                    '        "Do you want to update now?", vbExclamation + vbYesNo, "Update Confirmation!")
                            
                    'If Response = vbYes Then
                        'aryRep_Profile(6)
    '                    "C:\Program Files\Modcon\Collections\Updater\collections_updater.exe"
                        'Call Shell("Z:\QB\collections\Updater\Collections_Updater.exe", vbNormalFocus)
                        
                    'End If
                    
                    
                    
                    'so updates don't happen while debugging
                    If sUser = "charty" Then
                        Response = MsgBox("Do you wish to update?", vbExclamation + vbYesNo, "Update the 'Updater' Error")
                        If Response = vbYes Then
                            bUpdateTime = True
                            If aryRep_Profile(6) = "0" Then
                                Call Shell("Z:\QB\collections\Updater\Collections_Updater.exe", vbNormalFocus)
                            Else
                                Call Shell("C:\Program Files\Modcon\Collections\Updater\collections_updater.exe", vbNormalFocus)
                            End If
                        End If
                    Else
                        bUpdateTime = True
                        If aryRep_Profile(6) = "0" Then
                            Call Shell("Z:\QB\collections\Updater\Collections_Updater.exe", vbNormalFocus)
                        Else
                            Call Shell("C:\Program Files\Modcon\Collections\Updater\collections_updater.exe", vbNormalFocus)
                        End If
                    End If
                        
                        
                        
                End If
            ElseIf sVersionResponse = 2 Then
                MsgBox "Errors: Unable to acquire update information at this time." & vbNewLine & vbNewLine & _
                        "You may continue to work."
            End If
        'Else
        '    MsgBox "Errors: Modcon Collections Registry Keys Missing!!!!"
        '    Unload Me
        'End If
        
        If bUpdateTime = False Then
        
            prcPresetProfileAttributes
            
            If funGrabProfile = 1 Then
            
                prcGrabProfileDtls
                prcInitFormVars
                prcGrabQbxProperties
                funGrabImportLevels
                funGrabRegions
                
                Me.Enabled = True
                Me.WindowState = 2
                frmMain.StatusBar1.Panels.Item(3).Text = sUser
                funCheckUpdateMgr
                            
                If sUser = "vdelgado" Then
                    frmPriorityAlerts.Show
                Else
                    sKillPriorityAlert = True
                End If
                
                'check for which forms will be started up
                'prcShowFrmInvoiceQry
                'frmEmailImport.Show
            Else
                MsgBox "The current user does not have permission to access this program."
                Unload Me
            End If
            
        End If
    Else
        MsgBox "An Internal Error has occured, please call chris (concerning: '0x0D008' sql)"
    End If
End Sub

Sub prcDepends()
    Dim sFileAry(5) As String
    Dim sPath As String, stemp As String, sRegTemp As String
    Dim i As Integer
    Dim sCopyFrom As String, sCopyTo As String
    
On Error GoTo DependErr:
    
    sFileAry(0) = "tdbgpp7.dll"
    sFileAry(1) = "TABCTL32.OCX"
    sFileAry(2) = "tdbg7.ocx"
    sFileAry(3) = "MSCAL.OCX"
    sFileAry(4) = "Msflxgrd.ocx"
    
    sPath = "c:\" & sGRootDir & "\system32\"
    
    For i = 0 To UBound(sFileAry) - 1
        stemp = sPath & sFileAry(i)
        
        If DoesFileExist(stemp) = False Then
            sCopyFrom = "Z:\QB\Components\" & sFileAry(i)
            sCopyTo = "c:\" & sGRootDir & "\system32\"
            If funCopyFiles(sCopyFrom, sCopyTo) = True Then
                sRegTemp = "regsvr32.exe /s " & stemp
                Call Shell(sRegTemp, vbNormalFocus)
            End If
        End If
    Next i
    Exit Sub
    
DependErr:
    MsgBox "prcDepends-" & Err.Number & ": " & vbNewLine & Err.Description
End Sub


Function funCopyFiles(sSource As String, sDest As String) As Boolean
    Dim fs, f

On Error GoTo CopyErr:
    funCopyFiles = False
    
    Screen.MousePointer = vbHourglass
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set f = fs.GetFile(sSource)
    f.Copy sDest
    
    funCopyFiles = True
    Screen.MousePointer = vbDefault
    Exit Function
        
CopyErr:
        Screen.MousePointer = vbDefault
        MsgBox "funCopyFiles-" & vbNewLine & sSource & vbNewLine & sDest & vbNewLine & Err.Number & ": " & vbNewLine & Err.Description
        Exit Function
End Function

''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
Sub prcCheckForUpdatedUpdater()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsUpdaterChk    As New ADODB.Recordset
    Dim sErrMsg         As String
    Dim sUpdateMsg      As String
    Dim sNewVersion     As String
    Dim sSourceLoc      As String
    Dim sDestinationLoc      As String

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then Exit Sub
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_update_list where update_name = 'Updater' and update_active = '1'  "
            
    Set rsUpdaterChk = cmdCommand.Execute
    
    If Not rsUpdaterChk.EOF Then
        rsUpdaterChk.MoveFirst
        
        sNewVersion = Trim(rsUpdaterChk!update_version) & ""
        sSourceLoc = LCase(Trim(rsUpdaterChk!update_location)) & ""
        sDestinationLoc = LCase(Trim(rsUpdaterChk!update_destination)) & ""
        
        If aryRep_Profile(6) = "0" Or aryRep_Profile(6) < sNewVersion Then
            
            If DoesFileExist(sSourceLoc) = True Then
                sErrMsg = funCopyFile(sSourceLoc, sDestinationLoc)
                If sErrMsg = 9 Then
                    sUpdateMsg = sNewVersion & ":OK"
                Else
                    sUpdateMsg = aryRep_Profile(6) & ":Err(" & sErrMsg & ")->, ver;" & sNewVersion
                End If
                prcUpdatedUpdaterVersion4User sUpdateMsg
            End If
            
        End If
        
    End If
            
    Set rsUpdaterChk = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Update the 'Updater' Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsUpdaterChk = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Sub prcUpdatedUpdaterVersion4User(sMsg As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsUpdaterChk    As New ADODB.Recordset

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then Exit Sub
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " update qb_profiles " & _
                            " set profiles_last_updater_version = '" & sMsg & "' " & _
                            " where profiles_username = '" & sUser & "'  "
            
    cmdCommand.Execute
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Update the 'Updater' version for " & sUser & " Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Function funCopyFile(sSource As String, sDest As String) As Integer
    Dim sSourcePath As String
    Dim sDestinationPath As String
    Dim fs, f, s

On Error GoTo CopyErr:

    'MsgBox sSource & ", " & sDest
    
    Screen.MousePointer = vbHourglass
    
    funCopyFile = funCopyFile + 1
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    funCopyFile = funCopyFile + 3
    'MsgBox sDest
    Set f = fs.GetFile(sSource)
    f.Copy sDest
    
    funCopyFile = funCopyFile + 5
    Screen.MousePointer = vbDefault
    Exit Function
        
CopyErr:
        funCopyFile = funCopyFile + 11
        Screen.MousePointer = vbDefault
        
        Exit Function
End Function

Function funCopyFile22(sSourcePath As String, sDestinationPath As String) As Boolean
    Dim sError

    Dim fso As New Scripting.FileSystemObject
    Dim f As Object
    
On Error GoTo errorOpenFile:

    'funCopyFile = False
    
    fso.CopyFile sSourcePath, sDestinationPath, True
    
    'funCopyFile = True
    
    Set f = Nothing
    Exit Function
    
errorOpenFile:
    MsgBox Err.Number & " " & Err.Description
    'funCopyFile = False
    Set f = Nothing
End Function

Function funGrabImportLevels() As String
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsImportLvls    As New ADODB.Recordset
    Dim i As Integer

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then Exit Function
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_importance_levels order by import_id asc "
            
    Set rsImportLvls = cmdCommand.Execute
    
    If Not rsImportLvls.EOF Then
        i = 0
        ReDim aryGImportLvl(rsImportLvls.RecordCount, 2)
        rsImportLvls.MoveFirst
        While Not rsImportLvls.EOF
            aryGImportLvl(i, 0) = Trim(rsImportLvls!import_active) & ""
            aryGImportLvl(i, 1) = Trim(rsImportLvls!import_id) & ""
            aryGImportLvl(i, 2) = Trim(rsImportLvls!import_name) & ""
            i = i + 1
            rsImportLvls.MoveNext
        Wend
    Else
        ReDim aryGImportLvl(0)
    End If
            
    Set rsImportLvls = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "funGrabImportLevels Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsImportLvls = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function


Function funGrabRegions() As String
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsRegions    As New ADODB.Recordset
    Dim i As Integer

On Error GoTo errHandle:
    
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then Exit Function
            
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_regions order by reg_index asc "
            
    Set rsRegions = cmdCommand.Execute
    
    If Not rsRegions.EOF Then
        ReDim aryGRegions(rsRegions.RecordCount + 1, 1)
        aryGRegions(0, 0) = "All"
        i = 1
        rsRegions.MoveFirst
        While Not rsRegions.EOF
            aryGRegions(i, 0) = Trim(rsRegions!reg_name) & ""
            i = i + 1
            rsRegions.MoveNext
        Wend
    Else
        ReDim aryGRegions(0)
    End If
            
    Set rsRegions = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "funGrabImportLevels Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsRegions = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function


Sub prcGetUserInfo()
    Dim Response
    Dim cmdCommand          As New ADODB.Command
    Dim parParameter        As New ADODB.Parameter
    Dim rsFullUserName      As New ADODB.Recordset
    Dim sBreakSplit
        
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qb_profiles where profiles_username = '" & sUser & "' "
        
    Set rsFullUserName = cmdCommand.Execute
            
    If Not rsFullUserName.EOF Then
        rsFullUserName.MoveFirst
        aryRep_Profile(0) = Trim(rsFullUserName!profiles_username) & ""
        aryRep_Profile(1) = Trim(rsFullUserName!profiles_full) & ""
        aryRep_Profile(2) = Trim(rsFullUserName!profiles_fullname) & ""
        aryRep_Profile(3) = Trim(rsFullUserName!profiles_email) & ""
        aryRep_Profile(4) = Trim(rsFullUserName!profiles_phone) & ""
        aryRep_Profile(5) = Trim(rsFullUserName!profiles_fax) & ""
        sBreakSplit = Split(Trim(rsFullUserName!profiles_last_updater_version), ":")
        If UBound(sBreakSplit) > 0 Then
            aryRep_Profile(6) = sBreakSplit(0)
        Else
            aryRep_Profile(6) = "0"
        End If
        sBreakSplit = Split(Trim(rsFullUserName!profiles_last_collection_version), ":")
        If UBound(sBreakSplit) > 0 Then
            aryRep_Profile(7) = sBreakSplit(0)
        Else
            aryRep_Profile(7) = "0"
        End If
        sVersion_Current = aryRep_Profile(7)
    End If
    
    Set rsFullUserName = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
        'frmLoginProgress.List1.AddItem "User Profile error: " & err.number & " - " & err.Description
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "prcGetUserInfo Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsFullUserName = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Sub

Sub prcDisabled()
    'frmMain.mnuTools
    frmMain.tbToolbar.Buttons.Item(11).Visible = False
    frmMain.tbToolbar.Buttons.Item(11).Enabled = False
    frmMain.tbToolbar.Buttons.Item(10).Visible = False
    frmMain.tbToolbar.Buttons.Item(10).Enabled = False
    frmMain.tbToolbar.Buttons.Item(9).Visible = False
    frmMain.tbToolbar.Buttons.Item(9).Enabled = False
    frmMain.tbToolbar.Buttons.Item(7).Visible = False
    frmMain.tbToolbar.Buttons.Item(7).Enabled = False
    frmMain.tbToolbar.Buttons.Item(5).Visible = False
    frmMain.tbToolbar.Buttons.Item(5).Enabled = False
    frmMain.tbToolbar.Buttons.Item(3).Visible = False
    frmMain.tbToolbar.Buttons.Item(3).Enabled = False
    frmMain.tbToolbar.Buttons.Item(1).Visible = False
    frmMain.tbToolbar.Buttons.Item(1).Enabled = False
    
    'frmMain.tbToolbar.Visible = False
    'frmMain.tbToolbar.Enabled = False
    
    'frmMain.mnuFile.Visible = False
    'frmMain.mnuFile.Enabled = False
        frmMain.mnu_callbacks.Enabled = False    'callbacks
        frmMain.mnu_callbacks.Visible = False
        frmMain.mnu_collections.Enabled = False    'collections
        frmMain.mnu_collections.Visible = False
    frmMain.mnuEdit.Visible = False
    frmMain.mnuEdit.Enabled = False
    frmMain.mnuTools.Visible = False
    frmMain.mnuTools.Enabled = False
        frmMain.mnu_manage_users.Enabled = False
        frmMain.mnu_manage_users.Visible = False
        frmMain.mnu_tools_options.Visible = False
        frmMain.mnu_tools_options.Enabled = False
    frmMain.mnuReports.Enabled = False
    frmMain.mnuReports.Visible = False
        frmMain.mnuReports.Enabled = False
        frmMain.mnuReports.Visible = False
        frmMain.mnuLST_Refresh.Visible = False
        frmMain.mnuAdmin_Tools.Visible = False
    
    'frmMain.mnuWindow.Visible = False
    'frmMain.mnuWindow.Enabled = False
End Sub

Sub prcEnabledView()
        'frmMain.tbToolbar.Visible = True
        'frmMain.tbToolbar.Enabled = True
        'frmMain.mnuFile.Enabled = True
        'frmMain.mnuFile.Visible = True
        'frmMain.mnuWindow.Enabled = True
        'frmMain.mnuWindow.Visible = True
        
    If sProfileAttrDtlsAry(1, 2) = 1 And sProfileAttrDtlsAry(1, 1) = 1 Then
        frmMain.mnu_callbacks.Enabled = True    'callbacks
        frmMain.mnu_callbacks.Visible = True
    End If
    
    If sProfileAttrDtlsAry(1, 1) = 1 Then
        frmMain.tbToolbar.Buttons.Item(1).Visible = True   'collection
        frmMain.tbToolbar.Buttons.Item(1).Enabled = True
        frmMain.mnu_collections.Enabled = True
        frmMain.mnu_collections.Visible = True
    End If
    
    If sProfileAttrDtlsAry(1, 25) = 1 Then
        frmMain.tbToolbar.Buttons.Item(3).Visible = True    'refresh
        frmMain.tbToolbar.Buttons.Item(3).Enabled = True
        frmMain.mnuEdit.Enabled = True
        frmMain.mnuEdit.Visible = True
    End If
    
    If sProfileAttrDtlsAry(1, 4) = 1 Then
        frmMain.tbToolbar.Buttons.Item(5).Visible = True   'reporting
        frmMain.tbToolbar.Buttons.Item(5).Enabled = True
        frmMain.mnuReports.Enabled = True
        frmMain.mnuReports.Visible = True
        frmMain.mnuTools.Enabled = True
        frmMain.mnuTools.Visible = True
    End If
    
    If sProfileAttrDtlsAry(1, 5) = 1 Then
        frmMain.tbToolbar.Buttons.Item(9).Visible = True    'manage
        frmMain.tbToolbar.Buttons.Item(9).Enabled = True
        frmMain.mnuTools.Enabled = True
        frmMain.mnuTools.Visible = True
        frmMain.mnu_manage_users.Enabled = True
        frmMain.mnu_manage_users.Visible = True
        frmMain.mnu_tools_options.Visible = True
        frmMain.mnu_tools_options.Enabled = True
    End If
    
    frmMain.mnuLST_Refresh.Visible = True
    frmMain.mnu_disable_auto_refresh.Visible = True
    frmMain.mnu_enable_auto_refresh.Visible = False
    'If sProfileAttrDtlsAry(1, 27) = 1 Then
    '    frmMain.tbToolbar.Buttons.Item(9).Visible = True   'search
    '    frmMain.tbToolbar.Buttons.Item(9).Enabled = True
    'End If
    
    'If sProfileAttrDtlsAry(1, 28) = 1 Then
    '    frmMain.tbToolbar.Buttons.Item(10).Visible = True    'help
    '    frmMain.tbToolbar.Buttons.Item(10).Enabled = True
    'End If
End Sub


Sub prclevel_1()
    prcEnabledView
    frmMain.tbToolbar.Buttons.Item(9).Visible = True    'manage
    frmMain.tbToolbar.Buttons.Item(9).Enabled = True
    frmMain.tbToolbar.Buttons.Item(7).Visible = True   'alert
    frmMain.tbToolbar.Buttons.Item(7).Enabled = True
    frmMain.mnuTools.Visible = True
    frmMain.mnuTools.Enabled = True
        frmMain.mnu_manage_users.Enabled = True
        frmMain.mnu_manage_users.Visible = True
        frmMain.mnu_tools_options.Visible = True
        frmMain.mnu_tools_options.Enabled = True
    frmMain.mnuAdmin_Tools.Visible = True
End Sub


Sub prclevel_2()
    prcEnabledView
    frmMain.tbToolbar.Buttons.Item(9).Visible = True    'manage
    frmMain.tbToolbar.Buttons.Item(9).Enabled = True
    frmMain.tbToolbar.Buttons.Item(7).Visible = True   'alert
    frmMain.tbToolbar.Buttons.Item(7).Enabled = True
    frmMain.mnu_tools_options.Visible = True
    frmMain.mnu_tools_options.Enabled = True
    frmMain.mnuTools.Visible = True
    frmMain.mnuTools.Enabled = True
End Sub


Sub prclevel_3()
    prcEnabledView
    frmMain.tbToolbar.Buttons.Item(7).Visible = True   'alert
    frmMain.tbToolbar.Buttons.Item(7).Enabled = True
    frmMain.mnu_tools_options.Visible = True
    frmMain.mnu_tools_options.Enabled = True
End Sub


Sub prclevel_4()
    prcEnabledView
    frmMain.tbToolbar.Buttons.Item(7).Visible = True   'alert
    frmMain.tbToolbar.Buttons.Item(7).Enabled = True
End Sub


Sub prclevel_5()
    prcEnabledView
End Sub

'Sub prcInitVars()
    'sProfileAttrNamesAry(0) = "Enable 'Collections' Window"
    'sProfileAttrNamesAry(1) = "Enable 'Callbacks' Window"
    'sProfileAttrNamesAry(2) = "Enable 'Print Page' Window"
    'sProfileAttrNamesAry(3) = "Enable 'Print Reports' Window"
    'sProfileAttrNamesAry(4) = "Enable 'Profile Managers' Window"
    'sProfileAttrNamesAry(5) = "Show 'Media' companies"
    'sProfileAttrNamesAry(6) = "Show companies with a zero balance"
    'sProfileAttrNamesAry(7) = "Able to view 'Customer Messages'"
    'sProfileAttrNamesAry(8) = "Able to Save 'Customer Messages'"
    'sProfileAttrNamesAry(9) = "Able to view 'Open Invoices'"
    'sProfileAttrNamesAry(10) = "Able to view 'Closed Invoices'"
    'sProfileAttrNamesAry(11) = "Number of Days, Months, or Years"
    'sProfileAttrNamesAry(12) = "Type: Days, Months, or Years"
    'sProfileAttrNamesAry(13) = "Be able to view Payments"
    'sProfileAttrNamesAry(14) = "Number of Days, Months, or Years"
    'sProfileAttrNamesAry(15) = "Type: Days, Months, or Years"
    'sProfileAttrNamesAry(16) = "Be able to see the 'Email a Customer' button"
    'sProfileAttrNamesAry(17) = "Be able to change 'Email a Customer' info"
    'sProfileAttrNamesAry(18) = "Be able to see the 'Alert' button"
    'sProfileAttrNamesAry(19) = "Be able to see the 'Print' button"
    'sProfileAttrNamesAry(20) = "Be able to print an 'Invoice' letter"
    'sProfileAttrNamesAry(21) = "Be able to print a 'Collection' letter"
    'sProfileAttrNamesAry(22) = "Be able to see the 'Print Report' button"
    'sProfileAttrNamesAry(23) = "Be able to see the Toolbar 'Collections' button"
    'sProfileAttrNamesAry(24) = "Be able to see the Toolbar 'Refresh' button"
    'sProfileAttrNamesAry(25) = "Be able to see the Toolbar 'Profile Manager' button"
    'sProfileAttrNamesAry(26) = "Be able to see the Toolbar 'Search' button"
    'sProfileAttrNamesAry(27) = "Be able to see the Toolbar 'Help' button"
    'sProfileAttrNamesAry(28) = "Be able to see the Toolbar 'Print Report' button"
    'sProfileAttrNamesAry(29) = "Be able to see the Menu 'Collections' button"
    'sProfileAttrNamesAry(30) = "Be able to see the Menu 'Calendar' button"
    'sProfileAttrNamesAry(31) = "Be able to see the Menu 'Callbacks' button"
    'sProfileAttrNamesAry(32) = "Be able to see the Menu 'Refresh' button"
    'sProfileAttrNamesAry(33) = "Be able to see the Menu 'Profile Manager' button"
    'sProfileAttrNamesAry(34) = "Be able to see the Menu 'Clear Windows' button"
    'sProfileAttrNamesAry(35) = "Be able to see the Menu 'Print Report' button"
    'sProfileAttrNamesAry(36) = "Be able to access 'Calendar' windows"
    'sProfileAttrNamesAry(37) = "Be able to access 'Callbacks' windows"
    'sProfileAttrNamesAry(38) = "Be able to access 'Email Customers' windows"
    'sProfileAttrNamesAry(39) = "Be able to access 'Collections' windows"
    'sProfileAttrNamesAry(40) = "Be able to access 'Print Page' windows"
    'sProfileAttrNamesAry(41) = "Be able to access 'Print Report' windows"


'End Sub

Sub prcPresetProfileAttributes()
    iEnableInvoiceQryToStartUp = 0
    iEnableCallbacksToStartUp = 0
    iEnablePrintPageToStartUp = 0
    iEnablePrintReportsToStartUp = 0
    iEnableManageUsersToStartUp = 0
    iShowMedia = 0
    iShowZeroBalance = 0
    iEnableViewingOfMsgs = 0
    iEnableSaveMsgs = 0
    iEnableViewInvoicesOpen = 0
    iEnableViewInvoicesClosed = 0
    iViewInvoicesNumber = 0
    iViewInvoicesDateType = 0
    iEnableViewPayements = 0
    iViewPaymentsNumber = 0
    iViewPaymentsDateType = 0
    iEnableEmailCustomer = 0
    iEnableEditEmailCustomer = 0
    iEnableAlertButton = 0
    iEnablePrintButton = 0
    iEnablePrintInvoicesletter = 0
    iEnablePrintCollectionLetter = 0
    iEnablePrintReportbutton = 0
    iEnableToolbarCollections = 0
    iEnableToolbarRefresh = 0
    iEnableToolbarManageUsers = 0
    iEnableToolbarSearch = 0
    iEnableToolbarHelp = 0
    iEnableToolbarPrintReport = 0
    iEnableDropCollections = 0
    iEnableDropCalendar = 0
    iEnableDropCallbacks = 0
    iEnableDropRefresh = 0
    iEnableDropManageUsers = 0
    iEnableDropClearWindows = 0
    iEnableDropPrintReport = 0
    iEnableAccessCalendar = 0
    iEnableAccessCallbacks = 0
    iEnableAccessEmailCustomer = 0
    iEnableAccessInvoiceqry = 0
    iEnableAccessPrintPage = 0
    iEnableAccessPrintReport = 0
    iEnableAccessQueryFilter = 0
End Sub


'insert_reg_changes_sp
Private Sub prcMCConn()

On Error Resume Next
    Dim UserName As String, serverName As String, DatabaseName As String, Password As String
               
    cnMC.Provider = "sqloledb"
    cnMC.CursorLocation = adUseClient
    cnMC.Properties("Data Source").Value = sMCServer
    cnMC.Properties("Initial Catalog").Value = sMCDatabase
    cnMC.Properties("User ID").Value = sMCUsername
    cnMC.Properties("Password").Value = sMCPassword
    cnMC.Open
    
    If cnMC.State = 1 Then
    End If
        'List1.AddItem "SQL connected..."
        'Timer2.Enabled = False
        'bConConfirm = True
        'Unload frmProcessing
    'Else
    '    'List1.AddItem "Trying to connect."
    '    serverName = "nao2"
    '    DatabaseName = "modcon"
    '    UserName = "sa"
    '    Password = "123456"
    '    cnMC.Provider = "sqloledb"
    '    cnMC.CursorLocation = adUseClient
    '    cnMC.Properties("Data Source").value = serverName
    '    cnMC.Properties("Initial Catalog").value = DatabaseName
    '    cnMC.Properties("User ID").value = UserName
    '    cnMC.Properties("Password").value = Password
    '    cnMC.Open
    '    'Unload frmProcessing
    '    If cnMC.State <> 1 Then
    '        MsgBox "SQL Not Connected!!!, Contact Chris H."
    '    End If
    'End If
    
End Sub



Function funCheckUpdateMgr() As Integer
    Dim Response
    Dim cmdCommand          As New ADODB.Command
    Dim parParameter        As New ADODB.Parameter
    Dim rsUpdateMgr         As New ADODB.Recordset
    Dim dTempDate           As Date
    
    funCheckUpdateMgr = 0
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qb_profiles order by profiles_update_request_handled_datestamp desc "
        
    Set rsUpdateMgr = cmdCommand.Execute
            
    If Not rsUpdateMgr.EOF Then
    
        rsUpdateMgr.MoveFirst
        'If Not rsUpdateMgr.EOF Then
        '    frmMain.StatusBar1.Panels.Item(4).Text = "last updated: " & Trim(rsUpdateMgr!profiles_update_request_handled_datestamp) & ""
        'Else
        '    frmMain.StatusBar1.Panels.Item(4).Text = ""
        'End If
                
        While Not rsUpdateMgr.EOF
            If "qbupdate" = Trim(rsUpdateMgr!profiles_username) Then
                frmMain.StatusBar1.Panels.Item(4).Text = "last updated: " & Trim(rsUpdateMgr!profiles_update_request_handled_datestamp) & ""
            End If
            If sUser = Trim(rsUpdateMgr!profiles_username) Then
                If IsNull(Trim(rsUpdateMgr!profiles_update_request_handled_datestamp)) Then
                    frmMain.mnuEDIT_REQUEST_UPDATE = True
                Else
                    dTempDate = DateAdd("n", 15, Trim(rsUpdateMgr!profiles_update_request_handled_datestamp))
                    If Now > dTempDate Then
                        'enable request again
                        frmMain.mnuEDIT_REQUEST_UPDATE.Enabled = True
                    End If
                End If
            End If
            rsUpdateMgr.MoveNext
        Wend
        
        
        
    Else
        frmMain.StatusBar1.Panels.Item(4).Text = ""
    End If
        
    Set rsUpdateMgr = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "funCheckUpdateMgr Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsUpdateMgr = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function


Function funCheckMgr(fForm As Form) As Integer
    Dim iCheckValue As Integer
    Dim i As Integer


    i = 0
    iCheckValue = funCheckUpdateMgr
            
    If iCheckValue = 1 Then
        frmProcessing.Text1.Text = "QuickBooks is being updated, please wait...."
        frmProcessing.Text1.Refresh
    End If
            
    While iCheckValue = 1 And i < 18
        Pause 6
        frmProcessing.Text1.Text = "Retrying attempt " & i & " of 18, please wait...."
        frmProcessing.Text1.Refresh
        iCheckValue = funCheckUpdateMgr
        i = i + 1
        frmProcessing.Text1.Text = "QuickBooks is still being updated, please wait...."
        frmProcessing.Text1.Refresh
    Wend
    
    If iCheckValue = 1 Then
        MsgBox "The Modcon Quickbooks Update Manager is taking too long to update information." & vbNewLine & "Please try again a couple minutes."
        Unload fForm
        funCheckMgr = 1
    ElseIf iCheckValue = 2 Then
        funCheckMgr = 2
    Else
        frmProcessing.Text1.Text = "Updated, please wait...."
        frmProcessing.Text1.Refresh
        funCheckMgr = 0
    End If
        
    iCheckValue = 0
    i = 0
End Function


Function funGrabProfile() As Integer
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsProfile     As New ADODB.Recordset
    
    funGrabProfile = 0
    prcDisabled
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_profile_sp"
    
    'reg_list_user
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, LCase(Trim(sUser)) & "")
    cmdCommand.Parameters.Append parParameter
        
    Set rsProfile = cmdCommand.Execute
    
    'frmLoginProgress.List1.AddItem "EOF: " & rsProfile.EOF
    If Not rsProfile.EOF Then
        Dim j As Integer
        
        'frmLoginProgress.List1.AddItem "In EOF"
        rsProfile.MoveFirst
        
            'username from db
            sProfileAttrDtlsAry(0, 0) = Trim(rsProfile!profiles_username)
            
            'security lvl
            sSecLvl = Trim(rsProfile!profiles_level)
            
            sSecEnabled = Trim(rsProfile!profiles_enabled)
            
            sFullName = Trim(rsProfile!profiles_fullname)
            
            
            'check for validation or info that will be used later
            If sProfileAttrDtlsAry(0, 0) = "" Or sSecLvl = "" Then
                Exit Function
            End If
                        
            'populate array with user profile attributes
            sProfileAttrDtlsAry(1, 1) = Trim(rsProfile![1])
            sProfileAttrDtlsAry(1, 2) = Trim(rsProfile![2])
            sProfileAttrDtlsAry(1, 3) = Trim(rsProfile![3])
            sProfileAttrDtlsAry(1, 4) = Trim(rsProfile![4])
            sProfileAttrDtlsAry(1, 5) = Trim(rsProfile![5])
            sProfileAttrDtlsAry(1, 6) = Trim(rsProfile![6])
            sProfileAttrDtlsAry(1, 7) = Trim(rsProfile![7])
            sProfileAttrDtlsAry(1, 8) = Trim(rsProfile![8])
            sProfileAttrDtlsAry(1, 9) = Trim(rsProfile![9])
            sProfileAttrDtlsAry(1, 10) = Trim(rsProfile![10])
            sProfileAttrDtlsAry(1, 11) = Trim(rsProfile![11])
            sProfileAttrDtlsAry(1, 12) = Trim(rsProfile![12])
            If Trim(rsProfile![13]) <> "days" And Trim(rsProfile![13]) <> "months" And Trim(rsProfile![13]) <> "years" Then
                sProfileAttrDtlsAry(1, 13) = "years"
            Else
                sProfileAttrDtlsAry(1, 13) = Trim(rsProfile![13])
            End If
            sProfileAttrDtlsAry(1, 14) = Trim(rsProfile![14])
            sProfileAttrDtlsAry(1, 15) = Trim(rsProfile![15])
            If Trim(rsProfile![16]) <> "days" And Trim(rsProfile![16]) <> "months" And Trim(rsProfile![16]) <> "years" Then
                sProfileAttrDtlsAry(1, 16) = "years"
            Else
                sProfileAttrDtlsAry(1, 16) = Trim(rsProfile![16])
            End If
            sProfileAttrDtlsAry(1, 17) = Trim(rsProfile![17])
            sProfileAttrDtlsAry(1, 18) = Trim(rsProfile![18])
            sProfileAttrDtlsAry(1, 19) = Trim(rsProfile![19])
            sProfileAttrDtlsAry(1, 20) = Trim(rsProfile![20])
            sProfileAttrDtlsAry(1, 21) = Trim(rsProfile![21])
            sProfileAttrDtlsAry(1, 22) = Trim(rsProfile![22])
            sProfileAttrDtlsAry(1, 23) = Trim(rsProfile![23])
            sProfileAttrDtlsAry(1, 24) = Trim(rsProfile![24])
            sProfileAttrDtlsAry(1, 25) = Trim(rsProfile![25])
            sProfileAttrDtlsAry(1, 26) = Trim(rsProfile![26])
            sProfileAttrDtlsAry(1, 27) = Trim(rsProfile![27])
            sProfileAttrDtlsAry(1, 28) = Trim(rsProfile![28])
            sProfileAttrDtlsAry(1, 29) = Trim(rsProfile![29])
            sProfileAttrDtlsAry(1, 30) = Trim(rsProfile![30])
            sProfileAttrDtlsAry(1, 31) = Trim(rsProfile![31])
            sProfileAttrDtlsAry(1, 32) = Trim(rsProfile![32])
            sProfileAttrDtlsAry(1, 33) = Trim(rsProfile![33])
            sProfileAttrDtlsAry(1, 34) = Trim(rsProfile![34])
            sProfileAttrDtlsAry(1, 35) = Trim(rsProfile![35])
            sProfileAttrDtlsAry(1, 36) = Trim(rsProfile![36])
            sProfileAttrDtlsAry(1, 37) = Trim(rsProfile![37])
            sProfileAttrDtlsAry(1, 38) = Trim(rsProfile![38])
            sProfileAttrDtlsAry(1, 39) = Trim(rsProfile![39])
            sProfileAttrDtlsAry(1, 40) = Trim(rsProfile![40])
            sProfileAttrDtlsAry(1, 41) = Trim(rsProfile![41])
            sProfileAttrDtlsAry(1, 42) = Trim(rsProfile![42])
            sProfileAttrDtlsAry(1, 43) = Trim(rsProfile![43])
            sProfileAttrDtlsAry(1, 44) = Trim(rsProfile![44])
            sProfileAttrDtlsAry(1, 45) = Trim(rsProfile![45])
            sProfileAttrDtlsAry(1, 46) = Trim(rsProfile![46])
            
            prcSetSecurityLevel
            
            'index the array
            For j = 1 To iProfileCount
                sProfileAttrDtlsAry(0, j) = j
            Next
            
            Timer1.Enabled = True
            frmMain.StatusBar1.Panels.Item(5).Text = "Auto Refresh - On"
            Timer1.Interval = 60000
            funGrabProfile = 1
    Else
        'frmLoginProgress.List1.AddItem "No User Info Found."
    End If
    
    Set rsProfile = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Function
    
errHandle:
        'frmLoginProgress.List1.AddItem "User Profile error: " & err.number & " - " & err.Description
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "funGrabProfile Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsProfile = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function


Sub prcSetSecurityLevel()
            
    If sSecLvl = 1 Then
        frmMain.prclevel_1
    End If
            
    If sSecLvl = 2 Then
        frmMain.prclevel_2
    End If
            
    If sSecLvl = 3 Then
        frmMain.prclevel_3
    End If
            
    If sSecLvl = 4 Then
        frmMain.prclevel_4
    End If
            
    If sSecLvl = 5 Then
        frmMain.prclevel_5
    End If
    
End Sub



Function prcGrabProfileDtls() As Integer
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsProfileDtls     As New ADODB.Recordset
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "grab_profile_dtls_sp"
        
    Set rsProfileDtls = cmdCommand.Execute
    
    If Not rsProfileDtls.EOF Then
        
        rsProfileDtls.MoveFirst
        While Not rsProfileDtls.EOF
            sProfileAttrDtlsAry(2, Trim(rsProfileDtls!attr_index)) = Trim(rsProfileDtls!attr_name)
            sProfileAttrDtlsAry(3, Trim(rsProfileDtls!attr_index)) = Trim(rsProfileDtls!attr_desc)
            sProfileAttrDtlsAry(4, Trim(rsProfileDtls!attr_index)) = Trim(rsProfileDtls!attr_enabled)
            rsProfileDtls.MoveNext
        Wend
 
    Else
        'List1.AddItem "No saved settings found, user input required."
    End If
    
    Set rsProfileDtls = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Screen.MousePointer = vbDefault
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "prcGrabProfileDtls Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsProfileDtls = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
End Function




Sub prcInitFormVars()
    sFrmReporting = 0
    sFrmProfiles = 0
    sFrmCalendar = 0
    sFrmCallbacks = 0
    sFrmEmailCustomer = 0
    sFrmInvoiceQry = 0
    sFrmPrintPage = 0
    sFrmProcessing = 0
    sFrmLogin = 0
    sFrmSecurityLock = 0
    sFrmQueryFilter = 0
    sFrmRemarks = 0
    sFrmPriorityAlerts = 0
    sfrmSearchfor = 0
    sfrmImportanceSettings = 0
    sfrmQuickBooksFaxes = 0
End Sub

Sub prcQuickBooksFaxes()
    If sfrmQuickBooksFaxes = 0 Then
        Load frmQuickBooksFaxes
        frmQuickBooksFaxes.Show
    Else
        MsgBox "The 'QuickBooks Faxes' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcfrmImportanceSettings()
    If sfrmImportanceSettings = 0 Then
        Load frmImportanceSettings
        frmImportanceSettings.Show
    Else
        MsgBox "The 'Importance Settings' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcShowSearchFor()
    If sfrmSearchfor = 0 Then
        Load frmSearchFor
        frmSearchFor.Show
    Else
        MsgBox "The 'Search For' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcShowPriorityAlerts()
    If sFrmPriorityAlerts = 0 Then
        Load frmPriorityAlerts
        frmPriorityAlerts.Show
    Else
        MsgBox "The 'Priority Alerts' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcShowemailimport()
    If sFrmEmailImport = 0 Then
        Load frmEmailImport
        frmEmailImport.Show
    Else
        MsgBox "The 'Customer Stats' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcShowFrmQBInStats()
    If sFrmQBInStats = 0 Then
        Load frmQBInStats
        frmQBInStats.Show
    Else
        MsgBox "The 'Customer Stats' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub


Sub prcShowFrmReporting()
    If sFrmReporting = 0 Then
        Load frmReporting
        frmReporting.Show
    Else
        MsgBox "The 'Reporting' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcShowFrmProfiles()
    If sFrmProfiles = 0 Then
        Load frmProfiles
        frmProfiles.Show
    Else
        MsgBox "The 'Profiles' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    
    End If
End Sub

Sub prcShowFrmProcessing()
    If sFrmProcessing = 0 Then
        Load frmProcessing
        frmProcessing.Show
    End If
End Sub

Sub prcShowFrmPrintPage()
    If sFrmPrintPage = 0 Then
        Load frmPrintPage
        frmPrintPage.Show
    Else
        MsgBox "The 'Print Page' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    End If
End Sub

Sub prcShowFrmEmailCustomer()
    If sFrmEmailCustomer = 0 Then
        Load frmEmailCustomer
        frmEmailCustomer.Show
    Else
        MsgBox "The 'Email Customer' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    End If
End Sub

Sub prcShowFrmRemarks()
    If sFrmRemarks = 0 Then
        Load frmRemarks
        frmRemarks.Show
    Else
        MsgBox "The 'Quick Notes' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    End If
End Sub
Sub prcShowFrmCallbacks()
    If sFrmCallbacks = 0 Then
        Load frmCallbacks
        frmCallbacks.Show
    Else
        MsgBox "The 'Callbacks' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    End If
End Sub

Sub prcShowFrmInvoiceQry()
    If sFrmInvoiceQry = 0 Then
        Load frmInvoiceQry
        frmInvoiceQry.Show
    Else
        MsgBox "The 'Collections' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    End If
End Sub

Sub prcShowFrmCalendar(ByVal who As Integer)
    If sFrmCalendar = 0 Then
        iCalendarRequest = who
        Load frmCalendar
        frmCalendar.Show
    Else
        MsgBox "The 'Calendar' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
    End If
End Sub

Sub prcUpdateFormSec(iOption As Integer)

    If sFrmInvoiceQry = 1 Then
        'print page
        If sProfileAttrDtlsAry(1, 3) = 1 Then
            frmInvoiceQry.Command4.Visible = True
        Else
            frmInvoiceQry.Command4.Visible = False
        End If
        'callback
        If sProfileAttrDtlsAry(1, 2) = 1 Then
            frmInvoiceQry.Command5.Visible = True
        Else
            frmInvoiceQry.Command5.Visible = False
        End If
        'email customer
        If sProfileAttrDtlsAry(1, 17) = 1 Then
            frmInvoiceQry.Picture2.Visible = True
        Else
            frmInvoiceQry.Picture2.Visible = False
        End If
        'email alert
        If sProfileAttrDtlsAry(1, 19) = 1 Then
            frmInvoiceQry.Picture3.Visible = True
        Else
            frmInvoiceQry.Picture3.Visible = False
        End If
        'View query filter settings
        If sProfileAttrDtlsAry(1, 43) = 1 Then
            frmInvoiceQry.Command6.Visible = True
        Else
            frmInvoiceQry.Command6.Visible = False
        End If
        If sProfileAttrDtlsAry(1, 43) = 0 And sFrmQueryFilter = 1 Then
            Unload frmQueryFilter
        End If
        'collections window
        If sProfileAttrDtlsAry(1, 1) = 0 Then
            Unload frmInvoiceQry
        End If
        If iOption = 1 Then
            frmInvoiceQry.prcRefresh
        End If
    End If
    
    If sFrmCallbacks = 1 Then
        'collections window
        If sProfileAttrDtlsAry(1, 1) = 0 Or sProfileAttrDtlsAry(1, 2) = 0 Then
            Unload frmCallbacks
        End If
    End If
    If sFrmCalendar = 1 Then
        
    End If
    If sFrmEmailCustomer = 1 Then
        
    End If
    If sFrmPrintPage = 1 Then
        
    End If
    If sFrmReporting = 1 Then
        
    End If
    If sFrmQueryFilter = 1 Then
        If sProfileAttrDtlsAry(1, 43) = 0 Then
            Unload frmQueryFilter
        End If
    End If
    'If sFrmSecurityLock <> 0 Then
    'End If
    'If sFrmLogin <> 0 Then
    'End If
    'If sFrmProfiles <> 0 Then
    'End If
    'If sFrmProcessing <> 0 Then
    'End If
    'If sFrmRemarks <> 0 Then
    'End If


End Sub
    
Sub prcClearWindows()
    
    If sFrmRemarks <> 0 Then
        Unload frmRemarks
    End If
    If sFrmInvoiceQry <> 0 Then
        Unload frmInvoiceQry
    End If
    If sFrmCallbacks <> 0 Then
        Unload frmCallbacks
    End If
    If sFrmCalendar <> 0 Then
        Unload frmCalendar
    End If
    If sFrmProcessing <> 0 Then
        Unload frmProcessing
    End If
    If sFrmEmailCustomer <> 0 Then
        Unload frmEmailCustomer
    End If
    If sFrmPrintPage <> 0 Then
        Unload frmPrintPage
    End If
    If sFrmProfiles <> 0 Then
        Unload frmProfiles
    End If
    If sFrmReporting <> 0 Then
        Unload frmReporting
    End If
    If sFrmSecurityLock <> 0 Then
        Unload frmSecurityLock
    End If
    If sFrmLogin <> 0 Then
        Unload frmLogin
    End If
    If sFrmQueryFilter <> 0 Then
        Unload frmQueryFilter
    End If
    If sFrmPriorityAlerts <> 0 Then
        Unload frmPriorityAlerts
    End If
    If sfrmSearchfor <> 0 Then
        Unload frmSearchFor
    End If
    If sfrmQuickBooksFaxes <> 0 Then
        Unload frmQuickBooksFaxes
    End If
    If sfrmImportanceSettings <> 0 Then
        Unload frmImportanceSettings
    End If
    
    
    'Unload Form1
    prcInitFormVars
End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    sKillPriorityAlert = True
    prcClearWindows
    cnMC.Close
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    
    'If sFrmLogin = 1 Then
    '    frmLogin.Left = frmMain.Width - 4400
    '    frmLogin.top = 100
    'End If
   '
   ' If sFrmSecurityLock = 1 Then
   '     frmSecurityLock.Left = frmMain.Width - 1120
   '     frmSecurityLock.top = frmMain.Height - 2350
   ' End If
    
    If Me.WindowState = 1 Then
        bFocus = True
    Else
        bFocus = False
    End If
    
    Dim client_rect As RECT
    Dim client_hwnd As Long '''''

    

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    '''''''''''''''''''' Copy the original picture into picStretched.
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
    
    ''''''''''''''''''''' Set the MDI form's picture.
    Picture = picStretched.Image '

    ''''''''''''''''''''''''' Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", _
        vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1

End Sub

Private Sub MDIForm_Terminate()
    'sKillPriorityAlert = True
    'prcClearWindows
    'cnMC.Close
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'sKillPriorityAlert = True
    'prcClearWindows
    'cnMC.Close
End Sub



Private Sub mnu_about_Click()
    frmAbout.Show
End Sub

Private Sub mnu_calendar_Click()
    prcShowFrmCalendar 0
End Sub

Private Sub mnu_callbacks_Click()
    prcShowFrmCallbacks
End Sub

Private Sub mnu_ClearWindows_Click()
    prcClearWindows
End Sub

Private Sub mnu_collections_Click()
    prcShowFrmInvoiceQry
End Sub

Private Sub mnu_cust_stats_Click()
    prcShowFrmQBInStats
End Sub

Private Sub mnu_disable_auto_refresh_Click()
    Timer1.Enabled = False
    frmMain.mnu_enable_auto_refresh.Visible = True
    frmMain.mnu_disable_auto_refresh.Visible = False
    frmMain.StatusBar1.Panels.Item(5).Text = "Auto Refresh - Off"
End Sub

Private Sub mnu_enable_auto_refresh_Click()
    Timer1.Enabled = True
    frmMain.mnu_disable_auto_refresh.Visible = True
    frmMain.mnu_enable_auto_refresh.Visible = False
    frmMain.StatusBar1.Panels.Item(5).Text = "Auto Refresh - On"
End Sub

Private Sub mnu_manage_users_Click()
    prcShowFrmProfiles
End Sub

Private Sub mnu_reporting_Click()
    prcShowFrmReporting
End Sub



Private Sub mnu_tools_options_alert_settings_Click()
    frmAlertSettings.Show
End Sub

Private Sub mnu_tools_options_importance_settings_Click()
    frmImportanceSettings.Show
End Sub

Private Sub mnuAdmin_reset_view_Click()
    frmAdmin_Reset.Show
End Sub

Private Sub mnuEDIT_REQUEST_UPDATE_Click()
    frmMain.mnuEDIT_REQUEST_UPDATE.Enabled = False
    prcRequestAUpdate
End Sub

Sub prcRequestAUpdate()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sQuery As String

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then Exit Sub

    sQuery = "update qb_profiles " & _
            " set profiles_update_request = '1', " & _
            " profiles_update_request_datestamp = '" & Now & "' " & _
            " where profiles_username = '" & sUser & "' "
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
        
    cmdCommand.CommandText = sQuery
    
    cmdCommand.Execute
    
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:

        Screen.MousePointer = vbDefault
        'prcMainLog Now & "-prcLogout-Error: " & err.Description & "~" & err.number & ", agent: " & frmMain.strUserTsr
        Resume Next
        
End Sub

Private Sub mnuEdit_Search_Click()
    
    frmSearchFor.Show
End Sub

Private Sub mnuEdit_UpdateCollectionsApp_Click()
    Call Shell("Z:\QB\collections\Updater\Collections_Updater.exe", vbNormalFocus)
    'Call Shell("C:\PROGRAM FILES\Modcon\Collections\Updater\Collections_Updater.exe", vbNormalFocus)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Sub ErrorLog(sFilename As String, sMessage As String, Err As Error)
    
    f = FreeFile
        Open "C:\TraceFile.txt" For Append As #f
        Print #f, Now & " Err = " & Err.Description & " " & Err.Number & " ; User:, Message:" & sMessage
        Close #f
        Screen.MousePointer = vbDefault
        
End Sub

Sub MessageLog(sMessage As String)
    
    f = FreeFile
        Open "C:\MessageLog.txt" For Append As #f
        Print #f, Now & " By " & sUser & vbNewLine & sMessage
        Close #f
        Screen.MousePointer = vbDefault
        
End Sub


Private Sub mnuArrangeWindows_Click()
   
    MousePointer = vbHourglass
    Arrange vbArrangeIcons
    MousePointer = vbDefault

End Sub

Private Sub mnuTileHorizontal_Click()
    MousePointer = vbHourglass
    Arrange vbTileHorizontal
    MousePointer = vbDefault
End Sub


Private Sub mnuTileVertical_Click()
   MousePointer = vbHourglass
    Arrange vbTileVertical
    MousePointer = vbDefault
End Sub


Private Sub mnuTransactions_Click()
    'frmTransactions.GRID = 0
    'frmTransactions.Show
End Sub

Private Sub PriorityAlertTimer_Timer()
    If Format(Now, "m") = 0 Then
        'prcFindPriorityAlerts
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response

On Error GoTo errHandle
    
    Select Case (Button.Key)
        Case ("Collections"): prcShowFrmInvoiceQry
        Case ("Callbacks"):   prcShowFrmCallbacks
        Case ("Calendar"):   prcShowFrmCalendar 0
        Case ("Refresh"):   ActiveForm.prcRefresh
        Case ("Profile Manager"):   prcShowFrmProfiles
        Case ("Priority Alerts"):   prcShowPriorityAlerts
        Case ("Reporting"): prcShowFrmReporting
        'Case ("Tools"):   'prcOrderForm
        'Case ("Search"):  'prcVerifier
        'Case ("Help"):
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
            MsgBox ActiveForm.Caption & " does not support this operation", vbInformation, "SYSTEM"
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Login run time error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub


Sub prcKillfrmInvoiceQry()
    Unload frmInvoiceQry
End Sub

Private Sub Timer1_Timer()
    Debug.Print "timer1"
    'And iMainLostFocus = 0
    
    If SQL_ReConnect_old(cnMC) = False Then
        frmMain.StatusBar1.Panels.Item(6).Text = "Not Connected."
        Exit Sub
    End If
    frmMain.StatusBar1.Panels.Item(6).Text = "Connected."
    
    If iInvoiceQryLostFocus = 0 Then
        If funGrabProfile = 0 Then
            MsgBox "You do not have sufficient priviledges."
            Unload Me
        End If
        prcUpdateFormSec 1
        funCheckUpdateMgr
        prcGrabQbxProperties
    End If
End Sub



Sub prcUpdateOneAttr(sNum As String, sVal As String, sProfile As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim sQuery As String

On Error GoTo errHandle:

    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State = 0 Then Exit Sub

    sQuery = "update qb_features " & _
            " set [" & sNum & "]='" & sVal & "' " & _
            " where features_index = (select profiles_index " & _
                    " From qb_profiles" & _
                    " where profiles_username='" & sProfile & "')"
    'Debug.Print sQuery
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
        
    cmdCommand.CommandText = sQuery
    
    cmdCommand.Execute
            
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:

        Screen.MousePointer = vbDefault
        'prcMainLog Now & "-prcLogout-Error: " & err.Description & "~" & err.number & ", agent: " & frmMain.strUserTsr
        Resume Next
        
End Sub


Sub prcUpdateProfile(sProfile As String, sValue As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
        
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "update_profile_sp"
    
    'profile
    Set parParameter = cmdCommand.CreateParameter(, adVarChar, adParamInput, 50, Trim(sProfile) & "")
    cmdCommand.Parameters.Append parParameter
    
    'value
    Set parParameter = cmdCommand.CreateParameter(, adInteger, adParamInput, , Trim(sValue) & "")
    cmdCommand.Parameters.Append parParameter
        
    cmdCommand.Execute
            
    Screen.MousePointer = vbDefault
    Set cmdCommand = Nothing
    Exit Sub
    
errHandle:

        Screen.MousePointer = vbDefault
        'prcMainLog Now & "-prcLogout-Error: " & err.Description & "~" & err.number & ", agent: " & frmMain.strUserTsr
        Resume Next
        
End Sub



Sub prcGrabQbxProperties()
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsGrabQbxProperty   As New ADODB.Recordset
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If frmMain.cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = frmMain.cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_properties "
        
    Set rsGrabQbxProperty = cmdCommand.Execute
    
    If Not rsGrabQbxProperty.EOF Then
        rsGrabQbxProperty.MoveFirst
        While Not rsGrabQbxProperty.EOF
            If Trim(rsGrabQbxProperty!property_name) = "html_printing" Then
                sGHtml_printing = Trim(rsGrabQbxProperty!property_value) & ""
            ElseIf Trim(rsGrabQbxProperty!property_name) = "image_modcon_c" Then
                sGImage_Modcon_C = Trim(rsGrabQbxProperty!property_value) & ""
            ElseIf Trim(rsGrabQbxProperty!property_name) = "html_reporting" Then
                sGHtml_Reporting = Trim(rsGrabQbxProperty!property_value) & ""
            ElseIf Trim(rsGrabQbxProperty!property_name) = "html_dealer_status" Then
                sGHtml_Dealer_Status = Trim(rsGrabQbxProperty!property_value) & ""
            ElseIf Trim(rsGrabQbxProperty!property_name) = "link_cmc" Then
                sGLink_Cmc = Trim(rsGrabQbxProperty!property_value) & ""
            End If
                
            rsGrabQbxProperty.MoveNext
        Wend
    End If
    
    Set rsGrabQbxProperty = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "prcGrabQbxProperties Error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
    Set rsGrabQbxProperty = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub




Sub prcDoRegSet()
    Dim vResult
    Dim sResult1
    Dim sResult2
    
On Error Resume Next
    
    'sResult1 = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "version")
    'Text1.Text = sResult1
    'sResult1 = Trim(Text1.Text)
    sResult2 = App.Major & "." & App.Minor & "." & App.Revision

        '''''''''''''''sys info'''''''''''''''
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main"
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "version", App.Major & "." & App.Minor & "." & App.Revision, REG_SZ
        'If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "version") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "version", App.Major & "." & App.Minor & "." & App.Revision, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "name") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "name", "empty", REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "datetime") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "datetime", Date, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "LastUser") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "LastUser", sUser, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "Computer") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "Computer", sComputer, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "SysInfo") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "SysInfo", sOS & " " & sOSBuild, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "IP") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "IP", strIP, REG_SZ
        
End Sub



Function funGrabProductUpdatesVersions() As String
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    Dim rsGrabVersions  As New ADODB.Recordset
    
On Error GoTo errHandle:

    funGrabProductUpdatesVersions = 0
    
    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Function
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " select * from qbx_update_list where update_name = 'Collections' and update_active = '1' "
        
    Set rsGrabVersions = cmdCommand.Execute
    
    If Not rsGrabVersions.EOF Then
        rsGrabVersions.MoveFirst
        sVersion_Latest = Trim(rsGrabVersions!update_version) & ""
        sVersion_Latest_Name = Trim(rsGrabVersions!update_name) & ""
        sVersion_Latest_DateTime = Trim(rsGrabVersions!update_datetime) & ""
        
        If sVersion_Latest = "" And sVersion_Latest_Name = "" And sVersion_Latest_DateTime = "" Then
            funGrabProductUpdatesVersions = 2
        End If
    Else
        funGrabProductUpdatesVersions = 1
    End If
       
    Set rsGrabVersions = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Function
    
errHandle:
    Select Case (Err.Number)
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "funGrabProductUpdatesVersions Error")
            If Response = vbYes Then Resume Else Exit Function
    End Select
    Set rsGrabVersions = Nothing
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Function


Function funGrabCurrentRegVersions() As String
    Dim sCheck As String
    
    sCheck = 0
On Error GoTo RegError:

    sVersion_Current = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "version")
    sVersion_Current = Replace(sVersion_Current, vbNullChar, "")
    'Text1.Text = sVersion_Current
    'sVersion_Current = Trim(Text1.Text)
    
    sVersion_Current_Name = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "name")
    sVersion_Current_Name = Replace(sVersion_Current_Name, vbNullChar, "")
    'Text1.Text = sVersion_Current_Name
    'sVersion_Current_Name = Trim(Text1.Text)
    
    sVersion_Current_DateTime = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Collections\main", "datetime")
    sVersion_Current_DateTime = Replace(sVersion_Current_DateTime, vbNullChar, "")
    'Text1.Text = sVersion_Current_DateTime
    'sVersion_Current_DateTime = Trim(Text1.Text)
    
    If sVersion_Current_DateTime = "" And sVersion_Current_Name = "" And sVersion_Current = "" Then
        sCheck = 1
    End If
    
    funGrabCurrentRegVersions = sCheck
    Exit Function
    
RegError:
    'Label4.Caption = "Error while checking the registry."
    funGrabCurrentRegVersions = 2
End Function






Sub prcLogIt(who As String, msg As String)
    Dim Response
    Dim cmdCommand      As New ADODB.Command
    Dim parParameter    As New ADODB.Parameter
    
On Error GoTo errHandle:

    MousePointer = vbHourglass
    SQL_ReConnect_old frmMain.cnMC
    If cnMC.State <> 1 Then
        Exit Sub
    End If
    
    Set cmdCommand.ActiveConnection = cnMC
    cmdCommand.CommandType = adCmdText
    cmdCommand.CommandText = " insert into qbx_log " & _
                            " ( log_datetime, log_by_who, log_msg ) " & _
                            " values " & _
                            " ( '" & Now & "', '" & who & "', '" & Left(msg, 290) & "' ) "
        
    cmdCommand.Execute
        
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Set parParameter = Nothing
    Set cmdCommand = Nothing
    Screen.MousePointer = vbDefault
End Sub

