VERSION 5.00
Begin VB.Form Query 
   Caption         =   "Invoice Query"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox IncludeLineItems 
      Caption         =   "Include Line Items"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Submit 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox ToTxnDate 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox FromTxnDate 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "To Txn Date (mm/dd/yyyy)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From Txn Date (mm/dd/yyyy)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   495
      Width           =   2055
   End
End
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: InvoiceQuery
'
' Description:  This sample demonstrates the use of QBFC 2.0,
'               by querying for invoices within a given date range.
'               It includes examples of the following:
'                   - Constructing a complex query
'                   - Looping through the list of invoices in a response
'                   - Getting detailed invoice information (invoice lines)
'                   - Reading data from an OR object
'                   - Checking fields that are not guaranteed in the response
'                     and obtaining data from them
'
' Created On: 01/17/2002
' Updated to QBFC 2.0: 08/2002
'
' Copyright © 2002 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Dim fromDate As String
Dim toDate As String
Dim invoiceCollection As Collection

' These constants hold the application name and ID plus the
' major and minor versions for QBFC 2.0.
Const cAppID = "Desktop VB QBFC InvoiceQuery"
Const cAppName = "Desktop VB QBFC InvoiceQuery"





'
' When the user clicks Submit, validate the data that has been entered
' and send the request to QuickBooks for processing.
'
Private Sub Submit_Click()

On Error GoTo ErrHandler

    If GetDates Then      'Check the validity of the input data
        If QueryInvoices Then   'If the dates are valid, use them in query
            ShowInvoices    'If there are any invoices, show them
        End If
    End If
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub



'
' Make sure input data is valid, in particular the dates.
'
Private Function GetDates()
    
On Error GoTo ErrHandler
    
    fromDate = Format(FromTxnDate.Text, "mm/dd/yyyy")
    toDate = Format(ToTxnDate.Text, "mm/dd/yyyy")
    
    If fromDate <> "" And Not IsDate(fromDate) Then
        MsgBox "From Txn Date format is wrong!"
        GetDates = False
        Exit Function
    End If
    
    If toDate <> "" And Not IsDate(toDate) Then
        MsgBox "To Txn Date format is wrong!"
        GetDates = False
        Exit Function
    End If

    GetDates = True
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    GetDates = False
    Exit Function
    
End Function



'
' Query QuickBooks for invoices between the specified dates
' using QBFC.
'
Private Function QueryInvoices()
    
On Error GoTo ErrHandler
    
    ' Step 1: Start session with QuickBooks
    Dim SessionManager As New QBFC3Lib.QBSessionManager
    SessionManager.OpenConnection cAppID, cAppName
    SessionManager.BeginSession "", QBFC3Lib.omDontCare

    ' Step 2: Create Message Set request
    Dim requestMsgSet As QBFC3Lib.IMsgSetRequest
    Set requestMsgSet = GetLatestMsgSetRequest(SessionManager)
    requestMsgSet.Attributes.OnError = QBFC3Lib.roeContinue
    
    ' Step 3: Create the query object needed to perform InvoiceQueryRq
    Dim invQuery As QBFC3Lib.IInvoiceQuery
    Set invQuery = requestMsgSet.AppendInvoiceQueryRq
    
    Dim invFilter As QBFC3Lib.IInvoiceFilter
    Set invFilter = invQuery.ORInvoiceQuery.InvoiceFilter
    
    ' If there's date info, create the filter and put it in the query
    If fromDate <> "" Or toDate <> "" Then
        If fromDate <> "" Then
            invFilter.ORDateRangeFilter.TxnDateRangeFilter. _
                ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue (CDate(fromDate))
        End If
        If toDate <> "" Then
            invFilter.ORDateRangeFilter.TxnDateRangeFilter. _
                ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue (CDate(toDate))
        End If
    End If
    invQuery.IncludeLineItems.SetValue (Me.IncludeLineItems)
        
    ' Step 4: Do the request
    Dim responseMsgSet As QBFC3Lib.IMsgSetResponse
    Set responseMsgSet = SessionManager.DoRequests(requestMsgSet)
    
    ' Terminate the session and connection,
    ' since we are done with the session manager
    SessionManager.EndSession
    SessionManager.CloseConnection
    
    ' Uncomment the following to see the request and response XML for debugging
    ' MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    ' MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
    
    ' Step 5: Interpret the response
    Dim rsList As QBFC3Lib.IResponseList
    Set rsList = responseMsgSet.ResponseList
    
    Dim response As QBFC3Lib.IResponse
    ' Retrieve the one response corresponding to our single request
    Set response = rsList.GetAt(0)
              
    If (response.StatusCode <> 0) Then
        If (response.StatusCode = 1) Then   ' No record found
            MsgBox "No invoice is found", vbInformation, "Message from QuickBooks"
        Else
            Dim msg
            msg = "Error occured.  Status Code = " & CStr(response.StatusCode) & _
                    ", Status Message = " & response.StatusMessage & _
                    ", Status Severity = " & response.StatusSeverity
            MsgBox msg, vbExclamation, "Message from QuickBooks"
        End If
        QueryInvoices = False
        Exit Function
    Else
        ' We have one or more invoices in the invoice list, which is the response.Detail
        Dim invoiceList As QBFC3Lib.IInvoiceRetList
        Set invoiceList = response.Detail
        Set invoiceCollection = New Collection
        Dim ndx
        For ndx = 0 To (invoiceList.Count - 1)
            Dim invoiceRet As QBFC3Lib.IInvoiceRet
            Set invoiceRet = invoiceList.GetAt(ndx)
            ' Add to the collection
            invoiceCollection.Add invoiceRet, invoiceRet.TxnID.GetValue
        Next
    End If
    
    QueryInvoices = True
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    QueryInvoices = False
    Exit Function
End Function



'
' Display invoices using the Display.frm form.
'
Private Sub ShowInvoices()

    Dim msg As String
    Dim invRet As QBFC3Lib.IInvoiceRet
    Dim i As Integer
    Const cMaxShown = 30    ' Show only a maximum of 30 invoices
    
    If invoiceCollection.Count > cMaxShown Then
        msg = "Showing " & CStr(cMaxShown) & " out of " & _
            CStr(invoiceCollection.Count) & " invoices" & vbCrLf
    End If
    
    i = 1
    For Each invRet In invoiceCollection
        msg = msg & vbCrLf & GetInvoiceRetDetail(invRet)
        If i = cMaxShown Then
            Exit For
        End If
        i = i + 1
    Next
    
    Dim frmDisplay As New Display
    frmDisplay.Text_Content = msg
    frmDisplay.Show vbModal, Me
        
    'Clear the collection
    Set invoiceCollection = Nothing
    Exit Sub
    
End Sub



'
' Retrieve details of an invoice from a particular IInvoiceRet object.
'
Private Function GetInvoiceRetDetail(invRet As QBFC3Lib.IInvoiceRet) As String

    Dim msg
    
    'Retrieve guaranteed fields
    msg = " TxnNumber = " & invRet.TxnNumber.GetValue & ", Customer = " & invRet.CustomerRef.FullName.GetValue
    
    'Retrive non-guaranteed fields
    If (Not (invRet.RefNumber Is Nothing)) Then
        msg = msg & ", RefNumber = " & invRet.RefNumber.GetValue
    End If
    
    If (Not (invRet.Memo Is Nothing)) Then
        msg = msg & ", Memo = " & invRet.RefNumber.GetValue
    End If
    
    'Retrieve invoice line list
    'Each line can be either InvoiceLineRet OR InvoiceLineGroupRet
    Dim orInvoiceLineRetList As QBFC3Lib.IORInvoiceLineRetList
    Set orInvoiceLineRetList = invRet.orInvoiceLineRetList
    If (Not (orInvoiceLineRetList Is Nothing)) Then
    
        Dim linendx, linendxMax
        linendxMax = orInvoiceLineRetList.Count - 1
        For linendx = 0 To linendxMax
            Dim orInvoiceLineRet As QBFC3Lib.IORInvoiceLineRet
            Set orInvoiceLineRet = orInvoiceLineRetList.GetAt(linendx)
            
            msg = msg & vbCrLf & vbTab & " Line: " & CStr(linendx)
            'Check what to retrieve from the orInvoiceLineRet object
            'based on the "ortype" property
            If (orInvoiceLineRet.ortype = QBFC3Lib.orilrInvoiceLineRet) Then
                
                If (Not (orInvoiceLineRet.InvoiceLineRet.Desc Is Nothing)) Then
                    msg = msg & ", Desc: " & orInvoiceLineRet.InvoiceLineRet.Desc.GetValue
                End If
                
                If (Not (orInvoiceLineRet.InvoiceLineRet.Amount Is Nothing)) Then
                    msg = msg & ", Amount: " & orInvoiceLineRet.InvoiceLineRet.Amount.GetValue
                End If
                                
                If (Not (orInvoiceLineRet.InvoiceLineRet.ItemRef Is Nothing)) Then
                    msg = msg & ", Quantity: " & orInvoiceLineRet.InvoiceLineRet.ItemRef.FullName.GetValue
                End If
                
                Dim orRate As QBFC3Lib.IORRate
                Set orRate = orInvoiceLineRet.InvoiceLineRet.orRate
                If (Not (orRate Is Nothing)) Then
                    If orRate.ortype = QBFC3Lib.orrRate Then
                        msg = msg & ", Rate: " & CStr(orRate.Rate.GetValue)
                    Else
                        msg = msg & ", RatePercent: " & CStr(orRate.RatePercent.GetValue)
                    End If
                End If
            
            ElseIf (orInvoiceLineRet.ortype = QBFC3Lib.orilrInvoiceLineGroupRet) Then
                msg = msg & ", Group Name: " & orInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName.GetValue
                msg = msg & ", Total Amount: " & CStr(orInvoiceLineRet.InvoiceLineGroupRet.TotalAmount.GetValue)
                
                If (Not (orInvoiceLineRet.InvoiceLineGroupRet.Desc Is Nothing)) Then
                    msg = msg & ", Desc: " & orInvoiceLineRet.InvoiceLineGroupRet.Desc.GetValue
                End If
                
            End If
        Next
    End If
  
    GetInvoiceRetDetail = msg
End Function



'
' Exit program.
'
Private Sub Exit_Click()
    Unload Me
End Sub



Public Function GetLatestMsgSetRequest(SessionManager As QBSessionManager) As IMsgSetRequest
    Dim supportedVersion As String
    supportedVersion = QBFCLatestVersion(SessionManager)
    If (supportedVersion >= "3.0") Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
    ElseIf (supportedVersion = "2.0") Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
    ElseIf (supportedVersion = "1.1") Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
    End If
End Function

Function QBFCLatestVersion(SessionManager As QBSessionManager) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = SessionManager.QBXMLVersionsForSession
    
    Dim msgset As QBFC3Lib.IMsgSetRequest
    'Use oldest version to ensure that we work with any QuickBooks (US)
    Set msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
    msgset.AppendHostQueryRq
    Dim QueryResponse As QBFC3Lib.IMsgSetResponse
    Set QueryResponse = SessionManager.DoRequests(msgset)
    Dim response As QBFC3Lib.IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single HostQuery request
    Set response = QueryResponse.ResponseList.GetAt(0)
    Dim HostResponse As IHostRet
    Set HostResponse = response.Detail
    Dim supportedVersions As IBSTRList
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.Count - 1
        vers = Val(supportedVersions.GetAt(i))
        If (vers > LastVers) Then
            LastVers = vers
            QBFCLatestVersion = supportedVersions.GetAt(i)
        End If
    Next i
End Function

