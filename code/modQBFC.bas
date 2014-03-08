Attribute VB_Name = "modQBFC"
'----------------------------------------------------------
' Copyright © 2003 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Public objSessionManager As New QBFC3Lib.QBSessionManager

'Set the max versions based on the version of QBFC we're using.  If we
'were to run against QB 2004 the supported version would be 3.0, but
'QBFC2_1 can only understand up to version 2.1, so we use these constants
'to make sure we don't try to create version 3.x messages with a QBFC that
'can't create them
Const MaxMajorVersion = 2
Const MaxMinorVersion = 1

Dim booSessionBegun As Boolean

Dim intMajorVersion As Integer
Dim intMinorVersion As Integer

'Declare global variables for keeping the Invoice Query Response object
'and the Invoice Ret object so we can access them later
Dim objInvoiceMsgSetResponse As QBFC3Lib.IMsgSetResponse
Dim objSavedInvoiceRet As QBFC3Lib.IInvoiceRet

Dim strSavedRequestCode As String


Public Sub QBFC_OpenConnectionBeginSession()

  On Error GoTo Errs
  
  objSessionManager.OpenConnection cAppID, cAppName
  objSessionManager.BeginSession companyFile, omMultiUser
  'Debug.Print companyFile
  booSessionBegun = True
  Exit Sub

Errs:
  If err.number = &H80040416 Then
    MsgBox "You must have QuickBooks running with a company" & vbCrLf & _
           "file open to use this program."
    objSessionManager.CloseConnection
    End
  ElseIf err.number = &H80040422 Then
    MsgBox "This QuickBooks company file is open in single user mode and" & vbCrLf & _
           "another application is already accessing it.  Please exit the" & vbCrLf & _
           "other application and run this application again."
    objSessionManager.CloseConnection
    End
  ElseIf err.number = &H80040308 Then
    MsgBox _
      "It appears that the qbXML Request Processor has not" & vbCrLf & _
      "been installed, indicating QuickBooks 2002 or later" & vbCrLf & _
      "may not have been installed.  Please run this sample" & vbCrLf & _
      "after installing QuickBooks 2003 and running the Upgrade."
  ElseIf err.number = &H8007007E Then
'    If QBFC2CA_IsntInstalled Then
      MsgBox _
        "QBFC2 isn't installed.  You need QBFC2 or QBFC2CA installed to" & vbCrLf & _
        "use QBFC with this sample program."
      End
'    Else
'      QBFCCA_OpenConnectionBeginSession
'      SetImplementation "QBFCCA"
'    End If
  Else
    'Debug.Print err.number
    If err.number = "-2147220470" Then
        If sUser = "mfishman" Or sUser = "icedeno" Then
            objSessionManager.CloseConnection
            companyFile = "c:\temp\Modern Consumer.QBW"
            QBFC_OpenConnectionBeginSession
            
        End If
    End If
    'MsgBox "QBFC_OpenConnectionBeginSession" & vbCrLf &
      'err.number & vbCrLf & Hex(err.number) & vbCrLf & _
      'err.Description
    
  End If
End Sub


Public Sub QBFC_EndSessionCloseConnection()

  If booSessionBegun Then
    objSessionManager.EndSession
    objSessionManager.CloseConnection
  End If
End Sub

Function QBFC_MaxVersionSupported() As String

Dim strVersions() As String

  strVersions = objSessionManager.QBXMLVersionsForSession
  QBFC_MaxVersionSupported = strVersions(UBound(strVersions))
  
  If InStr(1, strVersions(UBound(strVersions)), "CA") Then
    'We're in the rare situation where QBFC2 and QBFC2CA are both
    'installed and we're running against QBCA.  End session, close our
    'connection and open and begin with QBFCCA
'    QBFC_EndSessionCloseConnection
'    SetImplementation "QBFCCA"
'    QBFCCA_OpenConnectionBeginSession
'    QBFC_MaxVersionSupported = QBFCCA_MaxVersionSupported
'    Exit Function

' For now exit the program if we're dealing with the Canadian version of
' QB
    MsgBox "The Canadian version of QBFC does not support version 2.1 " & _
      "messages.  Exiting."
    End
  End If
  
  'Now make sure that the version of QBFC installed supports the
  'maximum version of qbXML that QuickBooks can handle
  Dim intQBFCMajorVersion As Integer
  Dim intQBFCMinorVersion As Integer
  Dim enumReleaseLevel As QBFC3Lib.ENReleaseLevel
  Dim intReleaseNumber As Integer
  
  objSessionManager.GetVersion _
    intQBFCMajorVersion, intQBFCMinorVersion, _
    enumReleaseLevel, intReleaseNumber
  
  intMajorVersion = Int(Left(strVersions(UBound(strVersions)), 1))
  intMinorVersion = Int(Right(strVersions(UBound(strVersions)), 1))
  
  If intMajorVersion > MaxMajorVersion Then
    intMajorVersion = MaxMajorVersion
    intMinorVersion = MaxMinorVersion
  End If
  
  QBFC_MaxVersionSupported = Trim(Str(intMajorVersion)) & "." & _
                             Trim(Str(intMinorVersion))
End Function


Public Sub QBFC_FillComboBox(cmbComboBox As ComboBox, _
                             strQueryType As String, _
                             strNameElement As String, _
                             strFilter As String, _
                             booMarkGroupItems As Boolean)

  'Clear the combo box
  cmbComboBox.Clear
  
  Dim strSplits() As String
  strSplits = Split(strQueryType, ",")
  
  Dim strNameElementSplits() As String
  strNameElementSplits = Split(strNameElement, ",")
  
  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest
  Dim objMsgSetResponse As QBFC3Lib.IMsgSetResponse
  Dim objResponse As QBFC3Lib.IResponse
  
  Dim objQuery
  Dim objRetList
  Dim objRet
      
  Dim numItems As Integer
  
  Dim i As Integer
  Dim j As Integer
  For i = 0 To UBound(strSplits)
  
    Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
    objMsgSetRequest.Attributes.OnError = roeContinue
    
    Select Case strSplits(i)
      Case Is = "Account"
        Set objQuery = objMsgSetRequest.AppendAccountQueryRq
      Case Is = "Class"
        Set objQuery = objMsgSetRequest.AppendClassQueryRq
      Case Is = "Customer"
        Set objQuery = objMsgSetRequest.AppendCustomerQueryRq
      Case Is = "CustomerMsg"
        Set objQuery = objMsgSetRequest.AppendCustomerMsgQueryRq
      Case Is = "ItemService"
        Set objQuery = objMsgSetRequest.AppendItemServiceQueryRq
      Case Is = "ItemInventory"
        Set objQuery = objMsgSetRequest.AppendItemInventoryQueryRq
      Case Is = "ItemInventoryAssembly"
        Set objQuery = objMsgSetRequest.AppendItemInventoryAssemblyQueryRq
      Case Is = "ItemNonInventory"
        Set objQuery = objMsgSetRequest.AppendItemNonInventoryQueryRq
      Case Is = "ItemOtherCharge"
        Set objQuery = objMsgSetRequest.AppendItemOtherChargeQueryRq
      Case Is = "ItemSubtotal"
        Set objQuery = objMsgSetRequest.AppendItemSubtotalQueryRq
      Case Is = "ItemGroup"
        Set objQuery = objMsgSetRequest.AppendItemGroupQueryRq
      Case Is = "ItemDiscount"
        Set objQuery = objMsgSetRequest.AppendItemDiscountQueryRq
      Case Is = "ItemPayment"
        Set objQuery = objMsgSetRequest.AppendItemPaymentQueryRq
      Case Is = "ItemSalesTax"
        Set objQuery = objMsgSetRequest.AppendItemSalesTaxQueryRq
      Case Is = "ItemSalesTaxGroup"
        Set objQuery = objMsgSetRequest.AppendItemSalesTaxGroupQueryRq
      Case Is = "SalesRep"
        Set objQuery = objMsgSetRequest.AppendSalesRepQueryRq
      Case Is = "SalesTaxCode"
        Set objQuery = objMsgSetRequest.AppendSalesTaxCodeQueryRq
      Case Is = "ShipMethod"
        Set objQuery = objMsgSetRequest.AppendShipMethodQueryRq
      Case Is = "StandardTerms"
        Set objQuery = objMsgSetRequest.AppendStandardTermsQueryRq
      Case Else
        MsgBox "Unknown type " & strSplits(i) & " passed to QBFC_FillComboBox"
    End Select
    
    If strFilter <> Empty Then
      Dim strTemp As String
      Dim strStartTag As String
      Dim strEndTag As String
      Dim strValue As String
      Dim intTagLength As Integer
      
      strTemp = strFilter
      Do While strTemp <> Empty
        'GetTags strTemp, strStartTag, strEndTag, intTagLength
        
        strValue = Left(strTemp, InStr(1, strTemp, strEndTag) - 1)
        strValue = Right(strValue, Len(strValue) - intTagLength)
        strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, strEndTag) + intTagLength))
      
        Select Case strStartTag
          Case Is = "<AccountType>"
            objQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.AddAsString strValue
          Case Else
            MsgBox "Unknown filter " & strStartTag & " in QBFC_FillComboBox"
        End Select
      Loop
    End If
    
    Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
    
    Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
    
    If objResponse.statusCode <> 0 Then
      If objResponse.statusCode <> 1 Then
        MsgBox "Status Code " & objResponse.statusCode & _
               " on call to QBFC_FillComboBox" & vbCrLf & _
               " for " & strSplits(i) & " list items"
      End If
    Else
    
      Set objRetList = objResponse.Detail
      numItems = objRetList.Count
  
      If UBound(strNameElementSplits) > 0 Then
        strNameElement = strNameElementSplits(i)
      End If
    
      For j = 0 To numItems - 1
        Set objRet = objRetList.GetAt(j)
        If strSplits(i) = "ItemGroup" And booMarkGroupItems Then
          cmbComboBox.AddItem objRet.Name.GetValue & _
            " - Group Item"
        Else
          If strNameElement = "FullName" Then
            cmbComboBox.AddItem objRet.FullName.GetValue
          Else
            cmbComboBox.AddItem objRet.Name.GetValue
          End If
        End If
      Next j
    End If ' If objResponse.StatusCode <> 0
  Next i
End Sub

Public Sub QBFC_FillInvoiceList(FG2 As MSFlexGrid, _
                                strRefNumber As String, _
                                strFromDateTime As String, _
                                strToDateTime As String, _
                                strDateQueryType As String, _
                                strDateMacro As String, _
                                strCustomerJob As String, _
                                booCustomerWithChildren As Boolean, _
                                strAccount As String, _
                                booAccountWithChildren As Boolean, _
                                strFromRefNumberRange As String, _
                                strToRefNumberRange As String, _
                                strRefNumberPiece As String, _
                                strRefNumberCriteria As String, _
                                strPaidStatus As String)

  Dim strTimeIncluded As String
  Dim stemp As String
  Dim strItems As String
  Dim strReturnedRefNumber As String
  Dim i As Integer, j As Integer, k As Integer
  Dim sInv
  Dim cTotalBalance As Currency
  Dim sObjMsg As String
  Dim tempcount As String
  
  
On Error Resume Next
  
  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest
  
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", 2, 1)
  'Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  strSavedRequestCode = _
    "  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest" & vbCrLf & _
    "  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest(""US""," & Str(intMajorVersion) & ", " & Str(intMinorVersion) & ")" & vbCrLf & _
    "  objMsgSetRequest.Attributes.OnError = roeContinue" & vbCrLf & vbCrLf
  
  Dim objInvoiceQuery As QBFC3Lib.IInvoiceQuery
  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq
  
  strSavedRequestCode = strSavedRequestCode & _
    "  Dim objInvoiceQuery As QBFC3Lib.IInvoiceQuery" & vbCrLf & _
    "  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq" & vbCrLf
    
  objInvoiceQuery.IncludeLinkedTxns.SetValue True
  objInvoiceQuery.IncludeLineItems.SetValue True
    
  If strRefNumber <> "" Then
    objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add strRefNumber
    strSavedRequestCode = strSavedRequestCode & vbCrLf & _
      "  objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add " & strRefNumber
  Else
    'Get the invoice lines so we can put the line count in the invoice information
    strSavedRequestCode = strSavedRequestCode & vbCrLf & _
      "  objInvoiceQuery.IncludeLineItems.SetValue True" & vbCrLf & _
      "  objInvoiceQuery.IncludeLinkedTxns.SetValue True"
      
    'Use With statements to reduce the size of our lines
    With objInvoiceQuery.ORInvoiceQuery.InvoiceFilter
        
      'We're limiting ourselves to the first 30 invoices to avoid too much info
      .MaxReturned.SetValue 30
    strSavedRequestCode = strSavedRequestCode & vbCrLf & _
      "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.MaxReturned.SetValue 30" & vbCrLf
      
      If strFromDateTime <> "" Or strToDateTime <> "" Then
        'We'll either be using a Modified date or a Txn date for filtering
        If strDateQueryType = "ModifiedDateRangeFilter" Then
          With .ORDateRangeFilter.ModifiedDateRangeFilter
            If strFromDateTime <> "" Then
              If InStr(1, strFromDateTime, ":") Then
                strFromDateTime = Replace(strFromDateTime, "T", " ")
                .FromModifiedDate.SetValue CDate(strFromDateTime), True
                strTimeIncluded = "True"
              Else
                .FromModifiedDate.SetValue CDate(strFromDateTime), False
                strTimeIncluded = "False"
              End If
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate.SetValue CDate(""" & strFromDateTime & """), " & strTimeIncluded
            End If
      
            If strToDateTime <> "" Then
              If InStr(1, strToDateTime, ":") Then
                strToDateTime = Replace(strToDateTime, "T", " ")
                .ToModifiedDate.SetValue CDate(strToDateTime), True
                strTimeIncluded = "True"
              Else
                .ToModifiedDate.SetValue CDate(strToDateTime), False
                strTimeIncluded = "False"
              End If
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.ToModifiedDate.SetValue CDate(""" & strToDateTime & """), " & strTimeIncluded
            End If
          End With '.ORDateRangeFilter.ModifiedDateRangeFilter

        'Since the to or from date string isn't blank and the date
        'query type wasn't modified that mean's were using the Txn date filter
        Else 'strDateQueryType = "ModifiedDateRangeFilter"
          With .ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter
            If strFromDateTime <> "" Then
              .FromTxnDate.SetValue CDate(strFromDateTime)
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue CDate(""" & strFromDateTime & """)"
            End If
      
            If strToDateTime <> "" Then
              .ToTxnDate.SetValue CDate(strToDateTime)
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue CDate(""" & strToDateTime & """)"
            End If
          End With
        End If 'strDateQueryType = "ModifiedDate"
      End If 'strFromDateTime <> "" Or strToDateTime <> ""
    
      If strDateMacro <> "" Then
        With .ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter
          .DateMacro.SetAsString strDateMacro
        End With
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.DateMacro.SetAsString " & strDateMacro
      End If
    
      If strCustomerJob <> "" Then
        If booCustomerWithChildren Then
          .EntityFilter.OREntityFilter.FullNameWithChildren.SetValue strCustomerJob
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameWithChildren.SetValue " & strCustomerJob
        Else
          .EntityFilter.OREntityFilter.FullNameList.Add strCustomerJob
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameList.Add " & strCustomerJob
        End If
      End If
    
      If strAccount <> "" Then
        If booAccountWithChildren Then
          .AccountFilter.ORAccountFilter.FullNameWithChildren.SetValue strAccount
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.AccountFilter.ORAccountFilter.FullNameWithChildren.SetValue " & strAccount
        Else
          .AccountFilter.ORAccountFilter.FullNameList.Add strAccount
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.AccountFilter.ORAccountFilter.FullNameList.Add " & strAccount
        End If
      End If
    
      If strFromRefNumberRange <> "" Then
        .ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.SetValue strFromRefNumberRange
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.SetValue " & strFromRefNumberRange
      End If
  
      If strToRefNumberRange <> "" Then
        .ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.SetValue strToRefNumberRange
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.SetValue " & strToRefNumberRange
      End If
    
      If strRefNumberPiece <> "" Then
        .ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue strRefNumberPiece
        .ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetAsString strRefNumberCriteria
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue " & strRefNumberPiece
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetAsString " & strRefNumberCriteria
      End If
    
      If strPaidStatus <> "" Then
        .PaidStatus.SetAsString strPaidStatus
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.PaidStatus.SetAsString " & strPaidStatus
      End If
    End With 'objInvoiceQuery.ORInvoiceQuery.InvoiceFilter
  End If 'strRefNumber <> ""
  
  
  '''''''''''''''''''''''InvoiceRetList''''''''''''''''''''
  Dim objMsgSetResponse As QBFC3Lib.IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  'Form1.Text2.Text = objMsgSetResponse.ToXMLString
  'Form1.Show
  
  Dim objResponse As QBFC3Lib.IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.statusCode = 1 Then
    'lstInvoices.AddItem "No invoices match the query filter used"
    Exit Sub
  End If
  
  Dim objInvoiceRetList As QBFC3Lib.IInvoiceRetList
  Set objInvoiceRetList = objResponse.Detail
  Dim objInvoiceRet As QBFC3Lib.IInvoiceRet
  
          
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
    cTotalBalance = "0.00"
    lInvoiceNumber = 0
    lInvoiceNextPlace = 1
    
    Invoicedone = ""
    InvoiceCount = objInvoiceRetList.Count
    
    ReDim InvoicesInfo(InvoiceCount, 15)
    'objInvoiceRetList.Count is the number of invoice numbers that will be stored
    '30 is the amount of invoice items that created the invoice substance
    '5 is the details in each item in the invoice
    ReDim InvoicesDetails(InvoiceCount, 30, 7)
  
  
    For i = 0 To InvoiceCount - 1
    
        Set objInvoiceRet = objInvoiceRetList.GetAt(i)
        If objInvoiceRet.RefNumber Is Nothing Then
            strReturnedRefNumber = "Un-numbered "
        Else
            lInvoiceNumber = objInvoiceRet.RefNumber.GetValue
            strReturnedRefNumber = "Invoice " & lInvoiceNumber
        End If
                
        tempcount = objInvoiceRet.Subtotal.GetAsString
        tempcount = objInvoiceRet.AppliedAmount
        tempcount = objInvoiceRet.balanceRemaining.GetAsString
        If tempcount <> "0.00" Then
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''record invoice info for reporting and letters''''''''''''''''
                            
                InvoiceNum = i
                
                
                'items in the invoices: date, invoice #, terms, due date, rep, account, total, payments/credits, current balance, total balance
                InvoicesInfo(i, 0) = Trim(objInvoiceRet.TxnDate.GetValue)
                InvoicesInfo(i, 1) = lInvoiceNumber
                InvoicesInfo(i, 2) = Trim(objInvoiceRet.TermsRef.FullName.GetValue)
                InvoicesInfo(i, 3) = Trim(objInvoiceRet.DueDate.GetValue)
                InvoicesInfo(i, 4) = Trim(objInvoiceRet.SalesRepRef.FullName.GetValue)
                InvoicesInfo(i, 5) = Trim(objInvoiceRet.ARAccountRef.FullName.GetValue)
                InvoicesInfo(i, 6) = Trim(objInvoiceRet.AppliedAmount.GetAsString)
                InvoicesInfo(i, 7) = Trim(objInvoiceRet.balanceRemaining.GetAsString)
                InvoicesInfo(i, 8) = Trim(objInvoiceRet.CustomerMsgRef.FullName.GetValue)
                InvoicesInfo(i, 9) = Trim(objInvoiceRet.Subtotal.GetAsString)
                
                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            FG2.Rows = lInvoiceNextPlace + 1
            
            FG2.Row = lInvoiceNextPlace
            FG2.Col = 0
            FG2.Text = objInvoiceRet.TxnDate.GetValue
            FG2.Col = 1
            FG2.Text = lInvoiceNumber
            FG2.Col = 2
            FG2.Text = objInvoiceRet.Subtotal.GetAsString
                    
            lInvoiceNextPlace = lInvoiceNextPlace + 1
            bPayments = False
                            
                            
            displayLists (sCompany)
            
            
            FG2.Rows = lInvoiceNextPlace + 1
            FG2.Row = lInvoiceNextPlace
            FG2.Col = 0
            FG2.Text = "Balance"
            FG2.Col = 1
            FG2.Text = lInvoiceNumber
            FG2.Col = 2
            FG2.Text = objInvoiceRet.balanceRemaining.GetAsString
            lInvoiceNextPlace = lInvoiceNextPlace + 1
                
        End If
    
    Next
  
End Sub



Public Sub QBFC_GetInvoice(txnId As String)

  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  
  Dim objInvoiceQuery As QBFC3Lib.IInvoiceQuery
  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq
  
  objInvoiceQuery.ORInvoiceQuery.TxnIDList.Add txnId
  objInvoiceQuery.IncludeLineItems.SetValue True
  
  Set objInvoiceMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As QBFC3Lib.IResponse
  Set objResponse = objInvoiceMsgSetResponse.ResponseList.GetAt(0)
  
  Dim objInvoiceRetList As QBFC3Lib.IInvoiceRetList
  Set objInvoiceRetList = objResponse.Detail
  
  Set objSavedInvoiceRet = objInvoiceRetList.GetAt(0)
End Sub




Public Sub QBFC_LoadInvoiceLineArray(strLineArray() As String)

  Dim objInvoiceLineRetList As QBFC3Lib.IORInvoiceLineRetList
  Set objInvoiceLineRetList = objSavedInvoiceRet.ORInvoiceLineRetList
  
  
  Dim objLine As QBFC3Lib.IInvoiceLineRet
  Dim objGroupLine As QBFC3Lib.IInvoiceLineGroupRet
  Dim objGroupLines As QBFC3Lib.IInvoiceLineRetList
  
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  
  j = 0
  For i = 0 To objInvoiceLineRetList.Count - 1
  
    j = j + 1
    If objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet Is Nothing Then
      strLineArray(j) = ItemLineInfo(objInvoiceLineRetList.GetAt(i).InvoiceLineRet) & _
                        "Item<spliter>Original"
    Else
      strLineArray(j) = GroupLineInfo(objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet) & _
                        "Group<spliter>Original"
      
      Set objGroupLine = objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet
      Set objGroupLines = objGroupLine.InvoiceLineRetList
      If objGroupLines.Count > 0 Then
        For k = 0 To objGroupLines.Count - 1
          j = j + 1
          Set objLine = objGroupLines.GetAt(k)
          strLineArray(j) = ItemLineInfo(objLine) & "SubItem<spliter>Original"
        Next k
      End If 'objGroupLines.length > 0
    End If 'objNode.nodeName = "InvoiceLineRet"
  Next i
End Sub




Public Sub QBFC_GetItemInfo(strItemFullName As String, _
                            strDesc As String, _
                            strRate As String, _
                            strRateOrPercent As String, _
                            strSalesTaxCode As String)

  strDesc = ""
  strRate = ""
  strRateOrPercent = "Rate"
  strSalesTaxCode = ""
  
  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  Dim objItemQuery As QBFC3Lib.IItemQuery
  Set objItemQuery = objMsgSetRequest.AppendItemQueryRq
  objItemQuery.ORListQuery.FullNameList.Add strItemFullName
  
  Dim objMsgSetResponse As QBFC3Lib.IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As QBFC3Lib.IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  Dim objItemRetList As QBFC3Lib.IORItemRetList
  Set objItemRetList = objResponse.Detail
  
  Dim enItemType As QBFC3Lib.ENORItemRet
  enItemType = objItemRetList.GetAt(0).ortype
  
  Dim objItemRet
  Select Case enItemType
    Case Is = orirItemServiceRet
      Set objItemRet = objItemRetList.GetAt(0).ItemServiceRet
    Case Is = orirItemNonInventoryRet
      Set objItemRet = objItemRetList.GetAt(0).ItemNonInventoryRet
    Case Is = orirItemOtherChargeRet
      Set objItemRet = objItemRetList.GetAt(0).ItemOtherChargeRet
    Case Is = orirItemInventoryRet
      Set objItemRet = objItemRetList.GetAt(0).ItemInventoryRet
    Case Is = orirItemInventoryAssemblyRet
      Set objItemRet = objItemRetList.GetAt(0).ItemInventoryAssemblyRet
    Case Is = orirItemSubtotalRet
      Set objItemRet = objItemRetList.GetAt(0).ItemSubtotalRet
    Case Is = orirItemDiscountRet
      Set objItemRet = objItemRetList.GetAt(0).ItemDiscountRet
    Case Is = orirItemPaymentRet
      Set objItemRet = objItemRetList.GetAt(0).ItemPaymentRet
    Case Is = orirItemSalesTaxRet
      Set objItemRet = objItemRetList.GetAt(0).ItemSalesTaxRet
    Case Is = orirItemGroupRet
      Set objItemRet = objItemRetList.GetAt(0).ItemGroupRet
  End Select
  
  If enItemType = orirItemServiceRet Or enItemType = orirItemNonInventoryRet Or _
     enItemType = orirItemOtherChargeRet Then
    If Not objItemRet.ORSalesPurchase.SalesOrPurchase Is Nothing Then
      If Not objItemRet.ORSalesPurchase.SalesOrPurchase.Desc Is Nothing Then
        strDesc = objItemRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue
      End If
      
      If Not objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice Is Nothing Then
        If Not objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price Is Nothing Then
          strRate = objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetAsString
        Else
          strRate = objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.PricePercent.GetAsString
          strRateOrPercent = "RatePercent"
        End If
      End If
    Else ' Since it isn't SalesOrPurchase it must be SalesAndPurchase
      If Not objItemRet.ORSalesPurchase.SalesAndPurchase.SalesDesc Is Nothing Then
        strDesc = objItemRet.ORSalesPurchase.SalesAndPurchase.SalesDesc.GetValue
      End If
      
      If Not objItemRet.ORSalesPurchase.SalesAndPurchase.SalesPrice Is Nothing Then
        strRate = objItemRet.ORSalesPurchase.SalesAndPurchase.SalesPrice.GetAsString
      End If
    End If
  ElseIf enItemType = orirItemInventoryRet Or _
         enItemType = orirItemInventoryAssemblyRet Then
    If Not objItemRet.SalesDesc Is Nothing Then
      strDesc = objItemRet.SalesDesc.GetValue
    End If
  
    If Not objItemRet.SalesPrice Is Nothing Then
      strRate = objItemRet.SalesPrice.GetAsString
    End If
  Else
    If Not objItemRet.ItemDesc Is Nothing Then
      strDesc = objItemRet.ItemDesc.GetValue
    End If
  End If
  
  If Not (enItemType = orirItemSubtotalRet Or _
          enItemType = orirItemPaymentRet Or _
          enItemType = orirItemSalesTaxRet Or _
          enItemType = orirItemGroupRet) Then
    If Not objItemRet.SalesTaxCodeRef Is Nothing Then
      strSalesTaxCode = objItemRet.SalesTaxCodeRef.FullName.GetValue
    End If
  End If
End Sub


Public Sub QBFC_LoadRequest(strRequestText As String)
  strRequestText = strSavedRequestCode
End Sub


Private Function ItemLineInfo(objInvoiceLine As QBFC3Lib.IInvoiceLineRet) As String

  Dim strLineInfo As String
  Dim strRateOrPercent As String

  strLineInfo = objInvoiceLine.TxnLineID.GetValue & "<spliter>"
  
  If Not objInvoiceLine.Quantity Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceLine.Quantity.GetAsString
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.ItemRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceLine.ItemRef.FullName.GetValue
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.Desc Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceLine.Desc.GetValue
  End If
  strLineInfo = strLineInfo & "<spliter>"

  If Not objInvoiceLine.ORRate Is Nothing Then
    If Not objInvoiceLine.ORRate.Rate Is Nothing Then
      strLineInfo = strLineInfo & objInvoiceLine.ORRate.Rate.GetAsString
      strRateOrPercent = "Rate"
    Else
      strLineInfo = strLineInfo & objInvoiceLine.ORRate.RatePercent.GetAsString
      strRateOrPercent = "RatePercent"
    End If
  Else
    strRateOrPercent = "Neither"
  End If
  strLineInfo = strLineInfo & "<spliter>"

  If Not objInvoiceLine.amount Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceLine.amount.GetAsString
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.ClassRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceLine.ClassRef.FullName.GetValue
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.ServiceDate Is Nothing Then
    strLineInfo = strLineInfo & _
      Format(objInvoiceLine.ServiceDate.GetValue, "YYYY-MM-DD")
  End If
  strLineInfo = strLineInfo & "<spliter>"
    
  If Not objInvoiceLine.SalesTaxCodeRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceLine.SalesTaxCodeRef.FullName.GetValue
  End If
  strLineInfo = strLineInfo & "<spliter>" & strRateOrPercent & _
    "<spliter><spliter>"
  
  
  ItemLineInfo = strLineInfo
End Function


Private Function GroupLineInfo(objInvoiceGroupLine As QBFC3Lib.IInvoiceLineGroupRet) As String

  Dim strLineInfo As String
  Dim strRateOrPercent As String

  strLineInfo = objInvoiceGroupLine.TxnLineID.GetValue & "<spliter>"
  
  If Not objInvoiceGroupLine.Quantity Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceGroupLine.Quantity.GetAsString
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceGroupLine.ItemGroupRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceGroupLine.ItemGroupRef.FullName.GetValue
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceGroupLine.Desc Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceGroupLine.Desc.GetValue
  End If
  strLineInfo = strLineInfo & "<spliter><spliter><spliter><spliter>"

  If Not objInvoiceGroupLine.ServiceDate Is Nothing Then
    strLineInfo = strLineInfo & _
      Format(objInvoiceGroupLine.ServiceDate.GetValue, "YYYY-MM-DD")
  End If
  strLineInfo = strLineInfo & "<spliter><spliter><spliter><spliter>"
  
  GroupLineInfo = strLineInfo
End Function






Public Sub QBFC_FillReportList()

 ' Dim strTimeIncluded As String
 ' Dim stemp As String
 ' Dim strItems As String
  'Dim strReturnedRefNumber As String
  Dim i As Integer, j As Integer, k As Integer
  'Dim sInv
  'Dim cTotalBalance As Currency
  'Dim sObjMsg As String
  'Dim tempcount As String
  
  Dim booCustomerWithChildren As Boolean
  Dim strAccount As String
  Dim booAccountWithChildren As Boolean
  Dim strPaidStatus As String
  Dim strCustomerJob As String
  
On Error Resume Next

    'frmReports.List2.AddItem "->In Fill"
  
  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest
  
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", 2, 1)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  strSavedRequestCode = _
    "  Dim objMsgSetRequest As QBFC3Lib.IMsgSetRequest" & vbCrLf & _
    "  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest(""US""," & Str(intMajorVersion) & ", " & Str(intMinorVersion) & ")" & vbCrLf & _
    "  objMsgSetRequest.Attributes.OnError = roeContinue" & vbCrLf & vbCrLf
  
  Dim objInvoiceQuery As QBFC3Lib.IInvoiceQuery
  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq
  
  strSavedRequestCode = strSavedRequestCode & _
    "  Dim objInvoiceQuery As QBFC3Lib.IInvoiceQuery" & vbCrLf & _
    "  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq" & vbCrLf
    
    objInvoiceQuery.IncludeLinkedTxns.SetValue True
    objInvoiceQuery.IncludeLineItems.SetValue True
    'objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add Trim(frmReports.Text2.Text)
    
    'frmReports.List2.AddItem "->set init vars"
                    
    'strCustomerJob = Trim(Text2.Text)
    booCustomerWithChildren = True
    strAccount = "Accounts Receivable"
    booAccountWithChildren = False
    strPaidStatus = "NotPaidOnly"
    
    With objInvoiceQuery.ORInvoiceQuery.InvoiceFilter
    
        If strCustomerJob <> "" Then
            If booCustomerWithChildren Then
                .EntityFilter.OREntityFilter.FullNameWithChildren.SetValue strCustomerJob
               Else
                .EntityFilter.OREntityFilter.FullNameList.Add strCustomerJob
            End If
        End If
          
        If strAccount <> "" Then
            If booAccountWithChildren Then
                .AccountFilter.ORAccountFilter.FullNameWithChildren.SetValue strAccount
            Else
                .AccountFilter.ORAccountFilter.FullNameList.Add strAccount
            End If
        End If
        
        If strPaidStatus <> "" Then
            .PaidStatus.SetAsString strPaidStatus
        End If
    End With
    
    
    'frmReports.List2.AddItem "->set filter rules, doing request now..."
    
  
  '''''''''''''''''''''''InvoiceRetList''''''''''''''''''''
  Dim objMsgSetResponse As QBFC3Lib.IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
    
    'frmReports.List2.AddItem "->output string"
    'frmReports.Text1.Text = objMsgSetResponse.ToXMLString
    'frmReports.List2.AddItem "->string done"

  Dim objResponse As QBFC3Lib.IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.statusCode = 1 Then
    'lstInvoices.AddItem "No invoices match the query filter used"
    'frmReports.List2.AddItem "->no results!!!!!!!!!!!!!!!!!!!"
    Exit Sub
  End If
  
  Dim objInvoiceRetList As QBFC3Lib.IInvoiceRetList
  Set objInvoiceRetList = objResponse.Detail
  Dim objInvoiceRet As QBFC3Lib.IInvoiceRet
  
  
        
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    'frmReports.List2.AddItem "->results being processed"
    
    InvoiceCount = objInvoiceRetList.Count
    
    'frmReports.List1.Clear
    
    Dim temp As String
    Dim sumBalance As Currency
    Dim invdesc As String
    Dim invdate As String
    Dim invbal As Currency
    Dim invcname As String
    Dim litemName As String
    
    For i = 0 To InvoiceCount - 1
    
        Set objInvoiceRet = objInvoiceRetList.GetAt(i)
        
        temp = Trim(objInvoiceRet.balanceRemaining.GetAsString)
        
        If temp <> "0.00" And temp <> "" Then
        

            sumBalance = 0
            
            If objInvoiceRet.RefNumber Is Nothing Then
            Else
                lInvoiceNumber = objInvoiceRet.RefNumber.GetValue
            End If
                    
                                    
            
            'frmReports.List1.AddItem ""
            'frmReports.List1.AddItem ""
            'frmReports.List1.AddItem "======================================================="
            
            'items in the invoices: date, invoice #, terms, due date, rep, account, total, payments/credits, current balance, total balance
            'frmReports.List1.AddItem "Customer:" & Trim(objInvoiceRet.CustomerMsgRef.FullName.GetValue)
            'frmReports.List1.AddItem "Date:" & Trim(objInvoiceRet.TxnDate.GetValue)
            'frmReports.List1.AddItem lInvoiceNumber
            'frmReports.List1.AddItem Trim(objInvoiceRet.TermsRef.FullName.GetValue)
            
            'frmReports.List1.AddItem ""
            

            'frmReports.List1.AddItem " LineItemCount: " & objInvoiceRet.ORInvoiceLineRetList.Count
            For k = 0 To objInvoiceRet.ORInvoiceLineRetList.Count - 1
                invbal = 0
                litemName = Trim(objInvoiceRet.ORInvoiceLineRetList.GetAt(k).InvoiceLineRet.ItemRef.FullName.GetValue)
            
                If litemName <> "" Then
                    invdesc = Trim(objInvoiceRet.ORInvoiceLineRetList.GetAt(k).InvoiceLineRet.Desc.GetValue)
                    invdate = Trim(objInvoiceRet.ORInvoiceLineRetList.GetAt(k).InvoiceLineRet.ServiceDate.GetValue)
                    invbal = Trim(objInvoiceRet.ORInvoiceLineRetList.GetAt(k).InvoiceLineRet.amount.GetAsString)
                    
                    'frmReports.List1.AddItem k & " LineItem:" & invdesc & " / " & invdate & " / " & invbal
                    sumBalance = sumBalance + Trim(objInvoiceRet.ORInvoiceLineRetList.GetAt(k).InvoiceLineRet.amount.GetAsString)
                End If
                
            Next
            'frmReports.List1.AddItem "sum:" & sumBalance


    
            'frmReports.List1.AddItem ""
            
            'frmReports.List1.AddItem " PaymentCount: " & objInvoiceRet.LinkedTxnList.Count
            For j = 0 To objInvoiceRet.LinkedTxnList.Count - 1
                'frmReports.List1.AddItem j & " Payment: " & Trim(objInvoiceRet.LinkedTxnList.GetAt(j).TxnDate.GetValue) & " / " & Trim(objInvoiceRet.LinkedTxnList.GetAt(j).amount.GetAsString)
            Next
            
            'frmReports.List1.AddItem ""
            
            'frmReports.List1.AddItem " Subtotal: " & Trim(objInvoiceRet.Subtotal.GetAsString)
            'frmReports.List1.AddItem " Credit: " & Trim(objInvoiceRet.AppliedAmount.GetAsString)
            'frmReports.List1.AddItem " Balance: " & Trim(objInvoiceRet.balanceRemaining.GetAsString)
            
            
            
            'frmReports.List1.AddItem
            'frmReports.List1.AddItem Trim(objInvoiceRet.DueDate.GetValue)
            'frmReports.List1.AddItem Trim(objInvoiceRet.SalesRepRef.FullName.GetValue)
            'frmReports.List1.AddItem Trim(objInvoiceRet.ARAccountRef.FullName.GetValue)
        
        End If
    Next
  
End Sub




