Attribute VB_Name = "modQBXMLRP"

Option Explicit

Dim objRequestProcessor As QBXMLRPLib.RequestProcessor

Dim strTicket As String
Dim booSessionBegun As Boolean

Dim strqbXMLVersionLine As String

Dim objSavedDOMDocument As New DOMDocument40
Dim objSavedInvoiceRet As IXMLDOMNode

Dim strSavedRequest As String



Public Sub QBXMLRP_OpenConnectionBeginSession()
  
  On Error GoTo Errs
  
  Set objRequestProcessor = New QBXMLRPLib.RequestProcessor
  
  objRequestProcessor.OpenConnection cAppID, cAppName
  strTicket = objRequestProcessor.BeginSession(companyFile, qbFileOpenDoNotCare)
  booSessionBegun = True
  Exit Sub
  
Errs:
  If Err.Number = &H80040416 Then
    MsgBox "You must have QuickBooks running with a company" & vbCrLf & _
           "file open to use this program."
    objRequestProcessor.CloseConnection
    End
  ElseIf Err.Number = &H80040422 Then
    MsgBox "This QuickBooks company file is open in single user mode and" & vbCrLf & _
           "another application is already accessing it.  Please exit the" & vbCrLf & _
           "other application and run this application again."
    objRequestProcessor.CloseConnection
    End
  ElseIf Err.Number = &H1AD Then
    MsgBox _
      "It appears that the qbXML Request Processor has not" & vbCrLf & _
      "been installed, indicating QuickBooks 2002 or later" & vbCrLf & _
      "may not have been installed.  Please run this sample" & vbCrLf & _
      "after installing QuickBooks 2003 and running the Upgrade."
    End
  ElseIf Err.Number = &H1AE Then
    MsgBox _
      "It appears that you have QuickBooks 2002 R1P installed." & vbCrLf & _
      "This program requires QuickBooks 2003 to work."
    End
  Else
    MsgBox Err.Number & vbCrLf & Hex(Err.Number) & vbCrLf & _
           Err.Description
    End
  End If
End Sub


Public Sub QBXMLRP_EndSessionCloseConnection()
  If booSessionBegun Then
    objRequestProcessor.EndSession strTicket
    objRequestProcessor.CloseConnection
  End If
End Sub



Public Sub QBXMLRP_FillInvoiceList(FG2 As MSFlexGrid, _
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

  Dim strXMLRequest As String
  
  If strRefNumber <> "" Then
    'We only need to query for the ref number, don't bother building
    'the XML with MSXML
    strXMLRequest = _
      "<?xml version=""1.0""?>" & strqbXMLVersionLine & _
      "<QBXML><QBXMLMsgsRq onError=""stopOnError""><InvoiceQueryRq>" & _
      "<RefNumber>" & strRefNumber & "</RefNumber>" & _
      "<IncludeLineItems>true</IncludeLineItems></InvoiceQueryRq>" & _
      "</QBXMLMsgsRq></QBXML>"
  Else 'strRefNumber <> ""
    Dim objDOMDocument As New DOMDocument40
    
    Dim objRootNode As IXMLDOMNode
    Dim objRequestNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim objAttribute As IXMLDOMAttribute
    
    CreateStandardRequestNode _
      False, "continueOnError", objDOMDocument, objRootNode, objRequestNode, objAttribute
    
    Dim objInvoiceQueryNode As IXMLDOMNode
    AddMSXMLNode "InvoiceQueryRq", objDOMDocument, objRequestNode, objInvoiceQueryNode
    
    'Limit our response to 30 invoices
    AddMSXMLElement "MaxReturned", "30", objDOMDocument, objInvoiceQueryNode, objElement
    
    If strFromDateTime <> Empty Or strToDateTime <> Empty Then
      Dim objDateTimeFilter As IXMLDOMNode
      AddMSXMLNode strDateQueryType, objDOMDocument, objInvoiceQueryNode, objDateTimeFilter
      If strDateQueryType = "ModifiedDateRangeFilter" Then
        If strFromDateTime <> Empty Then
          AddMSXMLElement _
            "FromModifiedDate", strFromDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
        If strToDateTime <> Empty Then
          AddMSXMLElement _
            "ToModifiedDate", strToDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
      Else 'strDateQueryType = "ModifiedDateRangeFilter"
        If strFromDateTime <> Empty Then
          AddMSXMLElement _
            "FromTxnDate", strFromDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
        If strToDateTime <> Empty Then
          AddMSXMLElement _
            "ToTxnDate", strToDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
      End If 'strDateQueryType = "ModifiedDateRangeFilter"
    End If 'strFromDateTime <> Empty Or strToDateTime <> Empty
    
    If strDateMacro <> Empty Then
      Dim objDateFilter As IXMLDOMNode
      AddMSXMLNode "TxnDateRangeFilter", objDOMDocument, objInvoiceQueryNode, objDateFilter
      AddMSXMLElement "DateMacro", strDateMacro, objDOMDocument, objDateFilter, objElement
    End If
    
    If strCustomerJob <> Empty Then
      Dim objEntityFilter As IXMLDOMNode
      AddMSXMLNode _
        "EntityFilter", objDOMDocument, objInvoiceQueryNode, objEntityFilter
      
      If booCustomerWithChildren Then
        AddMSXMLElement _
          "FullNameWithChildren", strCustomerJob, objDOMDocument, objEntityFilter, objElement
      Else
        AddMSXMLElement _
          "FullName", strCustomerJob, objDOMDocument, objEntityFilter, objElement
      End If
    End If
    
    If strAccount <> Empty Then
      Dim objAccountFilter As IXMLDOMNode
      AddMSXMLNode _
        "AccountFilter", objDOMDocument, objInvoiceQueryNode, objAccountFilter
      
      If booAccountWithChildren Then
        AddMSXMLElement _
          "FullNameWithChildren", strAccount, objDOMDocument, objAccountFilter, objElement
      Else
        AddMSXMLElement _
          "FullName", strAccount, objDOMDocument, objAccountFilter, objElement
      End If
    End If
    
    If strFromRefNumberRange <> Empty Or strToRefNumberRange <> Empty Then
      Dim objRefNumberRangeFilter As IXMLDOMNode
      AddMSXMLNode "RefNumberRangeFilter", objDOMDocument, objInvoiceQueryNode, objRefNumberRangeFilter
      If strFromRefNumberRange <> Empty Then
        AddMSXMLElement _
          "FromRefNumber", strFromRefNumberRange, objDOMDocument, objRefNumberRangeFilter, objElement
      End If
      If strToRefNumberRange <> Empty Then
        AddMSXMLElement _
        "ToRefNumber", strToRefNumberRange, objDOMDocument, objRefNumberRangeFilter, objElement
      End If
    End If 'strFromRefNumberRange <> Empty Or strToRefNumberRange <> Empty

    If strRefNumberPiece <> Empty Then
      Dim objRefNumberFilter As IXMLDOMNode
      AddMSXMLNode "RefNumberFilter", objDOMDocument, objInvoiceQueryNode, objRefNumberFilter
      AddMSXMLElement "MatchCriterion", strRefNumberCriteria, objDOMDocument, objRefNumberFilter, objElement
      AddMSXMLElement "RefNumber", strRefNumberPiece, objDOMDocument, objRefNumberFilter, objElement
    End If

    If strPaidStatus <> Empty Then
      AddMSXMLElement "PaidStatus", strPaidStatus, objDOMDocument, objInvoiceQueryNode, objElement
    End If
  
    AddMSXMLElement "IncludeLineItems", "true", objDOMDocument, objInvoiceQueryNode, objElement
    
    strXMLRequest = "<?xml version=""1.0""?>" & strqbXMLVersionLine & objRootNode.xml
  End If 'strRefNumber <> ""
  PrettyPrintXMLToFile strXMLRequest, "C:\debugrq.xml"
  strSavedRequest = PrettyPrintXMLToString(strXMLRequest)
  
  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  PrettyPrintXMLToFile strXMLResponse, "C:\debugrs.xml"
  
  objDOMDocument.async = False
  objDOMDocument.loadXML (strXMLResponse)
  
  Dim objInvoiceQueryNodeList As IXMLDOMNodeList
  Set objInvoiceQueryNodeList = objDOMDocument.getElementsByTagName("InvoiceQueryRs")
  
  Set objInvoiceQueryNode = objInvoiceQueryNodeList.Item(0)
  
  Dim objInvoiceQueryAttributes As IXMLDOMNamedNodeMap
  Set objInvoiceQueryAttributes = objInvoiceQueryNode.Attributes
  
  If objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error getting Invoices" & vbCrLf & vbCrLf & _
      "Error = " & _
      objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objInvoiceQueryAttributes.getNamedItem("statusMessage").nodeValue
      
    '''''''''''''lstInvoices.AddItem "No invoices match the query filter used"
    Exit Sub
  End If
  
  Dim objInvoiceRetNodeList As IXMLDOMNodeList
  Set objInvoiceRetNodeList = objDOMDocument.getElementsByTagName("InvoiceRet")
  
  Dim objInvoiceRet As IXMLDOMNode
  Dim objNodeList As IXMLDOMNodeList
  Dim intItems As Integer
  Dim strReturnedRefNumber As String
  Dim strItems As String
  
  Dim i As Integer
  
  FG2.Clear
  FG2.Rows = objInvoiceRetNodeList.length
  
  For i = 0 To objInvoiceRetNodeList.length - 1
    Set objInvoiceRet = objInvoiceRetNodeList.Item(i)
    
    If Not objInvoiceRet.selectSingleNode("RefNumber") Is Nothing Then
      strReturnedRefNumber = _
        "Invoice " & objInvoiceRet.selectSingleNode("RefNumber").Text
    Else
      strReturnedRefNumber = "Un-numbered "
    End If
    
    Set objNodeList = objInvoiceRet.selectNodes("InvoiceLineRet")
    If objNodeList Is Nothing Then
      intItems = 0
    Else
      intItems = objNodeList.length
    End If
    Set objNodeList = objInvoiceRet.selectNodes("InvoiceLineGroupRet")
    If Not objNodeList Is Nothing Then
      intItems = intItems + objNodeList.length
    End If
    
    strItems = Str(intItems)
    If Len(strItems) = 1 Then strItems = "  " & strItems
    If Len(strItems) = 2 Then strItems = " " & strItems
    
    FG2.Row = i + 2
    FG2.Col = 0
    FG2.Text = strReturnedRefNumber
    FG2.Col = 1
    FG2.Text = objInvoiceRet.selectSingleNode("TxnDate").Text
    FG2.Col = 2
    FG2.Text = objInvoiceRet.selectSingleNode("BalanceRemaining").Text
    
    'lstInvoices.AddItem _
    '  strReturnedRefNumber & "     " & _
    '  objInvoiceRet.selectSingleNode("TxnDate").Text & _
    '  "     " & strItems & " items     " & _
    '  objInvoiceRet.selectSingleNode("CustomerRef").selectSingleNode("FullName").Text & _
    '  "     Balance " & objInvoiceRet.selectSingleNode("BalanceRemaining").Text & _
    '  "     " & objInvoiceRet.selectSingleNode("TxnID").Text
      
  Next i
End Sub





