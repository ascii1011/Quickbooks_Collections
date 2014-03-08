Attribute VB_Name = "modGeneric"
Option Explicit

Dim Implementation As String
Dim strInvoiceLineInfo As String

Dim strqbXMLLevel As String
Dim booSupportsModify As Boolean
Dim booSupportsDateTime As Boolean

Public Function QuickBooksVersionOK() As Boolean
  'The OpenConnectionBeginSession routines will exit this sample program
  'if installation or access problems are found
  
    
  
  'Now check to see what version of QuickBooks is running and what version
  'of qbXML it supports
  strqbXMLLevel = GetMaxVersionSupported
  
  'Check for version 2.1 or 3.X
  'If Not (InStr(1, strqbXMLLevel, "2.1") > 0) And Not (InStr(1, strqbXMLLevel, "3.") > 0) Then
    'MsgBox "The configuration you're running against " & _
    '  "does not support Invoice Modify." & vbCrLf & vbCrLf & _
    '  "You will be able to query " & _
    '  "for invoices, but you will not be able to modify them."
    booSupportsModify = False
  'Else
    'booSupportsModify = True
  'End If
  
  If InStr(1, strqbXMLLevel, "2") Or InStr(1, strqbXMLLevel, "3") Then
    booSupportsDateTime = True
  Else
    booSupportsDateTime = False
  End If
  
  QuickBooksVersionOK = True
End Function




Function GetMaxVersionSupported() As String
  If Implementation = "QBXMLRP" Then
    'GetMaxVersionSupported = QBXMLRP_MaxVersionSupported
  ElseIf Implementation = "QBFC" Then
    GetMaxVersionSupported = QBFC_MaxVersionSupported
'  Else
'    GetMaxVersionSupported = QBFCCA_MaxVersionSupported
  End If
End Function
    
    
    










