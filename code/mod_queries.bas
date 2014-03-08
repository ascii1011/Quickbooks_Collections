Attribute VB_Name = "mod_queries"
Option Explicit


Private Type Query_Security_Struct
    bView_Cust_Media As Boolean
    bView_Cust_Open As Boolean
    bView_Cust_Closed As Boolean
    bView_Cust_Credits As Boolean
    bView_Cust_Zero As Boolean
    bView_Inv_Zero As Boolean
    bView_Inv_Payments As Boolean
End Type

Public qSec As Query_Security_Struct

Private Type Query_Cust_List_Struct
    Account As String
    Balance As String
    Customer As String
    Importance As String
    Rep As String
    Status As String
    UpFront As String
    Region As String
    OrderName As String
    OrderBy As String
    How As String
    Value As String
    Query As String
    Where As String
End Type

Public CustList As Query_Cust_List_Struct

'redefine qSec structure as no access
Public Sub reDim_qSec()
    qSec.bView_Cust_Media = False
    qSec.bView_Cust_Open = False
    qSec.bView_Cust_Closed = False
    qSec.bView_Cust_Credits = False
    qSec.bView_Cust_Zero = False
    qSec.bView_Inv_Zero = False
    qSec.bView_Inv_Payments = False
End Sub

'redefine CustList structure as no search parameters
Public Sub reDim_CustList()
    CustList.Account = ""
    CustList.Balance = ""
    CustList.Customer = ""
    CustList.Importance = ""
    CustList.Rep = ""
    CustList.Status = ""
    CustList.UpFront = ""
    CustList.Region = ""
    CustList.OrderName = ""
    CustList.OrderBy = ""
    CustList.How = ""
    CustList.Value = ""
    CustList.Query = " select " & _
                    " cust_listid, cust_accountnumber, cust_fullname, " & _
                    " cust_totalbalance_money, cust_jobstatus, cust_salesrepref_fullname, " & _
                    " importance_name, importance_upfront, cust_accountnumber_numeric " & _
                    " from qbx_cust "
    CustList.Where = " where cust_fullname <> 'Customer' "
End Sub



Public Function QryBuildCustList() As String
    
    CustList.Query = CustList.Query & _
                    QrySec(CustList.Where, False) & _
                    QryCSearch(CustList.Where)
                    
    QryBuildCustList = CustList.Query
End Function

Public Function QryCSearch(Where As String) As String
    QryCSearch = Where
    reDim_CustList
    
    QryCSearch = ifEmpty(QryCSearch, " where importance_upfront = '" & CustList.UpFront & "' ")
    
    QryCSearch = ifEmpty(QryCSearch, " importance_name = '" & CustList.Importance & "' ")
    
    If CustList.Region <> "" Then QryCSearch = ifEmpty(QryCSearch, frmInvoiceQry.funStates(CustList.Region))
    
    If frmInvoiceQry.sOrderBy = "" Then
        CustList.OrderName = " cust_fullname "
        CustList.OrderBy = " asc "
        frmInvoiceQry.sOrderBy = " cust_fullname asc "
        QryCSearch = QryCSearch & " order by " & CustList.OrderName & " " & CustList.OrderBy
    End If
    
End Function

Public Function QrySec(Where As String, bJoined As Boolean) As String
    Dim jm_a As String, jm_b As String
    
    QrySec = Where
    reDim_qSec
    
    jm_a = ""
    jm_b = ""
    If bJoined = True Then
        jm_a = "a."
        jm_b = "b."
    End If
    
    If qSec.bView_Cust_Media = False Then
        'not able to view media, so ('awarded' = media customers)
        QrySec = ifEmpty(QrySec, jm_b & "cust_jobstatus <> 'awarded' and " & jm_b & " cust_accountnumber <> '' ")
    End If
            
    If qSec.bView_Cust_Open = True And qSec.bView_Cust_Closed = False Then
        'if only allowed to view open customer records
        QrySec = ifEmpty(QrySec, " CONVERT(int, " & jm_b & "cust_totalbalance_money) <> '0' ")
    ElseIf qSec.bView_Cust_Open = False And qSec.bView_Cust_Closed = True Then
        'if only allowed to view closed customer records
        QrySec = ifEmpty(QrySec, " CONVERT(int, " & jm_b & "cust_totalbalance_money) = '0' ")
    ElseIf qSec.bView_Cust_Open = False And qSec.bView_Cust_Closed = False Then
        'if not allowed, search for nothing of importance.
        QrySec = " where " & jm_b & "cust_jobstatus = 'Security Risk' "
        Exit Function
    End If
    
    If qSec.bView_Cust_Credits = False Then
        'if not able to view credits
        QrySec = ifEmpty(QrySec, " sign(" & jm_b & "cust_totalbalance_money) != CONVERT(money, '-1') ")
    End If
    
    'If jm_a <> "" Then QrySec = ifEmpty(QrySec, " " & jm_a & "alert_id = " & jm_b & "cust_listid ")
    
End Function

Private Function ifEmpty(old_qry As String, cur_qry As String) As String
    If old_qry = "" Then
        old_qry = " where " & cur_qry
    Else
        old_qry = old_qry & " and " & cur_qry
    End If
End Function




'Function funReturnPriviledgeRestrictions_old(ExtraWhere As String) As String
'    funReturnPriviledgeRestrictions = QrySec(ExtraWhere)
'End Function
Function funReturnPriviledgeRestrictions(ExtraWhere As String) As String
    Dim strXtraSQL As String
            
        'show media means that cust_jobstatus = 'awarded'
        If sProfileAttrDtlsAry(1, 6) = 0 Then
            strXtraSQL = " where b.cust_jobstatus <> 'awarded' "
        End If
                
        'Able to view "Open Invoices"
        If sProfileAttrDtlsAry(1, 10) = 0 Then
            'Able to view "Closed Invoices"
            If sProfileAttrDtlsAry(1, 11) = 0 Then
                If strXtraSQL = "" Then
                    strXtraSQL = " where CONVERT(int, b.cust_totalbalance_money) = '0' "
                Else
                    strXtraSQL = strXtraSQL & " and CONVERT(int, b.cust_totalbalance_money) = '0' "
                End If
            Else
                If strXtraSQL = "" Then
                    strXtraSQL = " where CONVERT(int, b.cust_totalbalance_money) = '0' "
                Else
                    strXtraSQL = strXtraSQL & " and CONVERT(int, b.cust_totalbalance_money) = '0' "
                End If
            End If
        Else
            If sProfileAttrDtlsAry(1, 11) = 0 Then
                If strXtraSQL = "" Then
                    strXtraSQL = " where CONVERT(int, b.cust_totalbalance_money) <> '0' "
                Else
                    strXtraSQL = strXtraSQL & " and CONVERT(int, b.cust_totalbalance_money) <> '0' "
                End If
            Else
                If strXtraSQL = "" Then
                    strXtraSQL = " where CONVERT(int, b.cust_totalbalance_money) <> '0' "
                Else
                    strXtraSQL = strXtraSQL & " and CONVERT(int, b.cust_totalbalance_money) <> '0' "
                End If
            End If
        End If
        
        'view credits
        If sProfileAttrDtlsAry(1, 44) = 0 Then
            If strXtraSQL = "" Then
                strXtraSQL = " where sign(b.cust_totalbalance_money) != CONVERT(money, '-1') "
            Else
                strXtraSQL = strXtraSQL & " and sign(b.cust_totalbalance_money) != CONVERT(money, '-1') "
            End If
        End If
        
        If strXtraSQL = "" Then
            If ExtraWhere <> "" Then strXtraSQL = " where " & ExtraWhere
        Else
            If ExtraWhere <> "" Then strXtraSQL = strXtraSQL & " and " & ExtraWhere
        End If
        
        strXtraSQL = strXtraSQL & " and a.alert_id = b.cust_listid "
    
    funReturnPriviledgeRestrictions = strXtraSQL
    
End Function

