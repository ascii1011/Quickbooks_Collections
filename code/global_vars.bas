Attribute VB_Name = "global_vars"
Option Explicit

''''''''''''''''''''''''''rules''''''''''''''''''''''''''''
'0 = false
'1 = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Global Const sMCServer = "modcon0009"
Global Const sMCDatabase = "modcon"
Global Const sMCUsername = "sa"
Global Const sMCPassword = "passwd"

' These constants hold the application name and ID plus the
' major and minor versions for QBFC 2.0.
Global Const cAppID = "8673"
Global Const cAppName = "MC Collection"
'Global Const companyFile = "C:\Program Files\Intuit\QuickBooks Premier - Professional Services Edition\sample_consulting business.qbw"
'Global Const companyFile = "c:\temp\Modern Consumer.QBW"
'Global Const companyFile = "z:\steve\qb\Modern Consumer.QBW"
'Global Const companyFile = "z:\Modern Consumer.QBW"
'Global Const companyFile = "Z:\QB\collections\charty\Modern Consumer.QBW"
Public companyFile As String

Public bFocus As String

'''''''''''''''''''''''''sql property values
Public sGHtml_printing As String
Public sGImage_Modcon_C As String
Public sGHtml_Reporting As String
Public sGHtml_Dealer_Status As String
Public sGLink_Cmc As String
'''''''''''''''''''''''''''''''''''''''''''''

Global sCompany As String
Public iCalendarRequest As Integer
Public ListID   As String
Global bPayments As Boolean
Global iUpdateTries As Integer

Global sInvoices() As String
Global sReceivedPayments() As String
Global sTotalAmountPending As String
Global sSecLvl As String
Global sSecEnabled As String
Global sFullName As String
Global sAccountNumber As String
Global sProcessingText As String

Global aryGImportLvl() As String
Global aryGRegions() As String


Dim Implementation As String
Dim strInvoiceLineInfo As String

Dim strqbXMLLevel As String
Dim booSupportsModify As Boolean
Dim booSupportsDateTime As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Reporting structure
Public Type Reporting_Structure
    Tpl_Type As String   'Template to use
    Tpl_Html As String   'Template Html Header/css
    Tpl_Title As String   'Template Title
    Tpl_Body As String   'Template Body
    Cust As String  'Customer info
    Inv As String   'Invoice info
    Imp As String   'importance info
    iNoteCount As String    'amount to display
    Note As String  'note info
    t_numItems As String
    t_ifound As String
    t_sTBal As String
    t_dTxnDate4 As String
    t_dTxnDate3 As String
    c_txnId As String
    c_Rep As String
    c_Name As String
    c_Bal As String
    c_Contact As String
    c_Phone As String
    c_Status As String
    c_WebStatus As String
    c_Importancelvl As String
    c_Upfront As String
End Type

Public Report As Reporting_Structure

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This information is pulled as needed from the QB reps profile
Public aryRep_Profile(8) As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'for preview and printing of collection and invoice letters,
'use for only one customer's information

'customer information billing/shipping info
Public Billing(8) As String
Public Shipping(8) As String
Public sBilling As String
Public sShipping As String
Public CustomerMainInfo(11) As String
Public sEmailAttachmentMessage As String

'structure replacing above array CustomerMainInfo
Public Type CustomerInfo
    CompanyName As String
    CompanyFullName As String
    FirstName As String
    LastName As String
    BillAddress1 As String
    BillAddress2 As String
    BillAddress3 As String
    BillAddress4 As String
    ShipAddress1 As String
    ShipAddress2 As String
    ShipAddress3 As String
    ShipAddress4 As String
    Balance As Currency
    TotalBalance As Currency
End Type

Public CurrCust As CustomerInfo

'this is a way to track the payments that are referenced by the invoice number
'a local function will track and display the payments that are not associated with any invoice number and will then display them after everything else.
Public Type Payments
    inv_txnid_link As String
    inv_txndate As String
    inv_pay_amount As String
    inv_pay_refnumber As String
    display_this As Boolean
End Type

Public PaymentsByInvoiceOnly() As Payments

' 1 # of items that are in this invoice
' 2 items in the invoices: date, invoice #, terms, due date, rep, account, total, payments/credits, current balance, total balance
Public InvoicesInfo() As String

' 1 grab all the items in the invoice: item, quantity, description, rate, amount
Public InvoicesDetails() As String
Public InvoiceCount As Integer
Public InvoiceNum As String
Public Invoicedone As String
Public InvoiceNumTemp As String

Public lInvoiceNumber As Long
Public iInvoiceQryLostFocus As Integer
Public iMainLostFocus As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''tells whether a form is open already or not
'''''''''''' 0 = not open '''''' 1 = open ''''''''''''''

Public sFrmCalendar As Integer
Public sFrmCallbacks As Integer
Public sFrmEmailCustomer As Integer
Public sFrmInvoiceQry As Integer
Public sFrmPrintPage As Integer
Public sFrmProcessing As Integer
Public sFrmProfiles As Integer
Public sFrmReporting As Integer
Public sFrmLogin As Integer
Public sFrmSecurityLock As Integer
Public sFrmQueryFilter As Integer
Public sFrmRemarks As Integer
Public sFrmQBInStats As Integer
Public sFrmEmailImport As Integer
Public sFrmPriorityAlerts As Integer
Public sfrmSearchfor As Integer
Public sfrmImportanceSettings As Integer
Public sfrmQuickBooksFaxes As Integer


Public sKillPriorityAlert As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'options
'0 = no
'1 = yes

Global Const iProfileCount = "46"
Public sProfileAttrNamesAry(iProfileCount) As String
Public sProfileAttrDtlsAry(8, iProfileCount) As String

''''''''''''''''''Startup Options
'enable the invoiceqry form to startup right away when program is started
'sProfileAttrNamesAry( = "Enable 'Collections' Window"
Public iEnableInvoiceQryToStartUp As Integer
'enable the callbacks form to startup right away when program is started
'sProfileAttrNamesAry(1) = "Enable 'Callbacks' Window"
Public iEnableCallbacksToStartUp As Integer
'enable the PrintPage form to startup right away when program is started
'sProfileAttrNamesAry(2) = "Enable 'Print Page' Window"
Public iEnablePrintPageToStartUp As Integer
'enable the PrintReports form to startup right away when program is started
'sProfileAttrNamesAry(3) = "Enable 'Print Reports' Window"
Public iEnablePrintReportsToStartUp As Integer
'enable the Manage users form to startup right away when program is started
'sProfileAttrNamesAry(4) = "Enable 'Profile Managers' Window"
Public iEnableManageUsersToStartUp As Integer

''''''''''''''''''viewing attributes
'show media companies
'sProfileAttrNamesAry(5) = "Show 'Media' companies"
Public iShowMedia As Integer
'show companies that have a zero balance
'sProfileAttrNamesAry(6) = "Show companies with a zero balance"
Public iShowZeroBalance As Integer
'be able to view customer messages added by this application (in sql)
'sProfileAttrNamesAry(7) = "Able to view 'Customer Messages'"
Public iEnableViewingOfMsgs As Integer
'be able to save a customer message
'sProfileAttrNamesAry(8) = "Able to Save 'Customer Messages'"
Public iEnableSaveMsgs As Integer

'be able to view invoices
    'be able to see Open Invoices
    'sProfileAttrNamesAry(9) = "Able to view 'Open Invoices'"
    Public iEnableViewInvoicesOpen As Integer
    'be able to see Closed Invoices
    'sProfileAttrNamesAry(10) = "Able to view 'Closed Invoices'"
    Public iEnableViewInvoicesClosed As Integer
    'be able to view Invoices within the last ... days
    'this will be the positive number only
    'sProfileAttrNamesAry(11) = "Number of Days, Months, or Years"
    Public iViewInvoicesNumber As Double
    'this will tell whether it is the day, month, or year
    'm = month, d = day, y = year
    'sProfileAttrNamesAry(12) = "Type: Days, Months, or Years"
    Public iViewInvoicesDateType As String
'be able to view payments
'sProfileAttrNamesAry(13) = "Be able to view Payments"
Public iEnableViewPayements As Integer
    'be able to view payments within the last ... days
    'this will be the positive number only
    'sProfileAttrNamesAry(14) = "Number of Days, Months, or Years"
    Public iViewPaymentsNumber As Double
    'this will tell whether it is the day, month, or year
    'm = month, d = day, y = year
    'sProfileAttrNamesAry(15) = "Type: Days, Months, or Years"
    Public iViewPaymentsDateType As String


''''''''''''''''''viewing window buttons
'be able to email a customer
'sProfileAttrNamesAry(16) = "Be able to see the 'Email a Customer' button"
Public iEnableEmailCustomer As Integer
    'be able to change email information
    'sProfileAttrNamesAry(17) = "Be able to change 'Email a Customer' info"
    Public iEnableEditEmailCustomer As Integer
'be able to send an alert about a customer
'sProfileAttrNamesAry(18) = "Be able to see the 'Alert' button"
Public iEnableAlertButton As Integer
'be able to print....
'sProfileAttrNamesAry(19) = "Be able to see the 'Print' button"
Public iEnablePrintButton As Integer
    'be able to print invoices
    'sProfileAttrNamesAry(20) = "Be able to print an 'Invoice' letter"
    Public iEnablePrintInvoicesletter As Integer
    'be able to print collection letters
    'sProfileAttrNamesAry(21) = "Be able to print a 'Collection' letter"
    Public iEnablePrintCollectionLetter As Integer
'be able to print a report
'sProfileAttrNamesAry(22) = "Be able to see the 'Print Report' button"
Public iEnablePrintReportbutton As Integer


''''''''''''''''''View Toolbar buttons
'be able to view collections button
'sProfileAttrNamesAry(23) = "Be able to see the Toolbar 'Collections' button"
Public iEnableToolbarCollections As Integer
'be able to view refresh button
'sProfileAttrNamesAry(24) = "Be able to see the Toolbar 'Refresh' button"
Public iEnableToolbarRefresh As Integer
'be able to view Profile Manager button
'sProfileAttrNamesAry(25) = "Be able to see the Toolbar 'Profile Manager' button"
Public iEnableToolbarManageUsers As Integer
'be able to view search button
'sProfileAttrNamesAry(26) = "Be able to see the Toolbar 'Search' button"
Public iEnableToolbarSearch As Integer
'be able to view help button
'sProfileAttrNamesAry(27) = "Be able to see the Toolbar 'Help' button"
Public iEnableToolbarHelp As Integer
'be able to view Print Report button
'sProfileAttrNamesAry(28) = "Be able to see the Toolbar 'Print Report' button"
Public iEnableToolbarPrintReport As Integer


''''''''''''''''''View top dropdown menu items
'be able to view collections
'sProfileAttrNamesAry(29) = "Be able to see the Menu 'Collections' button"
Public iEnableDropCollections As Integer
'be able to view calendar
'sProfileAttrNamesAry(30) = "Be able to see the Menu 'Calendar' button"
Public iEnableDropCalendar As Integer
'be able to view callbacks
'sProfileAttrNamesAry(31) = "Be able to see the Menu 'Callbacks' button"
Public iEnableDropCallbacks As Integer
'be able to view refresh
'sProfileAttrNamesAry(32) = "Be able to see the Menu 'Refresh' button"
Public iEnableDropRefresh As Integer
'be able to view Profile Manager
'sProfileAttrNamesAry(33) = "Be able to see the Menu 'Profile Manager' button"
Public iEnableDropManageUsers As Integer
'be able to view clear windows
'sProfileAttrNamesAry(34) = "Be able to see the Menu 'Clear Windows' button"
Public iEnableDropClearWindows As Integer
'be able to view print report
'sProfileAttrNamesAry(35) = "Be able to see the Menu 'Print Report' button"
Public iEnableDropPrintReport As Integer


''''''''''''''''''be able to access forms
'be able to access calendar
'sProfileAttrNamesAry(36) = "Be able to access 'Calendar' windows"
Public iEnableAccessCalendar As Integer
'be able to access callbacks
'sProfileAttrNamesAry(37) = "Be able to access 'Callbacks' windows"
Public iEnableAccessCallbacks As Integer
'be able to access emailcustomers
'sProfileAttrNamesAry(38) = "Be able to access 'Email Customers' windows"
Public iEnableAccessEmailCustomer As Integer
'be able to access invoiceqry
'sProfileAttrNamesAry(39) = "Be able to access 'Collections' windows"
Public iEnableAccessInvoiceqry As Integer
'be able to access print page
'sProfileAttrNamesAry(40) = "Be able to access 'Print Page' windows"
Public iEnableAccessPrintPage As Integer
'be able to access print report
'sProfileAttrNamesAry(41) = "Be able to access 'Print Report' windows"
Public iEnableAccessPrintReport As Integer

'be able to access filter settings
'sProfileAttrNamesAry(42) = "Be able to access 'Query Filters' windows"
Public iEnableAccessQueryFilter As Integer

'''''''added 01/03/05
'be able to view credited accounts.  accounts that have an excess of money
'sProfileAttrNamesAry(43) = "Be able to view 'Credited Accounts' windows"
Public iEnableViewCreditedCustomer As Integer
'be able to view credited invoices.  accounts that have an excess of money
'sProfileAttrNamesAry(44) = "Be able to view "Credited Invoices" windows"
Public iEnableViewCreditedInvoices As Integer
'show invoices that have a zero balance (similar to #6)
'sProfileAttrNamesAry(45) = "Show Invoices with a zero balance"
Public iShowInvoiceZeroBalance As Integer
'''''''

    
