Attribute VB_Name = "mod_sql"
Option Explicit


Private Type SQL_Con_Struct
    Con As New ADODB.Connection
    Cmd As New ADODB.Command
    Error As String
End Type

Private SQLcn As SQL_Con_Struct

Public Type Sql_Results_Struct
    Query As String
    Data As New ADODB.Recordset
End Type

Public Rs As Sql_Results_Struct



Public Function SQL_ReConnect_old(cn As ADODB.Connection) As Boolean
    SQL_ReConnect_old = False
    
On Error GoTo SQLError:

    SQL_ReConnect_old = SQL_Status_old(cn)
    If SQL_ReConnect_old = True Then Exit Function

    cn.Provider = "sqloledb"
    cn.CursorLocation = adUseClient
    cn.Properties("Data Source").Value = "192.168.168.51"
    cn.Properties("Initial Catalog").Value = "modcon"
    cn.Properties("User ID").Value = "sa"
    cn.Properties("Password").Value = "passwd"
    cn.Open
    
    If cn.State = 1 Then SQL_ReConnect_old = True
    Exit Function
    
SQLError:
    SQLErrorAction Err
End Function

Public Function SQL_Status_old(cn As ADODB.Connection) As Boolean
    SQL_Status_old = False
    
On Error GoTo SQLError:
    If cn.State = 1 Then SQL_Status_old = True
    Exit Function
    
SQLError:
    SQLErrorAction Err
End Function


Public Function SQL_Con() As Boolean
    SQL_Con = SQL_Connect
End Function
Public Function SQL_Connect() As Boolean
    SQL_Connect = False
    
On Error GoTo SQLError:

    SQL_Connect = SQL_Status
    If SQL_Connect = True Then Exit Function

    SQLcn.Con.Provider = "sqloledb"
    SQLcn.Con.CursorLocation = adUseClient
    SQLcn.Con.Properties("Data Source").Value = "192.168.168.51"
    SQLcn.Con.Properties("Initial Catalog").Value = "modcon"
    SQLcn.Con.Properties("User ID").Value = "sa"
    SQLcn.Con.Properties("Password").Value = "passwd"
    'Con.Properties("Data Source").Value = func2.funMCServer
    'Con.Properties("Initial Catalog").Value = func2.funMCDatabase
    'Con.Properties("User ID").Value = func2.funMCUsername
    'Con.Properties("Password").Value = func2.funMCPassword
    SQLcn.Con.Open
    
    If SQLcn.Con.State = 1 Then SQL_Connect = True
    Exit Function
    
SQLError:
    SQLErrorAction Err
End Function

Private Sub SQL_Init(sQuery As String)
    
On Error GoTo SQLError:
    Set SQLcn.Cmd.ActiveConnection = SQLcn.Con
    SQLcn.Cmd.CommandType = adCmdText
    SQLcn.Cmd.CommandText = sQuery
    Exit Sub
    
SQLError:
    SQLErrorAction Err
End Sub

Public Sub SQL_RS_Clear(rsData As ADODB.Recordset)
    
On Error GoTo SQLError:
    Set rsData = Nothing
    Exit Sub
    
SQLError:
    SQLErrorAction Err
End Sub
    
Public Sub SQL_Close_Clear(rsData As ADODB.Recordset)
    SQL_Close
    SQL_RS_Clear rsData
End Sub


Public Sub SQL_Query_auto(sQuery As String, rsData As ADODB.Recordset)
    Set rsData = Nothing
    If SQL_Connect = True Then SQL_Query sQuery, rsData
End Sub

Public Sub SQL_Query(sQuery As String, rsData As ADODB.Recordset)
    SQL_Init sQuery
    
On Error GoTo SQLError:
    Set rsData = SQLcn.Cmd.Execute
    Exit Sub
    
SQLError:
    SQLErrorAction Err
End Sub

Public Sub SQL_Update_auto(sQuery As String)
    If SQL_Connect = True Then SQL_Update sQuery
End Sub

Public Sub SQL_Update(sQuery As String)
    SQL_Init sQuery
    
On Error GoTo SQLError:
    SQLcn.Cmd.Execute
    Exit Sub
    
SQLError:
    SQLErrorAction Err
End Sub

Public Function SQL_Status() As Boolean
    SQL_Status = False
    
On Error GoTo SQLError:
    If SQLcn.Con.State = 1 Then SQL_Status = True
    Exit Function
    
SQLError:
    SQLErrorAction Err
End Function

Public Function SQL_Close() As Boolean
    SQL_Close = False
    
On Error GoTo SQLError: Err
    If SQLcn.Con.State = 1 Then SQLcn.Con.Close
    If SQLcn.Con.State = 0 Then SQL_Close = True
    Exit Function
    
SQLError:
    SQLErrorAction Err
End Function


Private Sub SQLErrorAction(eErrs As ErrObject)
    Dim errLoop As ErrObject, Response
    SQLcn.Error = "Errors: " & vbNewLine

    For Each errLoop In eErrs
        SQLcn.Error = SQLcn.Error & "Info: " & Err.source & " (" & Err.Number & ") - " & vbNewLine
        SQLcn.Error = SQLcn.Error & Err.Description & vbNewLine
        If errLoop.HelpFile = "" Then
            SQLcn.Error = SQLcn.Error & "No Helpfile available" & vbNewLine
        Else
            SQLcn.Error = SQLcn.Error & "Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext & vbNewLine
        End If
    Next
        
    Response = MsgBox("Would you like to Continue?, " & SQLcn.Error & vbNewLine & "Continue?", vbCritical + vbYesNo, "SQL Error")
    If Response = vbYes Then Resume Else Exit Sub
    
End Sub



