Attribute VB_Name = "modGenApps"
Option Explicit




Function funFormatDecimal(sNumber As String) As String
    Dim sVar As String
    Dim sRemake As String
    Dim i As Integer
    Dim ilength As Integer
    Dim iendlength As Integer
    Dim bDecimalFlag As Boolean
                
    bDecimalFlag = False
    ilength = Len(sNumber)
    iendlength = 0
                
    For i = 0 To ilength - 1
        sVar = Left(sNumber, 1)
        If bDecimalFlag = False Then
            If IsNumeric(sVar) Then
                sRemake = sRemake & sVar
            ElseIf sVar = "." Then
                bDecimalFlag = True
                sRemake = sRemake & sVar
            ElseIf sVar = "-" Then
                sRemake = sRemake & sVar
            End If
        Else
            If IsNumeric(sVar) Then
                sRemake = sRemake & sVar
                iendlength = iendlength + 1
            End If
        End If
        sNumber = Right(sNumber, Len(sNumber) - 1)
    Next i
    
    If bDecimalFlag = False Then
        sRemake = sRemake & ".00"
    Else
        If iendlength = 1 Then
            sRemake = sRemake & "0"
        End If
    End If
    
    funFormatDecimal = sRemake
    
End Function

Function funFormatCurr2String(sNumber As Currency) As String
    Dim sVar As String
    Dim sRemake As String
    Dim i As Integer
    Dim ilength As Integer
    Dim iendlength As Integer
    Dim bDecimalFlag As Boolean
    Dim sTmpNumber As String
    
    sTmpNumber = sNumber
    sRemake = ""
    bDecimalFlag = False
    ilength = Len(sTmpNumber)
    iendlength = 0
    
    For i = 0 To ilength - 1
        sVar = Left(sTmpNumber, 1)
        If bDecimalFlag = False Then
            If IsNumeric(sVar) Then
                sRemake = sRemake & sVar
            ElseIf sVar = "." Then
                bDecimalFlag = True
                sRemake = sRemake & sVar
            End If
        Else
            If IsNumeric(sVar) Then
                sRemake = sRemake & sVar
                iendlength = iendlength + 1
            End If
        End If
        sTmpNumber = Right(sTmpNumber, Len(sTmpNumber) - 1)
    Next i
    
    
    
    
    If bDecimalFlag = False Then
        sRemake = sRemake & ".00"
    Else
        If iendlength = 1 Then
            sRemake = sRemake & "0"
        End If
    End If
    
    Dim sTempEnd As String
    Dim sTempComma2 As String
    Dim sTempMoney As String
    Dim iCommaCount As Integer
    sTempMoney = ""
    
    'strip end off
    sTempEnd = Right(sRemake, 3)
    sTempComma2 = Left(sRemake, Len(sRemake) - 3)
    sRemake = sTempComma2
    
    If Len(sRemake) >= 4 Then
        iCommaCount = 0
        For i = 0 To Len(sRemake) - 1
            sTempMoney = Right(sTempComma2, 1) & sTempMoney
            sTempComma2 = Left(sTempComma2, Len(sTempComma2) - 1)
            iCommaCount = iCommaCount + 1
            If iCommaCount = 3 And i < Len(sRemake) - 1 Then
                sTempMoney = "," & sTempMoney
                iCommaCount = 0
            End If
        Next i
        sTempMoney = sTempMoney & sTempEnd
    Else
        sTempMoney = sRemake & sTempEnd
    End If
    
    
    funFormatCurr2String = sTempMoney
    
End Function

