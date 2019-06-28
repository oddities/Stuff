Option Explicit

Public Conn As Object
Public USERFULLNAME As String
Public USERLANID As String
Public ibOutcome() As String, obOutcome() As String



'minimize maximize
Public Declare Function FindWindowA& Lib "User32" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function GetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&)
Public Declare Function SetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)
 
' Déclaration des constantes
Public Const GWL_STYLE As Long = -16
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_FULLSIZING = &H70000
'minimize maximize

Sub connectToDB(pConn As Object, pDBPath As String, pUserId As String, pPassword As String, pDBType As String)
    On Error GoTo ErrorBlock
    Dim sconnect As String
    
    If Not pConn Is Nothing Then
        If pConn.State = 1 Then
            pConn.Close
        End If
    End If
    Set pConn = Nothing

    Set pConn = CreateObject("ADODB.Connection")
    
    If pDBType = "Text" Then
        'connect string for database as .csv file (text file)
        sconnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & pDBPath & ";" & _
                       "Extended Properties=""text;HDR=Yes"""
    End If
    
    If pDBType = "Excel" Then
        'connect string for database as excel file
        sconnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pDBPath & _
                       ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"
    End If

    If pDBType = "MSAccess" Then
        'connect string for database as ms access
        sconnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pDBPath & ";"
    End If
    
    If pDBType = "Oracle" Then
        sconnect = "Provider=MSDAORA;Data Source=FTCST01P_SSL;User ID=" & pUserId & ";Password=" & pPassword
    End If
    pConn.Open sconnect
Done:
    Exit Sub
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.connectToDB", Erl
    Resume Done
End Sub

Sub disconnectFromDB(pConn As Object)
    On Error GoTo ErrorBlock
    If Not pConn Is Nothing Then
        If pConn.State = 1 Then
            pConn.Close
        End If
    End If
    
    Set pConn = Nothing

Done:
    Exit Sub
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.disconnectFromDB", Erl
    Resume Done
End Sub

Sub closeRecordSet(pRs As Object)
    On Error GoTo ErrorBlock
    If Not pRs Is Nothing Then
        If pRs.State = 1 Then
            pRs.Close
        End If
    End If
    
    Set pRs = Nothing

Done:
    Exit Sub
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.closeRecordSet", Erl
    Resume Done
End Sub


Sub fetchInputFileData(pPathname As String)
    On Error GoTo ErrorBlock
    
    Dim fullpath As String
    
    fullpath = Application.GetOpenFilename _
    (Title:="Please choose a file", _
    Filefilter:="Access DB Files,*.accdb")
    
    If fullpath = "False" Then
        MsgBox "No File Specified.", vbExclamation, "ERROR"
        GoTo Done
    End If
    pPathname = fullpath
Done:
    Exit Sub
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.fetchInputFileData", Erl
    Resume Done
End Sub

Function fetchRowNo(pParaName As String) As Integer
    On Error GoTo ErrorBlock
    Dim noOfRows As Long, row As Long

    ConfigSheet.Activate
    noOfRows = Cells(ActiveSheet.Rows.count, 2).End(xlUp).row
    
    For row = 1 To noOfRows
        If Trim(Cells(row, 2)) = pParaName Then
            Exit For
        End If
    Next row
    fetchRowNo = row
Done:
    Exit Function
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.fetchRowNo", Erl
    Resume Done
End Function

Function FileExists(FilePath As String) As Boolean
    On Error GoTo ErrorBlock
    Dim TestStr As String
    
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
Done:
    Exit Function
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.FileExists", Erl
    Resume Done
End Function


Function AuthenticateUsers() As Boolean
    On Error GoTo ErrorBlock
    Dim sSqlQry As String, counter As Integer, found As Boolean

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    USERLANID = Environ$("Username")
    
    connectToDB Conn, ThisWorkbook.DBpath, "", "", "MSAccess"
    If Conn.State = 0 Then
        MsgBox "Not connected to DB!!!"
        GoTo Done
    End If
    sSqlQry = "Select fullname " & _
              "from AuthenticatedUsers " & _
              "where lanid = '" & USERLANID & "' " & _
              "and AccessLevel in ('Manager','Agent')"
    rs.Open sSqlQry, Conn
    found = False
    If rs.EOF = False Then
        USERFULLNAME = rs(0)
        found = True
        
    End If
    
    closeRecordSet rs
    
    AuthenticateUsers = found
Done:
    closeRecordSet rs
    disconnectFromDB Conn
    Exit Function
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.AuthenticateUsers", Erl
    Resume Done
End Function

'minimize maximize
'Attention, envoyer après changement du caption de l'UF
Public Sub InitMaxMin(mCaption As String, Optional Max As Boolean = True, Optional Min As Boolean = True _
        , Optional Sizing As Boolean = True)
    On Error GoTo ErrorBlock
    Dim hWnd As Long
    hWnd = FindWindowA(vbNullString, mCaption)
    If Min Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MINIMIZEBOX
    If Max Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
    If Sizing Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_FULLSIZING
Done:
    Exit Sub
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.InitMaxMin", Erl
    Resume Done
End Sub
'minimize maximize

Sub readconfigfile()
    On Error GoTo ErrorBlock
    Dim noOfRows As Long, row As Long
    Dim ErrorFlag As Boolean, textLine As String, pos As Integer
    Dim sConfigName As String
    ErrorFlag = False
    ConfigSheet.Activate
    
    sConfigName = Application.ActiveWorkbook.Path & "\" & "config.prop"
    
    If FileExists(sConfigName) = True Then
        ThisWorkbook.DBpath = ""
        ThisWorkbook.CJDBpath = ""
        Open sConfigName For Input As #1
        Do Until EOF(1)
            Line Input #1, textLine
            'MsgBox "textLine = " & textLine
            pos = InStr(textLine, "=")
            If pos > 0 Then
                If Left(textLine, (pos - 1)) = "ACCESSDB" Then
                    ThisWorkbook.DBpath = Mid(textLine, pos + 1)
                End If
            End If
            If pos > 0 Then
                If Left(textLine, (pos - 1)) = "CJACCESSDB" Then
                    ThisWorkbook.CJDBpath = Mid(textLine, pos + 1)
                End If
            End If
        Loop
        Close #1
        If Len(ThisWorkbook.DBpath) = 0 Then
            MsgBox "FATAL ERROR: ACCESSDB parameter not found. Program exiting."
            ErrorFlag = True
            ActiveWorkbook.Close savechanges:=False
        Else
            If FileExists(ThisWorkbook.DBpath) = False Then
                MsgBox "FATAL ERROR: Access DB file " & ThisWorkbook.DBpath & " not found. Program exiting."
                ErrorFlag = True
                ActiveWorkbook.Close savechanges:=False
            End If
        End If
        If Len(ThisWorkbook.CJDBpath) = 0 Then
            MsgBox "FATAL ERROR: CJACCESSDB parameter not found. Program exiting."
            ErrorFlag = True
            ActiveWorkbook.Close savechanges:=False
        Else
            If FileExists(ThisWorkbook.CJDBpath) = False Then
                MsgBox "FATAL ERROR: Access DB file " & ThisWorkbook.CJDBpath & " not found. Program exiting."
                ErrorFlag = True
                ActiveWorkbook.Close savechanges:=False
            End If
        End If
    Else
        MsgBox "FATAL ERROR: File: " & sConfigName & " not found. Program exiting."
        ErrorFlag = True
        ActiveWorkbook.Close savechanges:=False
    End If

    If ErrorFlag = False Then
        If AuthenticateUsers() = False Then
            MsgBox "ERROR: User not allowed to use this file."
            'ErrorFlag = True
            ActiveWorkbook.Close savechanges:=False
'        Else
'            CustomerDataForm.Show
        End If
    End If
Done:
    Exit Sub
ErrorBlock:
    DisplayError Err.Source, Err.Description, "Module1.readconfigfile", Erl
    Resume Done
End Sub
