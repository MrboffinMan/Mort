VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QUERY_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Module      : QUERY_
' Type        : Class
' Description : Used to query SQL Server in Async fashion
'             : I really.... really... need to clean this ish up!
' --------------------------------------------------------------------------
' Procedures  : Example #1                              Variant
' --------------------------------------------------------------------------
' References  : N/A
' ==========================================================================

'Custom Events
Public Event QueryFinished(returnedRecordset As Variant, vHeaders As Variant)
Public Event NoRecordsReturned(sMessage As String)
Public Event ConnectionClosed()
Public Event QueryError(ByVal currErr As Object)
' Adds ADODB.Connection Events
Private WithEvents AsyncConnection As ADODB.Connection
Attribute AsyncConnection.VB_VarHelpID = -1

' Constants
Private Const iQUERY_TIMEOUT = 300

' Class Variables
Private sServerName As String
Private sQueryString As String
Private sStoredProcedure As String
Private vStoredProcedureArgs As Variant
Private bReturnEvent As Boolean
Private QryDest As QueryDest
Private UsingAsyncQuery As Boolean


Public Enum QueryDest
    destSQLServer = 0
    destExcel = 1
End Enum


Private Sub Class_Initialize()
    bReturnEvent = True
    UsingAsyncQuery = True
    QryDest = destSQLServer
End Sub

' Class Properties
Property Let QueryDestination(dest As QueryDest)
    QryDest = dest
End Property

Property Get GetQueryDestination() As String
    If QryDest = 0 Then
        GetQueryDestination = "destSQLServer"
    ElseIf QryDest = 1 Then
        GetQueryDestination = "destExcel"
    End If
End Property

Property Let ReturnEvents(bDecision As Boolean)
    bReturnEvent = bDecision
End Property
Property Let AddServer(sServer As String)
    sServerName = sServer
End Property

Property Let RawQueryString(sNewQueryString As String)
    sStoredProcedure = ""
    vStoredProcedureArgs = ""
    sQueryString = sNewQueryString
End Property

Property Get GetCurrentUser()
    GetCurrentUser = GetUserName()
End Property

Private Function GetUserName() As String
    On Error GoTo ErrHandler:
    GetUserName = LCase(Environ("username"))
    
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("GetUserName", Err, True)
End Function

Public Sub AddStoredProcedure(sProcName As String, Optional vArguments As Variant)
    ' Only takes 1D arrays
    sQueryString = ""
    
    sStoredProcedure = sProcName
    vStoredProcedureArgs = vArguments
End Sub


Property Let GenerateQueryString(sNewQueryString As String, vArguments As Variant)
    If Len(sQueryString) > 0 Then
        'Reset Query String to Nothing
        sQueryString = ""
        
        sQueryString = CreateQuery(sNewQueryString, vArguments)
    End If
End Property

Property Get Reset()
    sServerName = ""
    sQueryString = ""
End Property

Property Let UseAsync(use As Boolean)
    UsingAsyncQuery = use
End Property

' Sub Functions
Public Sub RunProc()
    On Error GoTo ErrHandler:
    If sStoredProcedure <> "" Then
        If UsingAsyncQuery Then
            Call AsyncDatabaseQuery("Stored Procedure", sServerName)
        Else
            Call NonAsyncDatabaseQuery("Stored Procedure", sServerName)
        End If
   Else
        Debug.Print ("Something went wrong")
        Debug.Print ("Please check that you've added both a Stored Procedure & Server Name")
        Err.Raise vbObjectError + 380, "RunProc", "I'm having Param Problems"
    End If
Done:
    Exit Sub
ErrHandler:
    Call LOGGER_.print_error("RunProc", Err, True)
    'RaiseEvent QueryError(Error)
End Sub

Public Sub RunQuery()
    On Error GoTo ErrHandler:
    If sQueryString <> "" And sServerName <> "" And sStoredProcedure = "" And vStoredProcedureArgs = "" Then
        If UsingAsyncQuery Then
            Call AsyncDatabaseQuery("SQL Query", sServerName)
        Else
            Call NonAsyncDatabaseQuery("SQL Query", sServerName)
        End If
    Else
        Debug.Print ("Something went wrong")
        Debug.Print ("Please check that you've added both a Query String & Server Name")
        Err.Raise vbObjectError + 380, "RunQuery", "I'm having Param Problems"
    End If
Done:
    Exit Sub
ErrHandler:
    Call LOGGER_.print_error("RunQuery", Err, True)
   ' RaiseEvent QueryError(pError)

End Sub

Private Sub NonAsyncDatabaseQuery(sQueryType As String, ByVal sServerName As String)
    On Error GoTo ErrHandler:
    
    Dim conn As Object
    Dim rs As Object
    
    Dim sConnectionString As String
    Dim sQuery As String
    
    Set conn = New ADODB.Connection
    Set conn = GetNewConnection(sServerName)
    Set rs = New ADODB.recordset
    
    conn.CommandTimeout = iQUERY_TIMEOUT
    
    If sQueryType = "Stored Procedure" Then
        'Dim oCommand As Object
        If IsMissing(vStoredProcedureArgs) Then
            sQuery = sStoredProcedure
        Else
            sQuery = BuildStoredProcedureString(sStoredProcedure, vStoredProcedureArgs)
        End If
        Debug.Print ("Query Started")
        'Call AsyncConnection.Open(sQuery, CommandTypeEnum.adCmdText) ', ExecuteOptionEnum.adAsyncExecute)
        Set rs = conn.Execute(sQuery, CommandTypeEnum.adCmdText)
    Else
        Debug.Print ("Query Started")
        'Call AsyncConnection.Open(sQueryString, CommandTypeEnum.adCmdText) ', ExecuteOptionEnum.adAsyncExecute)
        Set rs = conn.Execute(sQueryString, CommandTypeEnum.adCmdText)
    End If
    
    Set conn = Nothing

    Debug.Print ("Query Finished")
    If rs.EOF = True Then
        If bReturnEvent = True Then
            RaiseEvent NoRecordsReturned("No records were returned with this query")
        End If
    Else
        If bReturnEvent = True Then
            RaiseEvent QueryFinished(ArrayFromRecordset(rs), GetQueryResultHeaders(rs))
        End If
    End If

    
    rs.Close
    Set rs = Nothing
Done:
    Exit Sub
ErrHandler:
    Call LOGGER_.print_error("NonAsyncDatabaseQuery", Err, True)
End Sub


Private Sub AsyncDatabaseQuery(sQueryType As String, ByVal sServerName As String)
    On Error GoTo ErrHandler:
    
    Dim sConnectionString As String
    Dim sQuery As String
    
    Set AsyncConnection = New ADODB.Connection
    Set AsyncConnection = GetNewConnection(sServerName)
    AsyncConnection.CommandTimeout = iQUERY_TIMEOUT
    
    If sQueryType = "Stored Procedure" Then
        'Dim oCommand As Object
        If IsMissing(vStoredProcedureArgs) Then
            sQuery = sStoredProcedure
        Else
            sQuery = BuildStoredProcedureString(sStoredProcedure, vStoredProcedureArgs)
        End If
        Call AsyncConnection.Execute(sQuery, CommandTypeEnum.adCmdText, ExecuteOptionEnum.adAsyncExecute)
    Else
        Call AsyncConnection.Execute(sQueryString, CommandTypeEnum.adCmdText, ExecuteOptionEnum.adAsyncExecute)
    End If
    Debug.Print ("Query Started")
Done:
    Exit Sub
ErrHandler:
    Call LOGGER_.print_error("AsyncDatabaseQuery", Err, True)
    RaiseEvent QueryError(Error)
End Sub

Private Function CreateQuery(ByVal inputQueryString As String, Inputs As Variant) As String
    Dim i As Long
    Dim j As Long
    Dim tmpStr As String
    Dim SearchValue As String
    Dim replaceValue As String
    Dim newString As String
    Dim aInputs() As Variant
    
    Debug.Print (TypeName(Inputs))
    On Error GoTo ErrHandler:
    If TypeName(Inputs) = "String" Then
        ReDim aInputs(1)
        aInputs(1) = Inputs
    Else
        aInputs = Inputs
        If InStr(inputQueryString, "{*}") <> 0 Then
            tmpStr = ""
            For j = LBound(aInputs) To UBound(aInputs)
                If j = LBound(aInputs) Then
                    tmpStr = tmpStr & "{" & j & "}"
                Else
                    tmpStr = tmpStr & ",{" & j & "}"
                End If
            Next j
            inputQueryString = replace(inputQueryString, "{*}", tmpStr)
        End If
    End If
    
    Debug.Print (inputQueryString)
    newString = inputQueryString
    For i = LBound(aInputs) To UBound(aInputs)
        SearchValue = "{" & i & "}"
        If UBound(aInputs, 2) > 0 Then
            replaceValue = aInputs(i, 1)
        Else
            replaceValue = aInputs(i)
        End If
        newString = replace(newString, SearchValue, replaceValue)
    Next i
    CreateQuery = newString
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("CreateQuery", Err, True)
End Function

Private Function BuildStoredProcedureString(sProcName As String, Optional vArguments As Variant) As String
    Dim i As Long
    Dim sArguments As String
    
    On Error GoTo ErrHandler:
    If IsMissing(vArguments) Then
        BuildStoredProcedureString = sProcName
    ElseIf (LBound(vArguments) - UBound(vArguments)) = 0 And Left(vArguments(LBound(vArguments)), 1) <> "'" And Left(vArguments(UBound(vArguments)), 1) <> "'" Then
        BuildStoredProcedureString = sProcName & " '" & CStr(vArguments(1)) & "'"
    ElseIf (LBound(vArguments) - UBound(vArguments)) = 0 And Left(vArguments(LBound(vArguments)), 1) = "'" And Right(vArguments(LBound(vArguments)), 1) = "'" Then
        BuildStoredProcedureString = sProcName & " " & CStr(vArguments(1))
    Else
        sArguments = ""
        For i = LBound(vArguments) To UBound(vArguments)
            If i = LBound(vArguments) Then
                sArguments = "'" & replace(vArguments(i), "'", "") & "'"
            Else
                sArguments = sArguments & ", '" & replace(vArguments(i), "'", "") & "'"
            End If
        Next i
        sArguments = Trim(sArguments)
        Debug.Print (sProcName & " " & sArguments)
        BuildStoredProcedureString = sProcName & " " & sArguments
    End If
Done:
    
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("BuildStoredProcedureString", Err, True)
End Function

Private Function GetNewConnection(ByVal serverName As String) As Object
    Dim objConn As Object
    Dim sConnectionStr As String
    Dim sServerName As String
    Const iAdStateOpenID As Integer = 1
    
    On Error GoTo ErrHandler:
    Set objConn = CreateObject("ADODB.Connection")
  
    sConnectionStr = GenerateConnectionString(QryDest, serverName)
    objConn.Open sConnectionStr
    
    If objConn.State = iAdStateOpenID Then
        Set GetNewConnection = objConn
    End If
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("GetNewConnection", Err, True)
End Function

Private Sub AsyncConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.recordset, ByVal pConnection As ADODB.Connection)
    On Error GoTo ErrHandler:
    If (adStatus = EventStatusEnum.adStatusOk) Then
        Debug.Print ("Query Finished")
        If pRecordset.EOF = True Then
            If bReturnEvent = True Then
                RaiseEvent NoRecordsReturned("No records were returned with this query")
            End If
        Else
            If bReturnEvent = True Then
                RaiseEvent QueryFinished(ArrayFromRecordset(pRecordset), GetQueryResultHeaders(pRecordset))
            End If
        End If
    End If
        
    If (pConnection.State = ObjectStateEnum.adStateOpen) Then
        pConnection.Close
        RaiseEvent ConnectionClosed
    End If
Done:
    Exit Sub
ErrHandler:
    Dim currError As Object
    Set currError = LOGGER_.clone_object(Err)
    Call LOGGER_.print_error("AsyncConnection_ExecuteComplete", currError, True)
    If Not pError Is Nothing Then
        RaiseEvent QueryError(pError)
    Else
        RaiseEvent QueryError(currError)
    End If
End Sub


Private Function GetQueryResultHeaders(recordset As Object) As Variant
    Dim resultHeaders As Variant
    Dim Field As Variant
    Dim i As Integer
    
    i = 1
    
    On Error GoTo ErrHandler:
    ReDim resultHeaders(1 To 1, 1 To recordset.Fields.count)
    
    For Each Field In recordset.Fields
        resultHeaders(1, i) = Field.Name
        i = i + 1
    Next Field
    GetQueryResultHeaders = resultHeaders
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("GetQueryResultHeaders", Err, True)
End Function


Public Function ArrayFromRecordset(oRecordset As Object) As Variant
    On Error GoTo ErrHandler:
    
    Dim vArray() As Variant
    Dim returnArr As Variant
    Dim i As Long
    Dim iNumFields As Long
    Dim Field As Variant
    
    Dim lCounter As Long
    Const lChunkSize As Long = 100000
    
    iNumFields = oRecordset.Fields.count
    ReDim vArray(1 To lChunkSize, 1 To iNumFields)
    
    Do While oRecordset.EOF = False
        lCounter = lCounter + 1
        If lCounter > UBound(vArray) Then
            vArray = ExpandArray(vArray, 1, lChunkSize, Empty)
            
            'ReDim Preserve vArray(1 To UBound(vArray) + lChunkSize, 1 To iNumFields)
        End If
        
        i = 1
        For Each Field In oRecordset.Fields
            vArray(lCounter, i) = Field.Value
            i = i + 1
        Next Field
        oRecordset.MoveNext
    Loop

    returnArr = ARRAY_.new_smaller_array(vArray, lCounter, iNumFields)
    ArrayFromRecordset = returnArr
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("ArrayFromRecordset", Err, True)
    RaiseEvent QueryError(Error)

End Function


Private Function GenerateConnectionString(dest As QueryDest, server As Variant) As Variant
    If dest = 0 Then
        GenerateConnectionString = "Provider=SQLOLEDB.1;Trusted_connection=yes;Data Source=" & server
    ElseIf dest = 1 Then
        GenerateConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & server & _
              ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
    End If
End Function



