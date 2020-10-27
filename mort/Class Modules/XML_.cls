VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XML_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name            : XML_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                                                 with XML
' ==========================================================================

Public Function array_to_xml(recordset As Variant, headers As Variant, Optional Remove_Amp_Spaces_Replace_w_underscores As Boolean = False) As String
    Dim sb As StringBuilder: Set sb = New StringBuilder
    Dim strTableElementName$, strRowElementName$
    Dim varColumnHeaders As Variant
    Dim varTable As Variant
    Dim intRow&, intCol&
    Dim currData$
     
    'Set custom names
    strTableElementName = "Table"
    strRowElementName = "Row"
    
    'Get Recordset data
    varTable = recordset
    varColumnHeaders = clean_headers(headers)

    ' // This may needs to be changed to accomdiate for
    ' // possible zero-index arrays
    sb.Append "<" & strTableElementName & ">"
        For intRow = 1 To UBound(varTable, 1)
            sb.Append "<" & strRowElementName & ">"
            For intCol = LBound(varTable, 2) To UBound(varTable, 2)
                currData = SQL_ReplaceRestricted(varTable(intRow, intCol))
                If Remove_Amp_Spaces_Replace_w_underscores = False Then
                    sb.Append "<" & varColumnHeaders(intCol - 1) & ">" & currData & "</" & varColumnHeaders(intCol - 1) & ">"
                Else
                    sb.Append "<" & SQL_RemoveAmp_and_Space_replace_Underscore(varColumnHeaders(intCol - 1)) & ">" & currData & "</" & SQL_RemoveAmp_and_Space_replace_Underscore(varColumnHeaders(intCol - 1)) & ">"
                End If
            Next
           sb.Append "</" & strRowElementName & ">"
        Next
    sb.Append "</" & strTableElementName & ">"
 
    array_to_xml = sb.ToString
End Function

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' // Support Functions
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Private Function clean_headers(ByVal headers As Variant) As Variant
    Dim results() As Variant
    Dim i&
    
    ReDim results(LBound(headers) To UBound(headers))
    For i = LBound(headers) To UBound(headers)
        results(i) = STRING_.replace(headers(i), " ", "_")
    Next i
    clean_headers = results
End Function

Private Function SQL_RemoveAmp_and_Space_replace_Underscore(ByVal record As Variant) As Variant
    If Not InStr(record, "&") = 0 Then
            record = replace(record, "&", "")
    End If
    If Not InStr(record, " ") = 0 Then
            record = replace(record, " ", "_")
    End If
    SQL_RemoveAmp_and_Space_replace_Underscore = record
End Function
Private Function SQL_ReplaceRestricted(ByVal record As String) As String
    If Not InStr(record, "'") = 0 Then
        record = replace(record, "'", "&apos;")
    End If
    
    If Not InStr(record, "&") = 0 Then
        record = replace(record, "&", "&amp;")
    End If
    
    If Not InStr(record, """") = 0 Then
        record = replace(record, """", "&quot;")
    End If
    SQL_ReplaceRestricted = record
End Function

