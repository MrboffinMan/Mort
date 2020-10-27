VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARRAY_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name            : ARRAY_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                                                 with Arrays
' ==========================================================================
Public Function single_dim_shrink_array(ByVal input_arr As Variant, new_total_rows As Long) As Variant
    Dim result() As Variant
    Dim i&
    
    ReDim result(LBound(input_arr) To new_total_rows)
    For i = LBound(input_arr) To new_total_rows
        result(i) = input_arr(i)
    Next i
    single_dim_shrink_array = result
End Function

Public Function remove_between(Data As Variant, sStart As String, sEnd As String) As Variant ', Optional DeleteRowsAfter As Long = 0) As Variant
    Dim results() As Variant: ReDim results(LBound(Data) To UBound(Data))
    Dim BAD As Boolean: BAD = False
    Dim idxCheck&
    Dim i&
    Dim j&
    
    j = 0
    For i = LBound(Data) To UBound(Data)
        If Not BAD Then
            idxCheck = InStr(Data(i), sStart)
            If idxCheck <> 0 Then
                BAD = True
                GoTo NextIteration:
            Else
                results(j) = Data(i)
                j = j + 1
            End If
        Else
            idxCheck = InStr(Data(i), sEnd)
            If idxCheck <> 0 Then
                'i = i + DeleteRowsAfter
                BAD = False
            End If
       End If
NextIteration:
    Next i
    remove_between = single_dim_shrink_array(results, j - 1)
    Erase Data
End Function

Public Function new_smaller_array(ByVal Data As Variant, ByVal rows As Long, ByVal cols As Long) As Variant
    ' // Original new_smaller_array
    Dim i As Long
    Dim j As Long
    Dim newArr() As Variant
    
    On Error GoTo ErrHandler:
    If UBound(Data) = rows Then
        new_smaller_array = Data
    End If
    
    ' NOTE index starts at 1
    ReDim newArr(1 To rows, 1 To cols)
    For i = 1 To rows
        For j = 1 To cols
            newArr(i, j) = Data(i, j)
        Next j
    Next i
    new_smaller_array = newArr
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("new_smaller_array", Err, True)
End Function

'Public Function new_smaller_array(ByVal data As Variant, ByVal rows&, ByVal cols&) As Variant
'    ' // Newest new_smaller_array: this one was giving problems with QUERY_
'    Dim result() As Variant
'    Dim i&, j&
'    Dim d_row_start_num&, d_col_start_num&
'
'    On Error GoTo ErrHandler:
'    ' If the number of rows said to shrink it to
'    ' is the same as the current size, we're done!
'    ' WHAT ABOUT THE COLUMNS?!??!
'    If UBound(data) = rows Then
'        new_smaller_array = data
'    End If
'
'    ReDim result(0 To rows, 0 To cols)
'    d_row_start_num = LBound(data, 1)
'    d_col_start_num = LBound(data, 2)
'
'    For i = 0 To rows - d_row_start_num
'        For j = 0 To cols - d_col_start_num
'            result(i, j) = data(i + d_row_start_num, j + d_col_start_num)
'        Next j
'    Next i
'    new_smaller_array = result
'Done:
'    Exit Function
'ErrHandler:
'    Call LOGGER_.print_error("new_smaller_array", Err, True)
'End Function
