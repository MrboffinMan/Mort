VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STRING_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class Name            : STRING_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class contains methods for working with Strings
' --------------------------------------------------------------------------
' TODO:
'   - Think and then execute how I want to handle inputted string variables.
'       - Do I want to use ByVal || ByRef?
'       - Take into consideration: Strings may have a large footprint in size.
' ==========================================================================
                                
Public Function ascii_letters() As String
        ascii_letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
End Function

Public Function ascii_lowercase() As String
        ascii_lowercase = "abcdefghijklmnopqrstuvwxyz"
End Function

Public Function ascii_uppercase() As String
        ascii_uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
End Function

Public Function digits() As String
        digits = "0123456789"
End Function

Public Function hexdigits() As String
        hexdigits = "0123456789abcdefABCDEF"
End Function

Public Function octdigits() As String
        octdigits = "01234567"
End Function

Public Function capitalize(X As String) As String
    capitalize = UCase$(Left$(LCase$(X), 1)) & LCase$(Right$(X, Len(X) - 1))
End Function

Public Function Length(X As String) As Long
    Length = Len(X)
End Function

Public Function endswith(X As String, suffix As String) As Boolean
        endswith = Right$(X, STRING_.Length(suffix)) = suffix
End Function

Public Function startswith(X As String, prefix As String) As Boolean
        startswith = Left$(X, STRING_.Length(prefix)) = prefix
End Function

Public Function find_text(X As String, search_value As String, Optional start_pos As Long = 0) As Boolean
    find_text = Find(search_value, X, start_pos)
End Function

Public Function replace(ByVal X As String, sOld As String, sNew As String) As String
    replace = WorksheetFunction.Substitute$(X, sOld, sNew)
End Function

Public Function strip(X As String) As String
        strip = Trim$(X)
End Function

Public Function split(X As String, Optional delimiter As String = "", Optional limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String()
    split = VBA.split(X, delimiter, limit, Compare)
End Function

Public Function split_trim(ByVal X As String, Optional delimiter As String = "", Optional limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String()
        Dim Arr() As String
        Dim i&
        
        Arr = STRING_.split(X, delimiter, limit, Compare)
        For i = LBound(Arr) To UBound(Arr)
                Arr(i) = STRING_.strip(Arr(i))
        Next i
        split_trim = Arr
End Function

Public Function slice(X As Variant, ByVal first As Long, ByVal last As Long) As String
    ' Starts at index 0
    slice = Mid$(X, (first + 1), (last + 1) - (first + 1))
End Function

Public Function title(X As String) As String
        title = WorksheetFunction.Proper$(X)
End Function

Public Function upper(X As String) As String
        upper = UCase$(X)
End Function

Public Function truncate(X As String, max_length As Long) As String
    Const Sep$ = "..."
    
    Dim str_length&: str_length = STRING_.Length(X)
    If max_length <= str_length Then
        Dim str_cut&:   str_cut = WorksheetFunction.Floor((max_length - STRING_.Length(Sep)) / 2, 1)
        Dim str_front$: str_front = STRING_.slice(X, 0, str_cut)
        Dim str_back$:  str_back = STRING_.slice(X, str_length - str_cut, str_length)
        truncate = STRING_.concat(Sep, str_front, str_back)
    Else
        truncate = X
        Exit Function
    End If
End Function

Public Function concat(Sep As String, ParamArray Args()) As Variant
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Concat
        ' By Chip Pearson, chip@cpearson.com, www.cpearson.com
        '                  www.cpearson.com/Excel/Concatenation.aspx
        ' This function concatenates all the elements in the Args array,
        ' delimited by the Sep character, into a single string. This function
        ' can be used in an array formula. There is a VBA imposed limit that
        ' a string in a passed in array (e.g.,  calling this function from
        ' an array formula in a worksheet cell) must be less than 256 characters.
        ' See the comments at STRING TOO LONG HANDLING for details.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim s As String
        Dim N As Long
        Dim M As Long
        Dim r As Range
        Dim NumDims As Long
        Dim LB As Long
        Dim IsArrayAlloc As Boolean

        '''''''''''''''''''''''''''''''''''''''''''
        ' If no parameters were passed in, return
        ' vbNullString.
        '''''''''''''''''''''''''''''''''''''''''''
        If UBound(Args) - LBound(Args) + 1 = 0 Then
                concat = vbNullString
                Exit Function
        End If

        For N = LBound(Args) To UBound(Args)
                ''''''''''''''''''''''''''''''''''''''''''''''''
                ' Loop through the Args
                ''''''''''''''''''''''''''''''''''''''''''''''''
                If IsObject(Args(N)) = True Then
                        '''''''''''''''''''''''''''''''''''''
                        ' OBJECT
                        ' If we have an object, ensure it
                        ' it a Range. The Range object
                        ' is the only type of object we'll
                        ' work with. Anything else causes
                        ' a #VALUE error.
                        ''''''''''''''''''''''''''''''''''''
                        If TypeOf Args(N) Is Excel.Range Then
                                '''''''''''''''''''''''''''''''''''''''''
                                ' If it is a Range, loop through the
                                ' cells and create append the elements
                                ' to the string S.
                                '''''''''''''''''''''''''''''''''''''''''
                                For Each r In Args(N).Cells
                                        If Len(r.text) > 0 Then
                                                s = s & r.text & Sep
                                        End If
                                Next r
                        Else
                                '''''''''''''''''''''''''''''''''
                                ' Unsupported object type. Return
                                ' a #VALUE error.
                                '''''''''''''''''''''''''''''''''
                                concat = CVErr(xlErrValue)
                                Exit Function
                        End If
                
                ElseIf IsArray(Args(N)) = True Then
                        '''''''''''''''''''''''''''''''''''''
                        ' ARRAY
                        ' If Args(N) is an array, ensure it
                        ' is an allocated array.
                        '''''''''''''''''''''''''''''''''''''
                        IsArrayAlloc = (Not IsError(LBound(Args(N))) And _
                                (LBound(Args(N)) <= UBound(Args(N))))
                        If IsArrayAlloc = True Then
                                ''''''''''''''''''''''''''''''''''''
                                ' The array is allocated. Determine
                                ' the number of dimensions of the
                                ' array.
                                '''''''''''''''''''''''''''''''''''''
                                NumDims = 1
                                On Error Resume Next
                                Err.Clear
                                NumDims = 1
                                Do Until Err.Number <> 0
                                        LB = LBound(Args(N), NumDims)
                                        If Err.Number = 0 Then
                                                NumDims = NumDims + 1
                                        Else
                                                NumDims = NumDims - 1
                                        End If
                                Loop
                                On Error GoTo 0
                                Err.Clear
                                ''''''''''''''''''''''''''''''''''
                                ' The array must have either
                                ' one or two dimensions. Greater
                                ' that two caues a #VALUE error.
                                ''''''''''''''''''''''''''''''''''
                                If NumDims > 2 Then
                                        concat = CVErr(xlErrValue)
                                        Exit Function
                                End If
                                If NumDims = 1 Then
                                        For M = LBound(Args(N)) To UBound(Args(N))
                                                If Args(N)(M) <> vbNullString Then
                                                        s = s & Args(N)(M) & Sep
                                                End If
                                        Next M
                                        
                                Else
                                        ''''''''''''''''''''''''''''''''''''''''''''''''
                                        ' STRING TOO LONG HANDLING
                                        ' Here, the error handler must be set to either
                                        '   On Error GoTo ContinueLoop
                                        '   or
                                        '   On Error GoTo ErrH
                                        ' If you use ErrH, then any error, including
                                        ' a string too long error, will cause the function
                                        ' to return #VALUE and quit. If you use ContinueLoop,
                                        ' the problematic value is ignored and not included
                                        ' in the result, and the result is the concatenation
                                        ' of all non-error values in the input. This code is
                                        ' used in the case that an input string is longer than
                                        ' 255 characters.
                                        ''''''''''''''''''''''''''''''''''''''''''''''''
                                        On Error GoTo ContinueLoop
                                        'On Error GoTo ErrH
                                        Err.Clear
                                        For M = LBound(Args(N), 1) To UBound(Args(N), 1)
                                                If Args(N)(M, 1) <> vbNullString Then
                                                        s = s & Args(N)(M, 1) & Sep
                                                End If
                                        Next M
                                        Err.Clear
                                        M = LBound(Args(N), 2)
                                        If Err.Number = 0 Then
                                                For M = LBound(Args(N), 2) To UBound(Args(N), 2)
                                                        If Args(N)(M, 2) <> vbNullString Then
                                                                s = s & Args(N)(M, 2) & Sep
                                                        End If
                                                Next M
                                        End If
                                        On Error GoTo ErrH:
                                End If
                        Else
                                If Args(N) <> vbNullString Then
                                        s = s & Args(N) & Sep
                                End If
                        End If
                        Else
                        On Error Resume Next
                        If Args(N) <> vbNullString Then
                                s = s & Args(N) & Sep
                        End If
                        On Error GoTo 0
                End If
ContinueLoop:
        Next N
        
        ' Remove the trailing Sep
        If Len(Sep) > 0 Then
                If Len(s) > 0 Then
                        s = Left(s, Len(s) - Len(Sep))
                End If
        End If
        
        concat = s
        ' Success. Get out.
        Exit Function
ErrH:
        ' Error. Return #VALUE
        concat = CVErr(xlErrValue)
End Function


Public Function pad_left(X As String, total_length As Long, pad_char As String) As String
    pad_left = VBA.String$(total_length - STRING_.Length(X), pad_char) & X
End Function

Public Function pad_right(X As String, total_length As Long, pad_char As String) As String
    pad_right = X & VBA.String$(total_length - STRING_.Length(X), pad_char)
End Function

Public Function IndexOf(ByVal X As String, search_value As String) As Long
    Dim result&: result = InStr(X, search_value)
    If result <> 0 Then
        IndexOf = result
    Else
        IndexOf = -1
    End If
End Function

Public Function push(X As String, character As String) As String
    push = X & character
End Function

Public Function pop(ByVal X As String) As String
    Dim i&: i = STRING_.Length(X) - 1
    pop = Left$(X, i)
End Function

Public Function count(X As String, search_value As String) As Long
    count = UBound(STRING_.split(X, search_value))
End Function
