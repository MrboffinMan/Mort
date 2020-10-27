VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FILE_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class Name            : FILE_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                                                 with Files
' ==========================================================================

Public Sub create(ByVal filepath$)
    On Error GoTo ErrHandler:
    ' // [filepath]: include filename & extension of file to create.
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CreateTextFile(filepath)
    Set fso = Nothing
Done:
    Exit Sub
ErrHandler:
    Call LOGGER_.print_error("FILE_::create", Err, True)
End Sub

Public Function exists(ByVal filepath$) As Boolean
        On Error Resume Next
        exists = STRING_.Length(Dir$(filepath))
End Function

Public Sub copy(ByVal Source$, ByVal destination$, Optional overwrite As Boolean = True)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(Source, destination, overwrite)
    Set fso = Nothing
End Sub

Public Sub move(ByVal Source$, ByVal destination$)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.MoveFile(Source, destination)
    Set fso = Nothing
End Sub

Public Sub delete(ByVal filepath$, Optional force As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.DeleteFile(filepath, force)
    Set fso = Nothing
End Sub
Public Function read_file_to_single_dim_array(ByVal filepath$) As Variant
    Const CHUNK_SIZE& = 100000
    Dim result() As Variant
    Dim in_stream&: in_stream = FreeFile
    Dim i&: i = 0
    Dim record$
    
    ReDim result(0 To CHUNK_SIZE)
    Open filepath For Input As #in_stream
    Do While Not EOF(in_stream)
        If i > UBound(result) Then
            ReDim Preserve result(LBound(result) To UBound(result) + CHUNK_SIZE)
        End If
        Line Input #in_stream, record
        result(i) = record
        i = i + 1
    Loop
    Close #in_stream
    read_file_to_single_dim_array = ARRAY_.single_dim_shrink_array(result, i - 1)
End Function

Public Function pdf_read(ByVal filepath$) As String
   Dim objAVDoc As New AcroAVDoc
   Dim objPDDoc As New AcroPDDoc
   Dim objPage As AcroPDPage
   Dim objSelection As AcroPDTextSelect
   Dim objHighlight As AcroHiliteList
   Dim pageNum&, tCount&
   Dim strText$
    
   strText = ""
   If objAVDoc.Open(filepath, "") Then
      Set objPDDoc = objAVDoc.GetPDDoc
      For pageNum = 0 To objPDDoc.GetNumPages() - 1
         Set objPage = objPDDoc.AcquirePage(pageNum)
         Set objHighlight = New AcroHiliteList
         objHighlight.Add 0, 10000 ' Adjust this up if it's not getting all the text on the page
         Set objSelection = objPage.CreatePageHilite(objHighlight)

         If Not objSelection Is Nothing Then
            For tCount = 0 To objSelection.GetNumText - 1
                ' // Possibly modify this to work with Stringbuilder
                ' // to increase its speed if there's every a problem...
               strText = strText & objSelection.GetText(tCount)
            Next tCount
         End If
      Next pageNum
      objAVDoc.Close 1
   End If

   pdf_read = strText
End Function

Public Function single_file_select_dialog() As String
    Dim intChoice%
    Dim strPath$
    
    On Error GoTo ErrHandler:
    ' Only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' Make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    ' Determine what choice the user made
    If intChoice <> 0 Then
        ' get the file path selected by user
        strPath = CStr(Application.FileDialog(msoFileDialogOpen).SelectedItems(1))
        single_file_select_dialog = strPath
    End If
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("single_file_select_dialog", Err, True)
End Function

Public Function multiple_file_select_dialog() As String()
    Dim intChoice%, numSelectedFiles%, i%
    Dim results() As String
    
    On Error GoTo ErrHandler:
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        numSelectedFiles = Application.FileDialog(msoFileDialogOpen).SelectedItems.count
        If numSelectedFiles = 1 Then
            ReDim results(0)
            results(0) = CStr(Application.FileDialog(msoFileDialogOpen).SelectedItems(1))
            multiple_file_select_dialog = results
        Else
            ReDim results(0 To numSelectedFiles - 1)
            For i = LBound(results) To UBound(results)
                results(i) = CStr(Application.FileDialog(msoFileDialogOpen).SelectedItems(i + 1))
            Next i
        End If
    Else
        ReDim results(0)
        results(0) = "No Files Selected"
    End If
    multiple_file_select_dialog = results
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("multiple_file_select_dialog", Err, True)
End Function


