VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LOGGER_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class Name            : LOGGER_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                                                 with Logging
' ==========================================================================

Public Sub print_error(sMethodName As String, vError As Object, LogToDB As Boolean)
    Dim oErrorObj As Object: Set oErrorObj = vError
    
    Debug.Print ("================================================")
    Debug.Print ("[Method]          => " & sMethodName)
    Debug.Print ("[Err Number]      => " & vError.Number)
    Debug.Print ("[Err Source]      => " & vError.Source)
    Debug.Print ("[Err Description] => " & vError.Description)
    Debug.Print ("================================================")
End Sub

Public Function clone_object(ByVal obj As Object) As Object
    Set clone_object = obj
End Function

