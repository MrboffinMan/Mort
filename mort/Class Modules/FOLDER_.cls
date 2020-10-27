VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FOLDER_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class Name            : FOLDER_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                                                 with Folders
' ==========================================================================

Public Function exists(ByVal folder_path$) As Boolean
        On Error Resume Next
        exists = (GetAttr(folder_path) And vbDirectory) = vbDirectory
End Function

Public Function list_files(ByVal folder_path$) As Variant
        Dim Arr As New ArrayList
        Dim last_char$
        Dim File$
        
        If Not exists(folder_path) Then
                Err.Raise Number:=vbObjectError + 76, Description:="Specified path does not exist"
                Exit Function
        End If
        
        last_char = Right$(folder_path, 1)
        
        If last_char <> "\" And last_char <> "/" Then
            If STRING_.count(folder_path, "\") > STRING_.count(folder_path, "/") Then
                folder_path = STRING_.push(folder_path, "\")
            Else
                folder_path = STRING_.push(folder_path, "/")
            End If
        End If
        
        File = Dir$(folder_path, vbReadOnly + vbHidden + vbSystem)
        Do While File <> ""
            Arr.Add (File)
            File = Dir()
        Loop
        list_files = Arr.toArray()
        Set Arr = Nothing
End Function
