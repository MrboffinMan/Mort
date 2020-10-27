VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name        : StringBuilder
' Module Type       : Class
' --------------------------------------------------------------------------
' Description       :
' ==========================================================================


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Windows API
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal Length As Long)

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Consumed Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Defined Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Enumerations & Types
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Variables
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Const DEFAULT_CAPACITY& = 16
Private string_buffer() As Byte
Private current_length&
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class: Initialization & Termination
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    ' // Each unicode char is 2 bytes
    ReDim string_buffer(0 To (DEFAULT_CAPACITY * 2) - 1)
End Sub
' ------------------
Private Sub Class_Terminate()
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Property Get Length() As Long
    Length = current_length * 0.5
End Property
' ------------------
Public Property Get Capacity() As Long
    Capacity = UBound(string_buffer)
End Property
' ------------------
Public Property Get ToString() As String
    ToString = Mid$(string_buffer, 1, Me.Length())
End Function
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Using Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Expand(Optional newSize& = 0)
    Select Case newSize
        Case Is = 0
            ReDim Preserve string_buffer(0 To (UBound(string_buffer) * 2) + 1)
        Case Else
            ReDim Preserve string_buffer(0 To newSize - 1)
    End Select
End Sub
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function Append(ByVal Str$) As StringBuilder
    On Error GoTo ErrHandler:
    
    Dim buffer&: buffer = UBound(string_buffer)
    Dim binary_length&: binary_length = LenB(Str)
    Dim memory_location As LongPtr: memory_location = StrPtr(Str)
    Dim new_length&: new_length = current_length + binary_length
    
    Select Case new_length
        Case Is < buffer
            CopyMemory VarPtr(string_buffer(current_length)), memory_location, binary_length
        Case Is < (buffer * 2)
            Expand
            CopyMemory VarPtr(string_buffer(current_length)), memory_location, binary_length
        Case Else
            Expand new_length
            CopyMemory VarPtr(string_buffer(current_length)), memory_location, binary_length
    End Select
    
    current_length = new_length
    Set Append = Me
Done:
    Exit Function
ErrHandler:
    Stop
    Resume
End Function
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' UserForm Builder Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
