VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class Name            : OS
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for working with anything
'                         that is relateed to the OS, Folders, or Files
' ==========================================================================


' // Public variables that will allow access to OS, Files, Folders class methods
Public File As New FILE_
Public Folder As New FOLDER_
Public Env As New ENVIRONMENT_
Public Performance As New PERFORMANCE_
Public Arr As New ARRAY_
Public Str As New STRING_
Public Xml As New XML_
