VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ENVIRONMENT_User_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name            : ENVIRONMENT_Comp_Info
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                         with Environments => User Specific Information
' ==========================================================================
Public Function username() As String
    On Error GoTo ErrHandler:
    username = LCase$(Environ("username"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("username", Err, True)
End Function


Public Function computer_name() As String
    On Error GoTo ErrHandler:
    computer_name = LCase$(Environ("COMPUTERNAME"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("computer_name", Err, True)
End Function

Public Function user_domain() As String
    On Error GoTo ErrHandler:
    user_domain = LCase$(Environ("USERDOMAIN"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("user_domain", Err, True)
End Function

Public Function terminal_client_name() As String
    On Error GoTo ErrHandler:
    terminal_client_name = LCase$(Environ("ClientName"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("terminal_client_name", Err, True)
End Function

Public Function terminal_session_name() As String
    On Error GoTo ErrHandler:
    terminal_session_name = LCase$(Environ("SESSIONNAME"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("terminal_session_name", Err, True)
End Function

Public Function logon_server() As String
    On Error GoTo ErrHandler:
    logon_server = LCase$(Environ("LOGONSERVER"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("logon_server", Err, True)
End Function



