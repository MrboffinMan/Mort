VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PERFORMANCE_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class Name            : PERFORMANCE_
' Module Type           : Class
' --------------------------------------------------------------------------
' Description           : This class is a container for methods dealing
'                         with Excel and its performance
' ==========================================================================

Public Sub turn_on()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
End Sub

Public Sub turn_off()
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
End Sub

