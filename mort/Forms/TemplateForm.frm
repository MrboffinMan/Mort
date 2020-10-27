VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "TemplateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' // ============================================= \\ '
' // Custom Events
' // ============================================= \\ '
Public Event Activate()
Public Event AddControl(ByVal Control As MSForms.Control)
Public Event BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event Click()
Public Event DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Public Event Deactivate()
Public Event Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
Public Event Initialize()
Public Event KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Public Event KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event Layout()
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event QueryClose(Cancel As Integer, CloseMode As Integer)
Public Event RemoveControl(ByVal Control As MSForms.Control)
Public Event Resize()
Public Event Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
Public Event Terminate()
Public Event Zoom(Percent As Integer)

' // ============================================= \\ '
' // Windows APIs
' // ============================================= \\ '
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal H_WINDOW As LongPtr, ByVal lngWinIdx As LongPtr, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal H_WINDOW As LongPtr, ByVal lngWinIdx As LongPtr) As LongPtr
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal H_WINDOW As LongPtr, ByVal crKey As Integer, ByVal bAlpha As Integer, ByVal dwFlags As LongPtr) As LongPtr
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal H_WINDOW As LongPtr) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal H_WINDOW As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal H_WINDOW As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Boolean) As LongPtr
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal H_WINDOW As LongPtr, lpPoint As POINTAPI) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (p As tCursor) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal HDC As LongPtr, ByVal nIndex As LongPtr) As LongPtr
Private Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi" (ByVal hWnd As LongPtr, ByVal attr As Integer, ByRef attrValue As Integer, ByVal attrSize As Integer) As LongPtr
Private Declare PtrSafe Function DwmExtendFrameIntoClientArea Lib "dwmapi" (ByVal hWnd As LongPtr, ByRef NEWMARGINS As MARGINS) As LongPtr

' // ============================================= \\ '
' // Variables for Windows APIs
' // ============================================= \\ '
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 '//// FRAME CHANGED SEND WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 '// DONT DO OWNER Z ORDERING
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Private Type MARGINS
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type tCursor
    Left As Long
    Top As Long
End Type

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000 '//// WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
Private Const GWL_EXSTYLE As Long = (-20) '//// OFFSET OF WINDOW EXTENDED STYLE
Private Const WS_EX_DLGMODALFRAME As Long = &H1 '//// CONTROLS IF WINDOW HAS AN ICON
Private Const SC_CLOSE As Long = &HF060
Private Const SW_SHOW As Long = 5
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_TRANSPARENT = &H20&
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private XWNDFORM, XWNDFORMEX As LongPtr

Public ENV_POS As Long
Public NEW_POS As Long
Public M_SNG_LEFT_POS, M_SNG_TOP_POS As Long

' // ============================================= \\ '
' // Factory method to get UserForm
' // ============================================= \\ '
Public Function GetTemplateForm(ByRef CallerWorkbook As Workbook) As TemplateForm
        Dim frm As TemplateForm: Set frm = New TemplateForm
        ' // set early-bound properties with intellisense
        frm.Caption = ""
        frm.Width = 300
        frm.Height = 270
        
        Set GetTemplateForm = frm
End Function


' // ============================================= \\ '
' // UserForm Events
' // ============================================= \\ '
Private Sub UserForm_Activate()
    RaiseEvent Activate
End Sub

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)
    RaiseEvent AddControl(Control)
End Sub

Private Sub UserForm_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent BeforeDragOver(Cancel, Control, Data, X, Y, State, Effect, Shift)
End Sub

Private Sub UserForm_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent BeforeDropOrPaste(Cancel, Control, Action, Data, X, Y, Effect, Shift)
End Sub

Private Sub UserForm_Click()
    RaiseEvent Click
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClick(Cancel)
End Sub

Private Sub UserForm_Deactivate()
    RaiseEvent Deactivate
End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    RaiseEvent Error(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub UserForm_Initialize()
    Dim ISTYLE, HWNDFORM As LongPtr
    Dim btrans As Byte
    btrans = 128
    Dim NEWMARGINS As MARGINS
    
    HWNDFORM = FindWindow(vbNullString, Me.Caption) '//// GET WINDOW

    ISTYLE = GetWindowLong(HWNDFORM, GWL_STYLE) '//// BASIC WINDOW STYLE FLAGS FOR THE FORM
    ISTYLE = ISTYLE And Not WS_CAPTION '//// NO CAPTION AREA
    SetWindowLong HWNDFORM, GWL_STYLE, ISTYLE '//// SET BASIC WINDOW STYLES

    ISTYLE = GetWindowLong(HWNDFORM, GWL_EXSTYLE) '//// BUILD EXTENDED WINDOW STYLE
    ISTYLE = ISTYLE And Not WS_EX_DLGMODALFRAME '//// NO BORDER

    'ISTYLE = ISTYLE Or WS_EX_LAYERED '//// ADD ONE COLOR TRANSPARENCE
    'ISTYLE = ISTYLE Or WS_EX_TRANSPARENT '//// ADD SEMI-TRANSPARENT WINDOW

    SetWindowLong HWNDFORM, GWL_EXSTYLE, ISTYLE

    'SetLayeredWindowAttributes HWNDFORM, vbCyan, btrans, LWA_ALPHA     '//// SEMI TRANSPARENT WINDOW
    'SetLayeredWindowAttributes HWNDFORM, vbCyan, btrans, LWA_COLORKEY  '//// COLOR SCREEN TRNSPARENCY

    XWNDFORM = FindWindow("ThunderDFrame", vbNullString) '//// GET NEW WINDOW

    DwmSetWindowAttribute XWNDFORM, 2, 2, 4 '//// DWMAPI

    With NEWMARGINS
        .Bottom = 0 '//// -1
        .Left = 0  '//// -1
        .Right = 0  '//// -1
        .Top = 1  '//// -1
    End With

    DwmExtendFrameIntoClientArea XWNDFORM, NEWMARGINS '//// DWMAPI
    
    DrawMenuBar HWNDFORM '//// CLEAN MENU BAR

    'RaiseEvent Initialize
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserForm_Layout()
    RaiseEvent Layout
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RaiseEvent QueryClose(Cancel, CloseMode)
End Sub

Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
    RaiseEvent RemoveControl(Control)
End Sub

Private Sub UserForm_Resize()
    RaiseEvent Resize
End Sub

Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    RaiseEvent Scroll(ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

Private Sub UserForm_Terminate()
    RaiseEvent Terminate
End Sub

Private Sub UserForm_Zoom(Percent As Integer)
    RaiseEvent Zoom(Percent)
End Sub

