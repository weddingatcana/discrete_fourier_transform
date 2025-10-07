'---------------------------------------------------------------------------------------
' 64-bit compatible WinAPI declarations and constants for the terminal window.
'---------------------------------------------------------------------------------------
Option Explicit

#If VBA7 Then
    ' For 64-bit Office
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (ByRef lpwcx As WNDCLASSEX) As Integer
    Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (ByRef lpMsg As MSG, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
    Private Declare PtrSafe Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG) As Long
    Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Long
    Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Private Declare PtrSafe Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
#Else
    ' For 32-bit Office
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As Long) As Long
    Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (ByRef lpwcx As WNDCLASSEX) As Integer
    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (ByRef lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
    Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG) As Long
    Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Long
    Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
    Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
#End If

' --- Structures (Type Definitions) ---
Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As LongPtr
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As LongPtr
    hIcon As LongPtr
    hCursor As LongPtr
    hbrBackground As LongPtr
    lpszMenuName As String
    lpszClassName As String
    hIconSm As LongPtr
End Type

Private Type MSG
    hwnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    pt_x As Long
    pt_y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' --- Constants ---
Private Const CLASS_NAME As String = "VBATerminalClass"
Private Const CS_HREDRAW As Long = &H2
Private Const CS_VREDRAW As Long = &H1
Private Const CW_USEDEFAULT As Long = &H80000000
Private Const WS_OVERLAPPEDWINDOW As Long = &HCF0000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VSCROLL = &H200000
Private Const SW_SHOWNORMAL As Long = 1
Private Const WM_CREATE         As Long = &H1
Private Const WM_DESTROY        As Long = &H2
Private Const WM_SIZE           As Long = &H5
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_CTLCOLOREDIT   As Long = &H133
Private Const ES_MULTILINE As Long = &H4
Private Const ES_AUTOVSCROLL As Long = &H40
Private Const VK_RETURN As Long = &HD
Private Const GWLP_WNDPROC As Long = -4
Private Const EM_SETSEL As Long = &HB1

' --- Global Variables ---
Private g_hwndMain As LongPtr
Private g_hwndEdit As LongPtr
Private g_pOldEditProc As LongPtr
Private g_sPrompt As String
Private g_hBrush As LongPtr

'========================================================================================
' Main Subroutine - Entry point for creating the terminal window
'========================================================================================
Public Sub CreateTerminal()
    Dim wc As WNDCLASSEX
    Dim msg As MSG
    Dim hInstance As LongPtr
    
    ' Get the application instance handle
    hInstance = GetModuleHandle(0)
    
    ' Define the window class
    wc.cbSize = Len(wc)
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnWndProc = AddressOf WndProc ' Set the callback procedure
    wc.hInstance = hInstance
    wc.lpszClassName = CLASS_NAME
    
    ' Register the window class
    If RegisterClassEx(wc) = 0 Then
        MsgBox "Failed to register window class."
        Exit Sub
    End If
    
    ' Create the main window
    g_hwndMain = CreateWindowEx(0, CLASS_NAME, "VBA Terminal", WS_OVERLAPPEDWINDOW Or WS_VISIBLE, _
                              CW_USEDEFAULT, CW_USEDEFAULT, 600, 400, 0, 0, hInstance, 0)
    
    If g_hwndMain = 0 Then
        MsgBox "Failed to create window."
        Exit Sub
    End If
    
    ShowWindow g_hwndMain, SW_SHOWNORMAL
    UpdateWindow g_hwndMain
    
    ' Message loop to process window events
    Do While GetMessage(msg, 0, 0, 0)
        TranslateMessage msg
        DispatchMessage msg
    Loop
End Sub

'========================================================================================
' Main Window Procedure - Handles messages for the parent window
'========================================================================================
Private Function WndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim rcClient As RECT
    
    Select Case uMsg
        Case WM_CREATE
            ' Get the username and create the prompt string
            g_sPrompt = GetUsernamePrompt()
            
            ' Create a brush for the grey background
            g_hBrush = CreateSolidBrush(RGB(45, 45, 45)) ' Dark grey background
            
            ' Get the client area dimensions
            GetClientRect hwnd, rcClient
            
            ' Create the Edit control to fill the entire client area
            g_hwndEdit = CreateWindowEx(0, "EDIT", "", _
                WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL, _
                0, 0, rcClient.Right - rcClient.Left, rcClient.Bottom - rcClient.Top, _
                hwnd, 0, GetModuleHandle(0), 0)
                
            ' Subclass the edit control to intercept its messages
            g_pOldEditProc = SetWindowLongPtr(g_hwndEdit, GWLP_WNDPROC, AddressOf EditProc)
            
            ' Set the initial prompt
            SetWindowText g_hwndEdit, g_sPrompt
            ' Move cursor to the end
            SendMessage g_hwndEdit, EM_SETSEL, -1, -1
            
        Case WM_CTLCOLOREDIT
            ' This message is sent by the edit control before it paints itself
            SetTextColor wParam, RGB(0, 255, 0)       ' Green text
            SetBkColor wParam, RGB(45, 45, 45)       ' Dark grey background
            WndProc = g_hBrush                      ' Return the brush handle for the background
            Exit Function
            
        Case WM_SIZE
            ' Resize the edit control when the main window is resized
            GetClientRect hwnd, rcClient
            MoveWindow g_hwndEdit, 0, 0, rcClient.Right, rcClient.Bottom, True
            
        Case WM_DESTROY
            ' Clean up resources
            DeleteObject g_hBrush
            ' Restore original edit procedure
            SetWindowLongPtr g_hwndEdit, GWLP_WNDPROC, g_pOldEditProc
            PostQuitMessage 0
            
        Case Else
            ' Handle all other messages with the default procedure
            WndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
            Exit Function
    End Select
End Function

'========================================================================================
' Subclassed Edit Control Procedure - Handles messages specifically for the Edit control
'========================================================================================
Private Function EditProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If uMsg = WM_KEYDOWN And wParam = VK_RETURN Then
        ' User pressed the Enter key
        Dim textLen As Long
        Dim sText As String
        Dim sLastLine As String
        Dim sCommand As String
        
        ' Get the current text from the edit control
        textLen = GetWindowTextLength(hwnd)
        sText = String(textLen + 1, 0)
        GetWindowText hwnd, sText, textLen + 1
        sText = Left(sText, textLen) ' Trim null terminator
        
        ' Extract the last line to get the command
        If InStrRev(sText, vbCrLf) > 0 Then
            sLastLine = Mid(sText, InStrRev(sText, vbCrLf) + 2)
        Else
            sLastLine = sText
        End If
        
        ' Isolate the command from the prompt
        sCommand = Trim(Replace(sLastLine, g_sPrompt, ""))
        
        If LCase(sCommand) = "clear" Then
            ' If the command is "clear", reset the window text to the prompt
            SetWindowText hwnd, g_sPrompt
        Else
            ' For any other command, add a new line and a new prompt
            SetWindowText hwnd, sText & vbCrLf & g_sPrompt
        End If
        
        ' Move cursor to the end of the text
        SendMessage hwnd, EM_SETSEL, -1, -1
        
        ' We handled the message, so return 0
        EditProc = 0
        Exit Function
    End If
    
    ' For all other messages, pass them to the original edit control procedure
    EditProc = CallWindowProc(g_pOldEditProc, hwnd, uMsg, wParam, lParam)
End Function

'========================================================================================
' Helper Function - Gets the current user's name and formats the prompt
'========================================================================================
Private Function GetUsernamePrompt() As String
    Dim buffer As String * 256
    Dim bufferLen As Long
    bufferLen = 255
    If GetUserName(buffer, bufferLen) <> 0 Then
        ' Success
        GetUsernamePrompt = Left(buffer, bufferLen - 1) & "$: "
    Else
        ' Failure
        GetUsernamePrompt = "user$: "
    End If
End Function
