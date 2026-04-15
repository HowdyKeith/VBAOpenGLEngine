Option Explicit

' =========================================================
' Win32GL.bas (v8.0 Master)
' Unified Win32 API, GDI, and WGL Library for VBA OpenGL
' =========================================================

' --- TYPES ---
Public Type POINTAPI: x As Long: y As Long: End Type

Public Type tagMSG
    hWnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    pt As POINTAPI
End Type

Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer: nVersion As Integer: dwFlags As Long: iPixelType As Byte
    cColorBits As Byte: cRedBits As Byte: cRedShift As Byte: cGreenBits As Byte
    cGreenShift As Byte: cBlueBits As Byte: cBlueShift As Byte: cAlphaBits As Byte
    cAlphaShift As Byte: cAccumBits As Byte: cAccumRedBits As Byte: cAccumGreenBits As Byte
    cAccumBlueBits As Byte: cAccumAlphaBits As Byte: cDepthBits As Byte: cStencilBits As Byte
    cAuxBuffers As Byte: iLayerType As Byte: bReserved As Byte: dwLayerMask As Long
    dwVisibleMask As Long: dwDamageMask As Long
End Type

Public Type WNDCLASSEX
    cbSize As Long: style As Long: lpfnWndProc As LongPtr: cbClsExtra As Long
    cbWndExtra As Long: hInstance As LongPtr: hIcon As LongPtr: hCursor As LongPtr
    hbrBackground As LongPtr: lpszMenuName As LongPtr: lpszClassName As LongPtr: hIconSm As LongPtr
End Type

' --- API: USER32 (Windowing & Messages) ---
Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
Public Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageA" (ByRef lpMsg As tagMSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare PtrSafe Function TranslateMessage Lib "user32" (ByRef lpMsg As tagMSG) As Long
Public Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (ByRef lpMsg As tagMSG) As LongPtr
Public Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare PtrSafe Function PostQuitMessage Lib "user32" (ByVal nExitCode As Long) As Long
Public Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (ByRef lpwcx As WNDCLASSEX) As Long
Public Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long
Public Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String) As Long

' --- API: GDI32 & OPENGL32 ---
Public Declare PtrSafe Function wglCreateContext Lib "opengl32" (ByVal hDC As LongPtr) As LongPtr
Public Declare PtrSafe Function wglMakeCurrent Lib "opengl32" (ByVal hDC As LongPtr, ByVal hglrc As LongPtr) As Long
Public Declare PtrSafe Function wglDeleteContext Lib "opengl32" (ByVal hglrc As LongPtr) As Long
Public Declare PtrSafe Function wglGetProcAddress Lib "opengl32" (ByVal name As String) As LongPtr
Public Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As LongPtr, pfd As PIXELFORMATDESCRIPTOR) As Long
Public Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hDC As LongPtr, ByVal format As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Public Declare PtrSafe Function SwapBuffers Lib "gdi32" (ByVal hDC As LongPtr) As Long

' --- API: KERNEL32 (System & Time) ---
Public Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long

' --- CONSTANTS: PFD ---
Public Const PFD_DRAW_TO_WINDOW = &H4, PFD_SUPPORT_OPENGL = &H20, PFD_DOUBLEBUFFER = &H1, PFD_TYPE_RGBA = 0, PFD_MAIN_PLANE = 0

' --- CONSTANTS: WINDOW STYLE ---
Public Const WS_OVERLAPPEDWINDOW = &HCF0000, WS_CLIPSIBLINGS = &H4000000, WS_CLIPCHILDREN = &H2000000
Public Const CS_HREDRAW = &H2, CS_VREDRAW = &H1, CS_OWNDC = &H20
Public Const SW_SHOW = 5, IDC_ARROW = 32512

' --- CONSTANTS: MESSAGES & INPUT ---
Public Const PM_REMOVE = 1, WM_QUIT = &H12, WM_CLOSE = &H10, WM_DESTROY = &H2
Public Const VK_W = &H57, VK_A = &H41, VK_S = &H53, VK_D = &H44, VK_SPACE = &H20, VK_ESCAPE = &H1B

' --- SHARED GLOBALS ---
Public g_hWnd As LongPtr
Public g_hDC As LongPtr
Public g_hRC As LongPtr

' --- ENGINE UTILITIES ---

Public Function MakeStandardPFD() As PIXELFORMATDESCRIPTOR
    With MakeStandardPFD
        .nSize = Len(MakeStandardPFD): .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA: .cColorBits = 32: .cDepthBits = 24: .iLayerType = PFD_MAIN_PLANE
    End With
End Function

Public Function WndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Select Case uMsg
        Case WM_CLOSE: Win32GL.DestroyWindow hWnd
        Case WM_DESTROY: Win32GL.PostQuitMessage 0
        Case Else: WndProc = Win32GL.DefWindowProc(hWnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function GetTime() As Double
    Dim freq As Currency, counter As Currency
    QueryPerformanceFrequency freq
    QueryPerformanceCounter counter
    GetTime = CDbl(counter) / CDbl(freq)
End Function

Public Function IsKeyDown(ByVal key As Long) As Boolean
    IsKeyDown = (GetAsyncKeyState(key) And &H8000) <> 0
End Function

Public Function GetAddress(ByVal fn As LongPtr) As LongPtr: GetAddress = fn: End Function

Public Function PumpMessages() As Boolean
    Dim m As tagMSG
    Do While PeekMessage(m, 0, 0, 0, PM_REMOVE)
        If m.message = WM_QUIT Then PumpMessages = False: Exit Function
        TranslateMessage m
        DispatchMessage m
    Loop
    PumpMessages = True
End Function


