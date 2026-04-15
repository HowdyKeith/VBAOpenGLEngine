Option Explicit

' ============================================================
' modWindowManager.v8.8.bas  FIXED
' WIN32 WINDOW & GL CONTEXT MANAGER
' FIXES:
'   - SetupPixelFormat was stubbed/commented out - now calls modWGLContext.InitOpenGL
'   - MSG struct was wrong size (array of Long instead of tagMSG) - now uses Win32GL.tagMSG
'   - WM_QUIT comparison used msg(1) on raw Long array - now uses proper struct
'   - CreateGLWindow no longer skips pixel format setup
'   - ShowWindow call added so the window is actually visible
' ============================================================

' --- INTERNAL STATE ---
Private m_hWnd As LongPtr
Private m_hDC  As LongPtr
Private m_hRC  As LongPtr

' --- PUBLIC GETTERS ---
Public Function GetLocalDC() As LongPtr:   GetLocalDC = m_hDC:   End Function
Public Function GetLocalRC() As LongPtr:   GetLocalRC = m_hRC:   End Function
Public Function GetLocalHWND() As LongPtr: GetLocalHWND = m_hWnd: End Function

' =========================
' WINDOW CREATION
' =========================
Public Function CreateGLWindow(ByVal Title As String, ByVal width As Long, ByVal Height As Long) As LongPtr

    ' 1. Register a proper window class (STATIC doesn't support OpenGL pixel formats)
    Dim wc As Win32GL.WNDCLASSEX
    wc.cbSize = Len(wc)
    wc.style = Win32GL.CS_HREDRAW Or Win32GL.CS_VREDRAW Or Win32GL.CS_OWNDC
    wc.lpfnWndProc = 0          ' Use DefWindowProc fallback (VBA can't set WndProc directly)
    wc.hInstance = 0
    wc.hCursor = Win32GL.LoadCursor(0, Win32GL.IDC_ARROW)
    wc.hbrBackground = 0
    wc.lpszClassName = StrPtr("VBAGLClass")
    Win32GL.RegisterClassEx wc

    ' 2. Create Window with clip flags required by OpenGL
    Dim style As Long
    style = Win32GL.WS_OVERLAPPEDWINDOW Or Win32GL.WS_CLIPSIBLINGS Or Win32GL.WS_CLIPCHILDREN
    m_hWnd = Win32GL.CreateWindowEx(0, "VBAGLClass", Title, style, 100, 100, width, Height, 0, 0, 0, 0)

    If m_hWnd = 0 Then
        Debug.Print "[WindowManager] Failed to create Window."
        Exit Function
    End If

    ' 3. Show the window
    Win32GL.ShowWindow m_hWnd, Win32GL.SW_SHOW

    ' 4. Set Pixel Format + Create GL Context (was previously stubbed out!)
    Dim ok As Boolean
    ok = modWGLContext.InitOpenGL(m_hWnd)

    If Not ok Then
        Debug.Print "[WindowManager] Failed to initialize OpenGL context."
        Win32GL.DestroyWindow m_hWnd
        m_hWnd = 0
        Exit Function
    End If

    ' 5. Mirror context handles from modWGLContext globals
    m_hDC = modWGLContext.g_hDC
    m_hRC = modWGLContext.g_hRC

    Debug.Print "[WindowManager] Window + GL context created OK."
    CreateGLWindow = m_hWnd
End Function

' =========================
' MESSAGE LOOP  (FIXED: uses proper tagMSG struct via Win32GL)
' =========================
Public Function ProcessMessages() As Boolean
    Dim msg As Win32GL.tagMSG

    Do While Win32GL.PeekMessage(msg, 0, 0, 0, Win32GL.PM_REMOVE) <> 0
        If msg.message = Win32GL.WM_QUIT Then
            ProcessMessages = False
            Exit Function
        End If
        Win32GL.TranslateMessage msg
        Win32GL.DispatchMessage msg
    Loop

    ProcessMessages = True
End Function

' =========================
' FRAME MANAGEMENT
' =========================
Public Sub PageFlip()
    If m_hDC <> 0 Then
        Win32GL.SwapBuffers m_hDC
    Else
        Debug.Print "[WindowManager] Warning: PageFlip failed - hDC is Null"
    End If
End Sub

' =========================
' CLEANUP
' =========================
Public Sub CloseGLWindow()
    modWGLContext.ShutdownOpenGL

    If m_hWnd <> 0 Then
        Win32GL.DestroyWindow m_hWnd
        m_hWnd = 0
    End If

    m_hDC = 0
    m_hRC = 0

    Debug.Print "[WindowManager] Context and Window Destroyed."
End Sub
