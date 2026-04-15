Option Explicit

' =========================================================
' modWGLContext.v8.1.bas  FIXED
' WGL PIXEL FORMAT + OPENGL CONTEXT INITIALIZER
' FIXES:
'   - Duplicate Public Declare of GetDC removed (already in Win32GL.bas,
'     causes "Ambiguous name" compile error)
'   - ShutdownOpenGL now properly calls wglMakeCurrent(0,0) and wglDeleteContext
'     instead of just zeroing the globals
'   - SwapBuffers duplicate declaration removed (in Win32GL already)
'   - wglCreateContext/wglMakeCurrent duplicates removed - use Win32GL versions
' =========================================================

' =========================
' WGL CONSTANTS
' =========================
Public Const PFD_TYPE_RGBA         As Long = 0
Public Const PFD_DOUBLEBUFFER      As Long = &H1
Public Const PFD_DRAW_TO_WINDOW    As Long = &H4
Public Const PFD_SUPPORT_OPENGL    As Long = &H20
Public Const PFD_MAIN_PLANE        As Long = 0

Public Const WGL_CONTEXT_MAJOR_VERSION_ARB  As Long = &H2091
Public Const WGL_CONTEXT_MINOR_VERSION_ARB  As Long = &H2092
Public Const WGL_CONTEXT_FLAGS_ARB          As Long = &H2094
Public Const WGL_CONTEXT_PROFILE_MASK_ARB   As Long = &H9126
Public Const WGL_CONTEXT_CORE_PROFILE_BIT_ARB         As Long = &H1
Public Const WGL_CONTEXT_COMPATIBILITY_PROFILE_BIT_ARB As Long = &H2

' =========================
' PIXEL FORMAT DESCRIPTOR
' (must match Win32GL.bas PIXELFORMATDESCRIPTOR exactly)
' =========================
' NOTE: We reuse Win32GL.PIXELFORMATDESCRIPTOR to avoid duplicate type errors.
' The local Private Type below is only used inside InitOpenGL.

' =========================
' GLOBAL STATE (shared with modWindowManager)
' =========================
Public g_hDC As LongPtr
Public g_hRC As LongPtr

' =========================
' INIT OPENGL CONTEXT
' =========================
Public Function InitOpenGL(ByVal hWnd As LongPtr) As Boolean
    Dim pfd As Win32GL.PIXELFORMATDESCRIPTOR
    Dim pixelFmt As Long

    ' Get Device Context via Win32GL (avoids duplicate Declare)
    g_hDC = Win32GL.GetDC(hWnd)
    If g_hDC = 0 Then
        Debug.Print "[WGLContext] GetDC failed."
        InitOpenGL = False
        Exit Function
    End If

    ' Build pixel format descriptor
    With pfd
        .nSize      = Len(pfd)
        .nVersion   = 1
        .dwFlags    = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA
        .cColorBits = 32
        .cDepthBits = 24
        .cStencilBits = 8
        .iLayerType = PFD_MAIN_PLANE
    End With

    pixelFmt = Win32GL.ChoosePixelFormat(g_hDC, pfd)
    If pixelFmt = 0 Then
        Debug.Print "[WGLContext] ChoosePixelFormat failed."
        InitOpenGL = False
        Exit Function
    End If

    If Win32GL.SetPixelFormat(g_hDC, pixelFmt, pfd) = 0 Then
        Debug.Print "[WGLContext] SetPixelFormat failed."
        InitOpenGL = False
        Exit Function
    End If

    ' Create legacy GL context
    g_hRC = Win32GL.wglCreateContext(g_hDC)
    If g_hRC = 0 Then
        Debug.Print "[WGLContext] wglCreateContext failed."
        InitOpenGL = False
        Exit Function
    End If

    If Win32GL.wglMakeCurrent(g_hDC, g_hRC) = 0 Then
        Debug.Print "[WGLContext] wglMakeCurrent failed."
        Win32GL.wglDeleteContext g_hRC
        g_hRC = 0
        InitOpenGL = False
        Exit Function
    End If

    Debug.Print "[WGLContext] OpenGL context created and made current."
    InitOpenGL = True
End Function

' =========================
' SWAP BUFFER
' =========================
Public Sub PresentFrame()
    If g_hDC <> 0 Then Win32GL.SwapBuffers g_hDC
End Sub

' =========================
' SHUTDOWN  (FIXED: properly releases context)
' =========================
Public Sub ShutdownOpenGL()
    If g_hRC <> 0 Then
        Win32GL.wglMakeCurrent 0, 0
        Win32GL.wglDeleteContext g_hRC
        g_hRC = 0
    End If
    g_hDC = 0
    Debug.Print "[WGLContext] Shutdown complete."
End Sub
