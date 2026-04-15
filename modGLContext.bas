Option Explicit

' =========================================================
' modGLContext.v8.1.bas  FIXED
' OPENGL CONTEXT + EXTENSION LOADER
' FIXES:
'   - CreateWindowEx last param passed as literal 0 (Long) for LongPtr param -
'     changed to 0& for proper type
'   - GL_LoadExtensions now delegates entirely to GL.GL_Init (avoids duplicate
'     pointer loading and ensures GL.bas owns all extension pointers)
'   - GL_Shutdown now calls modWGLContext.ShutdownOpenGL for proper teardown
' =========================================================

' NOTE: All API Declares live in Win32GL.bas

' =========================================================
' INIT OPENGL CONTEXT
' =========================================================
Public Function GL_InitContext(Optional ByVal width As Long = 800, _
                               Optional ByVal Height As Long = 600, _
                               Optional ByVal Title As String = "VBA OpenGL v8") As Boolean

    ' 1. Create window
    Win32GL.g_hWnd = Win32GL.CreateWindowEx(0, "STATIC", Title, 0, 0, 0, width, Height, 0, 0, 0, 0&)

    If Win32GL.g_hWnd = 0 Then
        Debug.Print "[GLContext] CreateWindowEx failed."
        GL_InitContext = False
        Exit Function
    End If

    ' 2. Init pixel format + GL context via WGLContext module
    Dim ok As Boolean
    ok = modWGLContext.InitOpenGL(Win32GL.g_hWnd)
    If Not ok Then
        GL_InitContext = False
        Exit Function
    End If

    ' Mirror handles into Win32GL globals
    Win32GL.g_hDC = modWGLContext.g_hDC
    Win32GL.g_hRC = modWGLContext.g_hRC

    ' 3. Load all extension pointers (delegates to GL.bas - single source of truth)
    GL.GL_Init

    GL_InitContext = True
End Function

' =========================================================
' SWAP / FRAME
' =========================================================
Public Sub GL_SwapBuffers()
    If Win32GL.g_hDC <> 0 Then Win32GL.SwapBuffers Win32GL.g_hDC
End Sub

' =========================================================
' SHUTDOWN
' =========================================================
Public Sub GL_Shutdown()
    modWGLContext.ShutdownOpenGL

    If Win32GL.g_hWnd <> 0 Then
        Win32GL.ReleaseDC Win32GL.g_hWnd, Win32GL.g_hDC
        Win32GL.DestroyWindow Win32GL.g_hWnd
    End If

    Win32GL.g_hRC  = 0
    Win32GL.g_hDC  = 0
    Win32GL.g_hWnd = 0
End Sub
