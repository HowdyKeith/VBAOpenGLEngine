Option Explicit

' ============================================================
' FirstPixel.bas  FIXED
' Simplest possible demo: one orange triangle on screen.
' FIXES:
'   - GL_DEPTH_TEST was unqualified (ambiguous between GL.bas and GLConstants.bas)
'     Now uses GL.GL_DEPTH_TEST
'   - GL_COLOR_BUFFER_BIT and GL_DEPTH_BUFFER_BIT now fully qualified
'   - m_Window, m_mesh, m_Shader moved to module-level (were already there, kept)
'   - Added proper cleanup order
' ============================================================

Private m_Window As OpenGLWindow
Private m_mesh   As TriangleMesh
Private m_Shader As ShaderProgram

Public Sub RunFirstPixel()
    Set m_Window = New OpenGLWindow

    If Not m_Window.Create("VBA OpenGL - First Pixel", 800, 600) Then
        MsgBox "Failed to create OpenGL window", vbCritical
        Exit Sub
    End If

    InitScene

    Do While m_Window.IsRunning()
        RenderFrame
        m_Window.Present
        DoEvents
    Loop

    ' Cleanup in reverse order
    If Not m_Shader Is Nothing Then m_Shader.Destroy
    If Not m_mesh   Is Nothing Then m_mesh.Destroy
    m_Window.Destroy

    Set m_Shader = Nothing
    Set m_mesh   = Nothing
    Set m_Window = Nothing
End Sub

Private Sub InitScene()
    GL.glClearColor 0.1, 0.1, 0.12, 1#
    GL.glEnable GL.GL_DEPTH_TEST

    Set m_mesh = New TriangleMesh
    m_mesh.UploadTriangle

    Set m_Shader = New ShaderProgram
    If Not m_Shader.CreateDefault() Then
        Debug.Print "[FirstPixel] Warning: Default shader failed to compile."
    End If
End Sub

Private Sub RenderFrame()
    GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT

    If Not m_Shader Is Nothing Then m_Shader.Use
    If Not m_mesh   Is Nothing Then m_mesh.Draw
End Sub
