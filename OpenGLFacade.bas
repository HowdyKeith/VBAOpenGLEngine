Option Explicit
' =========================================================
' OpenGLFacade.bas v4.3.3
' =========================================================

Public g_GL As OpenGLWindow

Public Sub Tick()
    If g_GL Is Nothing Then Exit Sub
    g_GL.Update
    g_GL.Render
    g_GL.Present
End Sub


Public Function Create(ByVal w As Long, ByVal h As Long, _
                  Optional ByVal Title As String = "VBA OpenGL Engine") As Boolean
    Set g_GL = New OpenGLWindow
    Create = g_GL.Create(Title, w, h)
    
    If Create Then
        Debug.Print "[Facade] Window + GL context created"
    Else
        Set g_GL = Nothing
    End If
End Function

Public Function IsRunning() As Boolean
    If g_GL Is Nothing Then Exit Function
    IsRunning = g_GL.IsRunning
End Function


Public Sub InitDemo()
    Debug.Print "[Facade] InitDemo stub - put your scene setup here"
End Sub

Public Sub Cleanup()
    If Not g_GL Is Nothing Then
        g_GL.CloseWindow
        Set g_GL = Nothing
    End If
End Sub




