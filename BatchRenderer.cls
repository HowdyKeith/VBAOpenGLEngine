Option Explicit

' =========================================================
' Class      : BatchRenderer
' Version    : v9.0  (VAO-First + Multi-Layout)
' Description: Collects and dispatches DrawCommands using VAOs only.
' =========================================================

Private Type DrawCommand
    vao As Long          ' ? only one ID needed now
    indexCount As Long
End Type

Private m_Cmd() As DrawCommand
Private m_Count As Long
Private m_Capacity As Long

Private Sub Class_Initialize()
    m_Capacity = 64
    ReDim m_Cmd(0 To m_Capacity - 1)
End Sub

Public Sub BeginFrame()
    m_Count = 0
End Sub

''' <summary>
''' Submit a fully-prepared MeshBuffer (VAO + layout already set)
''' </summary>
Public Sub Submit(ByVal vao As Long, ByVal indexCount As Long)
    If m_Count >= m_Capacity Then
        m_Capacity = m_Capacity * 2
        ReDim Preserve m_Cmd(0 To m_Capacity - 1)
    End If
    
    With m_Cmd(m_Count)
        .vao = vao
        .indexCount = indexCount
    End With
    
    m_Count = m_Count + 1
End Sub

Public Sub Flush()
    If m_Count = 0 Then Exit Sub
    
    Dim i As Long
    For i = 0 To m_Count - 1
        With m_Cmd(i)
            GL.glBindVertexArray .vao
            GL.glDrawElements GL_TRIANGLES, .indexCount, GL_UNSIGNED_INT, 0&
        End With
    Next i
    
    m_Count = 0
    GL.glBindVertexArray 0   ' good hygiene
End Sub

Private Sub Class_Terminate()
    Erase m_Cmd
End Sub
