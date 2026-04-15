Option Explicit

' ============================================================
' GPUBufferManager.cls  FIXED
' VBO/IBO ALLOCATION LAYER
' FIXES:
'   - glGenBuffers 1, ... called without module qualifier.
'     In VBA this resolves to nothing (no such unqualified function exists)
'     causing a silent compile error or wrong resolution.
'     Fixed to GL.glGenBuffers.
'   - glDeleteBuffers similarly qualified.
' ============================================================

Private Type BufferPair
    vbo        As Long
    IBO        As Long
    indexCount As Long
End Type

Private m_Buffers() As BufferPair
Private m_Count     As Long

Private Sub Class_Initialize()
    m_Count = 0
    ReDim m_Buffers(0 To 0)
End Sub

Public Function CreateBuffer() As Long
    Dim id As Long: id = m_Count
    m_Count = m_Count + 1
    ReDim Preserve m_Buffers(0 To m_Count - 1)

    ' FIXED: was bare glGenBuffers - now GL.glGenBuffers
    GL.glGenBuffers 1, m_Buffers(id).vbo
    GL.glGenBuffers 1, m_Buffers(id).IBO

    CreateBuffer = id
End Function

Public Function GetVBO(ByVal id As Long) As Long
    If id >= 0 And id < m_Count Then GetVBO = m_Buffers(id).vbo
End Function

Public Function GetIBO(ByVal id As Long) As Long
    If id >= 0 And id < m_Count Then GetIBO = m_Buffers(id).IBO
End Function

Public Sub SetIndexCount(ByVal id As Long, ByVal count As Long)
    If id >= 0 And id < m_Count Then m_Buffers(id).indexCount = count
End Sub

Public Function GetIndexCount(ByVal id As Long) As Long
    If id >= 0 And id < m_Count Then GetIndexCount = m_Buffers(id).indexCount
End Function

Public Sub DestroyBuffer(ByVal id As Long)
    If id < 0 Or id >= m_Count Then Exit Sub
    ' FIXED: qualified GL.glDeleteBuffers
    If m_Buffers(id).vbo <> 0 Then GL.glDeleteBuffers 1, m_Buffers(id).vbo: m_Buffers(id).vbo = 0
    If m_Buffers(id).IBO <> 0 Then GL.glDeleteBuffers 1, m_Buffers(id).IBO: m_Buffers(id).IBO = 0
End Sub

Public Sub DestroyAll()
    Dim i As Long
    For i = 0 To m_Count - 1: DestroyBuffer i: Next i
    m_Count = 0
    ReDim m_Buffers(0 To 0)
End Sub

Private Sub Class_Terminate()
    DestroyAll
End Sub
