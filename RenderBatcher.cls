Option Explicit
' ============================================================
' RenderBatcher.cls  v8.1 (FIXED - AddQuad + Destroy added)
' GPU Batch Builder (VBA OpenGL v7 Engine Core)
' ============================================================

' Internal quad vertex layout: X Y U V R G B A (8 Singles per vertex, 4 verts per quad = 32)
Private Type DrawCommand
    vao        As Long
    indexCount As Long
End Type

Private Type QuadVertex
    x As Single: y As Single
    u As Single: v As Single
    r As Single: g As Single: b As Single: a As Single
End Type

' --- Mesh draw queue ---
Private m_Cmds()   As DrawCommand
Private m_Count    As Long

' --- Quad batch ---
Private m_Quads()  As QuadVertex
Private m_QVerts   As Long          ' number of vertices accumulated
Private m_VAO      As Long
Private m_VBO      As Long
Private m_IBO      As Long
Private m_Ready    As Boolean

Private Const MAX_QUADS As Long = 4096
Private Const VERTS_PER_QUAD As Long = 4
Private Const IDX_PER_QUAD   As Long = 6

' ============================================================
' MESH DRAW QUEUE (used by RenderPipeline / BVH culling)
' ============================================================
Public Sub BeginFrame()
    m_Count = 0
    m_QVerts = 0
    ReDim m_Cmds(0 To 0)
End Sub

Public Sub Submit(ByVal vao As Long, ByVal indexCount As Long)
    m_Count = m_Count + 1
    ReDim Preserve m_Cmds(0 To m_Count - 1)
    m_Cmds(m_Count - 1).vao = vao
    m_Cmds(m_Count - 1).indexCount = indexCount
End Sub

Public Sub Flush()
    Dim i As Long
    ' 1. Draw queued meshes
    For i = 0 To m_Count - 1
        GL.glBindVertexArray m_Cmds(i).vao
        GL.glDrawElements GL.GL_TRIANGLES, m_Cmds(i).indexCount, GL.GL_UNSIGNED_INT, 0
    Next i
    m_Count = 0

    ' 2. Flush any accumulated quads
    If m_QVerts > 0 Then FlushQuads
End Sub

' ============================================================
' QUAD BATCH (used by SpriteRendererSystem)
' ============================================================
Public Sub AddQuad(ByVal x As Single, ByVal y As Single, _
                   ByVal w As Single, ByVal h As Single, _
                   ByVal UVX As Single, ByVal UVY As Single, _
                   ByVal UVW As Single, ByVal UVH As Single, _
                   ByVal r As Single, ByVal g As Single, _
                   ByVal b As Single, ByVal a As Single)

    If Not m_Ready Then InitQuadGPU
    If m_QVerts + VERTS_PER_QUAD > MAX_QUADS * VERTS_PER_QUAD Then FlushQuads

    ' Ensure array is large enough
    If m_QVerts + VERTS_PER_QUAD > UBound(m_Quads) + 1 Then
        ReDim Preserve m_Quads(0 To m_QVerts + VERTS_PER_QUAD + 63)
    End If

    Dim base As Long: base = m_QVerts

    ' Bottom-left
    m_Quads(base + 0).x = x:     m_Quads(base + 0).y = y
    m_Quads(base + 0).u = UVX:   m_Quads(base + 0).v = UVY + UVH
    m_Quads(base + 0).r = r: m_Quads(base + 0).g = g
    m_Quads(base + 0).b = b: m_Quads(base + 0).a = a

    ' Bottom-right
    m_Quads(base + 1).x = x + w: m_Quads(base + 1).y = y
    m_Quads(base + 1).u = UVX + UVW: m_Quads(base + 1).v = UVY + UVH
    m_Quads(base + 1).r = r: m_Quads(base + 1).g = g
    m_Quads(base + 1).b = b: m_Quads(base + 1).a = a

    ' Top-right
    m_Quads(base + 2).x = x + w: m_Quads(base + 2).y = y + h
    m_Quads(base + 2).u = UVX + UVW: m_Quads(base + 2).v = UVY
    m_Quads(base + 2).r = r: m_Quads(base + 2).g = g
    m_Quads(base + 2).b = b: m_Quads(base + 2).a = a

    ' Top-left
    m_Quads(base + 3).x = x:     m_Quads(base + 3).y = y + h
    m_Quads(base + 3).u = UVX:   m_Quads(base + 3).v = UVY
    m_Quads(base + 3).r = r: m_Quads(base + 3).g = g
    m_Quads(base + 3).b = b: m_Quads(base + 3).a = a

    m_QVerts = m_QVerts + VERTS_PER_QUAD
End Sub

Private Sub FlushQuads()
    If m_QVerts = 0 Then Exit Sub
    If Not m_Ready Then Exit Sub

    GL.glBindVertexArray m_VAO
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBO

    ' Upload accumulated quad vertices (stride = 8 singles = 32 bytes)
    Dim byteSize As LongPtr
    byteSize = m_QVerts * 32
    GL.glBufferData GL.GL_ARRAY_BUFFER, byteSize, VarPtr(m_Quads(0)), GL.GL_DYNAMIC_DRAW

    Dim quadCount As Long: quadCount = m_QVerts \ VERTS_PER_QUAD
    GL.glDrawElements GL.GL_TRIANGLES, quadCount * IDX_PER_QUAD, GL.GL_UNSIGNED_INT, 0
    GL.glBindVertexArray 0

    m_QVerts = 0
End Sub

Private Sub InitQuadGPU()
    If m_Ready Then Exit Sub

    ReDim m_Quads(0 To MAX_QUADS * VERTS_PER_QUAD - 1)

    ' Build index buffer: 2 triangles per quad = 6 indices
    Dim indices() As Long
    ReDim indices(0 To MAX_QUADS * IDX_PER_QUAD - 1)
    Dim q As Long, b As Long
    For q = 0 To MAX_QUADS - 1
        b = q * IDX_PER_QUAD
        Dim v As Long: v = q * VERTS_PER_QUAD
        indices(b) = v
        indices(b + 1) = v + 1
        indices(b + 2) = v + 2
        indices(b + 3) = v + 2
        indices(b + 4) = v + 3
        indices(b + 5) = v
    Next q

    GL.glGenVertexArrays 1, m_VAO
    GL.glGenBuffers 1, m_VBO
    GL.glGenBuffers 1, m_IBO

    GL.glBindVertexArray m_VAO

    ' VBO (dynamic)
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBO
    GL.glBufferData GL.GL_ARRAY_BUFFER, MAX_QUADS * VERTS_PER_QUAD * 32, 0, GL.GL_DYNAMIC_DRAW

    ' IBO (static - indices never change)
    GL.glBindBuffer GL.GL_ELEMENT_ARRAY_BUFFER, m_IBO
    GL.glBufferData GL.GL_ELEMENT_ARRAY_BUFFER, MAX_QUADS * IDX_PER_QUAD * 4, VarPtr(indices(0)), GL.GL_STATIC_DRAW

    ' Layout: loc=0 XY (2f), loc=1 UV (2f), loc=2 RGBA (4f) — stride 32 bytes
    GL.glEnableVertexAttribArray 0: GL.glVertexAttribPointer 0, 2, GL.GL_FLOAT, 0, 32, 0
    GL.glEnableVertexAttribArray 1: GL.glVertexAttribPointer 1, 2, GL.GL_FLOAT, 0, 32, 8
    GL.glEnableVertexAttribArray 2: GL.glVertexAttribPointer 2, 4, GL.GL_FLOAT, 0, 32, 16

    GL.glBindVertexArray 0
    m_Ready = True
End Sub

' ============================================================
' CLEANUP
' ============================================================
Public Sub Destroy()
    If m_VBO <> 0 Then GL.glDeleteBuffers 1, m_VBO
    If m_IBO <> 0 Then GL.glDeleteBuffers 1, m_IBO
    If m_VAO <> 0 Then GL.glDeleteVertexArrays 1, m_VAO
    m_Ready = False
    m_Count = 0
    m_QVerts = 0
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub


