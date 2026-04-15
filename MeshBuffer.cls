Option Explicit

' ============================================================
' Class       : MeshBuffer
' Version     : v8.2.0  (VAO + Multi-Layout)
' Description : GPU VBO/VAO Manager with pluggable vertex layouts
' ============================================================

Public vbo As Long
Public IBO As Long
Public vao As Long
Public IsUploaded As Boolean

Private m_IndexCount As Long
Private m_Layout As VertexLayoutType   ' NEW: stores which layout was used

Public Sub LoadMesh(ByRef mesh As MeshData, Optional ByVal layout As VertexLayoutType = Layout_Position)
    If mesh Is Nothing Then Exit Sub
    m_IndexCount = mesh.indexCount
    m_Layout = layout
End Sub

Public Sub Upload(ByRef mesh As MeshData, Optional ByVal layout As VertexLayoutType = Layout_Position)
    If mesh.VertexCount = 0 Then Exit Sub
    
    m_IndexCount = mesh.indexCount
    m_Layout = layout
    
    ' Generate objects once
    If vao = 0 Then GL.glGenVertexArrays 1, vao
    If vbo = 0 Then GL.glGenBuffers 1, vbo
    If IBO = 0 Then GL.glGenBuffers 1, IBO
    
    GL.glBindVertexArray vao
    
    ' VBO - vertex data (size depends on layout)
    Dim vertexByteSize As Long
    vertexByteSize = mesh.VertexCount * GetByteStride(layout)
    
    GL.glBindBuffer GL_ARRAY_BUFFER, vbo
    GL.glBufferData GL_ARRAY_BUFFER, vertexByteSize, VarPtr(mesh.vertices(0)), GL_STATIC_DRAW
    
    ' IBO
    GL.glBindBuffer GL_ELEMENT_ARRAY_BUFFER, IBO
    GL.glBufferData GL_ELEMENT_ARRAY_BUFFER, mesh.indexCount * 4, VarPtr(mesh.indices(0)), GL_STATIC_DRAW
    
    ' Apply correct attribute layout (this is the magic)
    Call ApplyLayout(layout)
    
    GL.glBindVertexArray 0
    IsUploaded = True
End Sub

Public Sub UpdateBuffer(ByRef mesh As MeshData)
    If Not IsUploaded Then
        Call Upload(mesh, m_Layout)
        Exit Sub
    End If
    
    GL.glBindBuffer GL_ARRAY_BUFFER, vbo
    GL.glBufferData GL_ARRAY_BUFFER, mesh.VertexCount * GetByteStride(m_Layout), VarPtr(mesh.vertices(0)), GL_STATIC_DRAW
    
    GL.glBindBuffer GL_ELEMENT_ARRAY_BUFFER, IBO
    GL.glBufferData GL_ELEMENT_ARRAY_BUFFER, mesh.indexCount * 4, VarPtr(mesh.indices(0)), GL_STATIC_DRAW
End Sub

Public Sub Bind()
    GL.glBindVertexArray vao
End Sub

Public Sub Draw()
    If Not IsUploaded Then Exit Sub
    GL.glDrawElements GL_TRIANGLES, m_IndexCount, GL_UNSIGNED_INT, 0&
End Sub

Public Sub Destroy()
    If vbo <> 0 Then GL.glDeleteBuffers 1, vbo: vbo = 0
    If IBO <> 0 Then GL.glDeleteBuffers 1, IBO: IBO = 0
    If vao <> 0 Then GL.glDeleteVertexArrays 1, vao: vao = 0
    IsUploaded = False
End Sub

' Optional helper if you want to expose the layout
Public Property Get CurrentLayout() As VertexLayoutType
    CurrentLayout = m_Layout
End Property

Public Sub UploadInstanced(ByRef InstanceData() As Single, ByVal instanceCount As Long)
    Dim instanceVBO As Long
    GL.glGenBuffers 1, instanceVBO
    GL.glBindBuffer GL_ARRAY_BUFFER, instanceVBO
    GL.glBufferData GL_ARRAY_BUFFER, instanceCount * 16 * 4, VarPtr(InstanceData(0)), GL_DYNAMIC_DRAW   ' e.g. 4x4 matrix = 16 floats
    
    GL.glBindVertexArray vao
    ' Assume attributes 3-6 for a full matrix (or pack into mat4)
    Dim i As Long
    For i = 0 To 3
        GL.glEnableVertexAttribArray 3 + i
        GL.glVertexAttribPointer 3 + i, 4, GL_FLOAT, 0, 64, i * 16   ' stride = 64 bytes (mat4)
        GL.glVertexAttribDivisor 3 + i, 1
    Next i
    GL.glBindVertexArray 0
End Sub

