Option Explicit

' =========================================================
' MeshSystem.v8.1.cls
' SAFE CPU MESH STORAGE + RENDER ENTRY POINT
' =========================================================

Private Type MeshData
    vertices() As Single
    indices() As Long
    vao As Long
    vbo As Long
    IBO As Long
    IsUploaded As Boolean
End Type

Private m_Meshes() As MeshData
Private m_Count As Long

Public Sub UploadToGPU(ByVal id As Long)
    If id < 0 Or id >= m_Count Then Exit Sub
    
    ' Using local reference for clarity
    Dim m As MeshData
    m = m_Meshes(id)
    If m.IsUploaded Then Exit Sub

    ' 1. Create and Bind VAO
    GL.glGenVertexArrays 1, m.vao
    GL.glBindVertexArray m.vao

    ' 2. Upload Vertices (VBO)
    GL.glGenBuffers 1, m.vbo
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m.vbo
    GL.glBufferData GL.GL_ARRAY_BUFFER, (UBound(m.vertices) + 1) * 4, VarPtr(m.vertices(0)), GL.GL_STATIC_DRAW

    ' 3. Upload Indices (IBO)
    GL.glGenBuffers 1, m.IBO
    GL.glBindBuffer GL.GL_ELEMENT_ARRAY_BUFFER, m.IBO
    GL.glBufferData GL.GL_ELEMENT_ARRAY_BUFFER, (UBound(m.indices) + 1) * 4, VarPtr(m.indices(0)), GL.GL_STATIC_DRAW

    ' 4. Define Layout (3 floats per vertex)
    GL.glEnableVertexAttribArray 0
    GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, 0, 12, 0

    ' 5. Unbind
    GL.glBindVertexArray 0
    
    m.IsUploaded = True
    m_Meshes(id) = m
    Debug.Print "Mesh " & id & " uploaded. VAO: " & m.vao
End Sub

Public Sub Render(ByVal id As Long)
    If id < 0 Or id >= m_Count Then Exit Sub
    
    Dim m As MeshData
    m = m_Meshes(id)
    
    If Not m.IsUploaded Then Exit Sub

    ' 1. Bind the VAO (State container)
    GL.glBindVertexArray m.vao
    
    ' 2. Execute the Draw Call
    ' GL_TRIANGLES, Number of Indices, GL_UNSIGNED_INT, Offset 0
    GL.glDrawElements GL.GL_TRIANGLES, UBound(m.indices) + 1, GL.GL_UNSIGNED_INT, 0
    
    ' 3. Unbind
    GL.glBindVertexArray 0
End Sub


