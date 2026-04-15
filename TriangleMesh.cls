Option Explicit

' ============================================================
' TriangleMesh.v8.4.cls  FIXED
' VAO/VBO GEOMETRY CONTAINER
' FIXES:
'   - GL_ARRAY_BUFFER, GL_STATIC_DRAW, GL_FLOAT, GL_TRIANGLES were unqualified
'     constants - worked only if GLConstants.bas was also defining them.
'     Now explicitly qualified as GL.GL_* to use the authoritative source.
'   - glEnableVertexAttribArray and glVertexAttribPointer were stub no-ops in
'     the old GL.bas - now work correctly with fixed GL.bas
'   - Destroy guards against double-free (VAO/VBO already 0)
' ============================================================

Private m_VAO         As Long
Private m_VBO         As Long
Private m_VertexCount As Long

Public Sub UploadTriangle()
    ' 3 vertices * 3 floats each = 9 Singles
    Dim verts(0 To 8) As Single
    verts(0) = -0.5: verts(1) = -0.5: verts(2) = 0#
    verts(3) =  0.5: verts(4) = -0.5: verts(5) = 0#
    verts(6) =  0#:  verts(7) =  0.5: verts(8) = 0#

    m_VertexCount = 3

    ' Generate VAO
    GL.glGenVertexArrays 1, m_VAO
    GL.glBindVertexArray m_VAO

    ' Generate VBO and upload data
    GL.glGenBuffers 1, m_VBO
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBO
    GL.glBufferData GL.GL_ARRAY_BUFFER, 9 * 4, VarPtr(verts(0)), GL.GL_STATIC_DRAW

    ' Vertex attribute 0: position (3 floats, no stride, offset 0)
    GL.glEnableVertexAttribArray 0
    GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, GL.GL_FALSE, 0, 0

    ' Unbind
    GL.glBindVertexArray 0
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, 0

    Debug.Print "[TriangleMesh] Uploaded triangle VAO=" & m_VAO & " VBO=" & m_VBO
End Sub

Public Sub Draw()
    If m_VAO = 0 Then
        Debug.Print "[TriangleMesh] Draw called but VAO=0 (not uploaded?)"
        Exit Sub
    End If
    GL.glBindVertexArray m_VAO
    GL.glDrawArrays GL.GL_TRIANGLES, 0, m_VertexCount
    GL.glBindVertexArray 0
End Sub

Public Sub Destroy()
    If m_VAO <> 0 Then
        GL.glDeleteVertexArrays 1, m_VAO
        m_VAO = 0
    End If
    If m_VBO <> 0 Then
        GL.glDeleteBuffers 1, m_VBO
        m_VBO = 0
    End If
    m_VertexCount = 0
End Sub
