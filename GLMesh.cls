Option Explicit

' ============================================================
' Class      : GLMesh
' Version    : v8.17
' Description: CPU Mesh Instance + Render Bridge
' ============================================================

Public id As Long
Public name As String
Public Active As Boolean

' --- GEOMETRY REFERENCES ---
Public MeshDataRef As MeshData
Public MeshBufferRef As MeshBuffer
Public MaterialRef As material

' --- TRANSFORM ---
Public Position As Vector3
Public Rotation As Vector3
Public ScaleVec As Vector3

' ============================================================
' INIT / TERMINATE
' ============================================================

Private Sub Class_Initialize()
    id = -1
    name = "New Mesh"
    Active = True

    Set MeshDataRef = Nothing
    Set MeshBufferRef = Nothing
    Set MaterialRef = Nothing

    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3

    ScaleVec.SetValues 1, 1, 1
End Sub

' ============================================================
' RENDER ENTRY POINT
' ============================================================

Public Sub Draw(ByRef shader As Object)
    If Not Active Then Exit Sub
    If shader Is Nothing Then Exit Sub
    If MaterialRef Is Nothing Then Exit Sub

    ' 1. Apply Material
    MaterialRef.Apply shader

    ' 2. Transform
    GL.glPushMatrix
        GL.glTranslatef Position.x, Position.y, Position.z
        GL.glRotatef Rotation.x, 1, 0, 0
        GL.glRotatef Rotation.y, 0, 1, 0
        GL.glRotatef Rotation.z, 0, 0, 1
        GL.glScalef ScaleVec.x, ScaleVec.y, ScaleVec.z

        ' GPU Path (High Performance)
        If Not MeshBufferRef Is Nothing Then
            MeshBufferRef.Draw
        ' CPU Path (Fallback / Debug)
        ElseIf Not MeshDataRef Is Nothing Then
            DrawImmediate
        End If
    GL.glPopMatrix
End Sub

' ============================================================
' IMMEDIATE MODE RENDER (FALLBACK)
' ============================================================

Private Sub DrawImmediate()
    If MeshDataRef Is Nothing Then Exit Sub

    ' 1. Pull the arrays into local variables to avoid "Wrong number of arguments"
    Dim v() As Single
    Dim idxArr() As Long
    
    v = MeshDataRef.vertices
    idxArr = MeshDataRef.indices
    
    Dim i As Long, idx As Long, vBase As Long
    Dim count As Long
    count = MeshDataRef.indexCount

    ' 2. Render using the local arrays
    GL.glBegin GL_TRIANGLES
    For i = 0 To count - 1 Step 3
        
        ' Vertex 0
        idx = idxArr(i): vBase = idx * 3
        GL.glVertex3f v(vBase), v(vBase + 1), v(vBase + 2)
        
        ' Vertex 1
        idx = idxArr(i + 1): vBase = idx * 3
        GL.glVertex3f v(vBase), v(vBase + 1), v(vBase + 2)
        
        ' Vertex 2
        idx = idxArr(i + 2): vBase = idx * 3
        GL.glVertex3f v(vBase), v(vBase + 1), v(vBase + 2)
        
    Next i
    GL.glEnd
End Sub
' ============================================================
' HELPERS
' ============================================================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub

Public Sub BindMeshData(ByRef mesh As MeshData): Set MeshDataRef = mesh: End Sub
Public Sub BindMeshBuffer(ByRef buffer As MeshBuffer): Set MeshBufferRef = buffer: End Sub
Public Sub BindMaterial(ByRef mat As material): Set MaterialRef = mat: End Sub


