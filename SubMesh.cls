Option Explicit

' ============================================================
' SubMesh.cls (ECS v8 CLEAN VERSION)
' ============================================================

' =========================
' GEOMETRY DATA
' =========================

Private m_Vertices() As Vector3
Private m_Normals() As Vector3
Private m_UVs() As Vector2
Private m_Indices() As Long

Private m_VertexCount As Long
Private m_IndexCount As Long

' =========================
' RANGE INFO
' =========================

Public baseVertex As Long
Public baseIndex As Long

' =========================
' MATERIAL LINK
' =========================

Public MaterialID As Long
Public textureKey As String

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    m_VertexCount = 0
    m_IndexCount = 0

    ReDim m_Vertices(0 To 0)
    ReDim m_Normals(0 To 0)
    ReDim m_UVs(0 To 0)
    ReDim m_Indices(0 To 0)

    baseVertex = 0
    baseIndex = 0

    MaterialID = -1
    textureKey = vbNullString

End Sub

' =========================
' ADD VERTEX
' =========================

Public Sub AddVertex(ByRef v As Vector3, ByRef n As Vector3, ByRef uv As Vector2)

    m_VertexCount = m_VertexCount + 1

    ReDim Preserve m_Vertices(0 To m_VertexCount - 1)
    ReDim Preserve m_Normals(0 To m_VertexCount - 1)
    ReDim Preserve m_UVs(0 To m_VertexCount - 1)

    Set m_Vertices(m_VertexCount - 1) = v
    Set m_Normals(m_VertexCount - 1) = n
    Set m_UVs(m_VertexCount - 1) = uv

End Sub

' =========================
' ADD INDEX
' =========================

Public Sub AddIndex(ByVal index As Long)

    m_IndexCount = m_IndexCount + 1
    ReDim Preserve m_Indices(0 To m_IndexCount - 1)

    m_Indices(m_IndexCount - 1) = index

End Sub

' =========================
' GETTERS
' =========================

Public Function VertexCount() As Long
    VertexCount = m_VertexCount
End Function

Public Function indexCount() As Long
    indexCount = m_IndexCount
End Function

Public Function GetVertex(ByVal i As Long) As Vector3
    Set GetVertex = m_Vertices(i)
End Function

Public Function GetNormal(ByVal i As Long) As Vector3
    Set GetNormal = m_Normals(i)
End Function

Public Function GetUV(ByVal i As Long) As Vector2
    Set GetUV = m_UVs(i)
End Function

Public Function GetIndex(ByVal i As Long) As Long
    GetIndex = m_Indices(i)
End Function

' =========================
' CLEAR
' =========================

Public Sub Clear()

    m_VertexCount = 0
    m_IndexCount = 0

    ReDim m_Vertices(0 To 0)
    ReDim m_Normals(0 To 0)
    ReDim m_UVs(0 To 0)
    ReDim m_Indices(0 To 0)

End Sub


