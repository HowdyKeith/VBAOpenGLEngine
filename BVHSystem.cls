Option Explicit

' =========================================================
' BVHSystem.v8.cls
' DYNAMIC REFIT BVH (SoA arrays, no object store)
' =========================================================

Private m_minX() As Single
Private m_minY() As Single
Private m_minZ() As Single

Private m_maxX() As Single
Private m_maxY() As Single
Private m_maxZ() As Single

Private m_left() As Long
Private m_right() As Long
Private m_entity() As Long
Private m_isLeaf() As Boolean

Private m_Count As Long

Public Sub Build(ByVal entityCount As Long)

    Dim i As Long
    m_Count = entityCount

    ReDim m_minX(m_Count - 1)
    ReDim m_minY(m_Count - 1)
    ReDim m_minZ(m_Count - 1)

    ReDim m_maxX(m_Count - 1)
    ReDim m_maxY(m_Count - 1)
    ReDim m_maxZ(m_Count - 1)

    ReDim m_left(m_Count - 1)
    ReDim m_right(m_Count - 1)
    ReDim m_entity(m_Count - 1)
    ReDim m_isLeaf(m_Count - 1)

    For i = 0 To m_Count - 1
        m_entity(i) = i
        m_isLeaf(i) = True
        m_left(i) = -1
        m_right(i) = -1

        UpdateLeaf i
    Next i

End Sub

Private Sub UpdateLeaf(ByVal id As Long)

    m_minX(id) = ComponentStore.AabbMinX(id)
    m_minY(id) = ComponentStore.AabbMinY(id)
    m_minZ(id) = ComponentStore.AabbMinZ(id)

    m_maxX(id) = ComponentStore.AabbMaxX(id)
    m_maxY(id) = ComponentStore.AabbMaxY(id)
    m_maxZ(id) = ComponentStore.AabbMaxZ(id)

End Sub

Public Sub Refit()

    Dim i As Long

    For i = 0 To m_Count - 1
        If m_isLeaf(i) Then UpdateLeaf i
    Next i

End Sub

' =========================================================
' PUBLIC ACCESSORS (required by RenderPipeline, BVHCulling)
' =========================================================
Public Property Get count() As Long
    count = m_Count
End Property

Public Function isLeaf(ByVal node As Long) As Boolean
    If node >= 0 And node < m_Count Then isLeaf = m_isLeaf(node)
End Function

Public Function GetEntity(ByVal node As Long) As Long
    If node >= 0 And node < m_Count Then GetEntity = m_entity(node)
End Function

Public Function GetLeft(ByVal node As Long) As Long
    If node >= 0 And node < m_Count Then GetLeft = m_left(node)
    Else: GetLeft = -1
End Function

Public Function GetRight(ByVal node As Long) As Long
    If node >= 0 And node < m_Count Then GetRight = m_right(node)
    Else: GetRight = -1
End Function

Public Function GetMinX(ByVal node As Long) As Single
    If node >= 0 And node < m_Count Then GetMinX = m_minX(node)
End Function

Public Function GetMinY(ByVal node As Long) As Single
    If node >= 0 And node < m_Count Then GetMinY = m_minY(node)
End Function

Public Function GetMinZ(ByVal node As Long) As Single
    If node >= 0 And node < m_Count Then GetMinZ = m_minZ(node)
End Function

Public Function GetMaxX(ByVal node As Long) As Single
    If node >= 0 And node < m_Count Then GetMaxX = m_maxX(node)
End Function

Public Function GetMaxY(ByVal node As Long) As Single
    If node >= 0 And node < m_Count Then GetMaxY = m_maxY(node)
End Function

Public Function GetMaxZ(ByVal node As Long) As Single
    If node >= 0 And node < m_Count Then GetMaxZ = m_maxZ(node)
End Function

' RenderPipeline compatibility stubs
' VAO/indexCount are stored in ComponentStore, not BVH.
' These delegate through ComponentStore for pipeline compatibility.
Public Function IsVisible(ByVal node As Long) As Boolean
    If node >= 0 And node < m_Count Then
        IsVisible = ComponentStore.IsActive(m_entity(node))
    End If
End Function

Public Function GetVAO(ByVal node As Long) As Long
    If node >= 0 And node < m_Count Then
        GetVAO = ComponentStore.MeshVAO(m_entity(node))
    End If
End Function

Public Function GetIndexCount(ByVal node As Long) As Long
    If node >= 0 And node < m_Count Then
        GetIndexCount = ComponentStore.MeshIndexCount(m_entity(node))
    End If
End Function

