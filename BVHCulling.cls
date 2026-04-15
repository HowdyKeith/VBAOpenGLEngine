Option Explicit

' ============================================================
' Class      : BVHCulling
' Version    : v8.43
' Description: Frustum Culling and BVH Traversal System
' ============================================================

Private m_BVH As Object
Private m_Frustum As frustum

Public Enabled As Boolean

Private Sub Class_Initialize()
    Enabled = True
End Sub

Public Sub Bind(ByRef bvh_obj As Object, ByRef frustum_obj As frustum)
    Set m_BVH = bvh_obj
    Set m_Frustum = frustum_obj
End Sub

Public Sub CullAndSubmit(ByRef queue As RenderQueue)
    If Not Enabled Then Exit Sub
    If m_BVH Is Nothing Or m_Frustum Is Nothing Then Exit Sub
    TraverseNode 0, queue
End Sub

Private Sub TraverseNode(ByVal node As Long, ByRef queue As RenderQueue)
    ' 1. FRUSTUM VISIBILITY CHECK
    If Not m_Frustum.IsBoxVisible(m_BVH.GetMinX(node), m_BVH.GetMinY(node), m_BVH.GetMinZ(node), _
                                  m_BVH.GetMaxX(node), m_BVH.GetMaxY(node), m_BVH.GetMaxZ(node)) Then
        Exit Sub
    End If

    ' 2. LEAF CHECK (Submit to Render Queue)
    If m_BVH.isLeaf(node) Then
        Dim e As Long
        e = m_BVH.GetEntity(node)

        If ComponentStore.IsActive(e) Then
            queue.Submit _
                ComponentStore.MeshVBO(e), _
                ComponentStore.MeshIBO(e), _
                ComponentStore.MeshIndexCount(e), _
                ComponentStore.MaterialID(e), _
                ComponentStore.PositionX(e), _
                ComponentStore.PositionY(e), _
                ComponentStore.PositionZ(e)
        End If
        Exit Sub
    End If

    ' 3. INTERNAL NODE (Traverse Children)
    Dim leftChild As Long, rightChild As Long
    leftChild = m_BVH.GetLeft(node)
    rightChild = m_BVH.GetRight(node)

    If leftChild <> -1 Then TraverseNode leftChild, queue
    If rightChild <> -1 Then TraverseNode rightChild, queue
End Sub

Private Sub Class_Terminate()
    Set m_BVH = Nothing
    Set m_Frustum = Nothing
End Sub




