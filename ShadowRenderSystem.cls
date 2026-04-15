Option Explicit
' =========================================================
' ShadowRenderSystem.cls (ECS v8 - FIXED)
' CPU-side shadow depth map pass using DirectionalLight
' and octree spatial query.
' =========================================================
Public SceneGraphRef As SceneGraph
Public LightRef As DirectionalLight

' =========================================================
' SHADOW PASS ENTRY
' =========================================================
Public Sub RenderShadowPass(ByRef tree As octree)
    If LightRef Is Nothing Then Exit Sub
    If SceneGraphRef Is Nothing Then Exit Sub

    LightRef.ClearShadowMap

    ' Collect visible node IDs via octree traversal
    Dim visible() As Long
    Dim visCount As Long
    ReDim visible(0 To tree.NodeCount - 1)
    visCount = 0

    Call TraverseOctree(tree, tree.Root, visible, visCount)

    ' Write depth for each visible node
    Dim i As Long
    For i = 0 To visCount - 1
        Dim nodeID As Long
        nodeID = visible(i)
        ' Use world matrix Z (column-major [14]) as depth approximation
        Dim wm As Matrix4
        wm = SceneGraphRef.GetWorldMatrix(nodeID)
        Dim depth As Single
        depth = wm.m(14)
        ' Map nodeID into a 2D shadow texel (simple hash into 256x256 map)
        Dim tx As Long, ty As Long
        tx = (nodeID Mod 256)
        ty = ((nodeID \ 256) Mod 256)
        LightRef.WriteDepth tx, ty, depth
    Next i
End Sub

' =========================================================
' OCTREE TRAVERSAL (replaces undefined QueryOctreeVisible)
' =========================================================
Private Sub TraverseOctree(ByRef tree As octree, _
                            ByVal nodeIdx As Long, _
                            ByRef visible() As Long, _
                            ByRef visCount As Long)
    If nodeIdx < 0 Or nodeIdx >= tree.NodeCount Then Exit Sub

    Dim n As OctreeNode
    n = tree.Nodes(nodeIdx)

    If n.isLeaf Then
        ' Add all objects in this leaf
        Dim i As Long
        For i = 0 To n.ObjectCount - 1
            If visCount <= UBound(visible) Then
                visible(visCount) = n.Objects(i)
                visCount = visCount + 1
            End If
        Next i
    Else
        ' Recurse into children
        Dim c As Long
        For c = 0 To 7
            If n.Children(c) >= 0 Then
                TraverseOctree tree, n.Children(c), visible, visCount
            End If
        Next c
    End If
End Sub

' =========================================================
' CLEANUP
' =========================================================
Private Sub Class_Terminate()
    Set SceneGraphRef = Nothing
    Set LightRef = Nothing
End Sub


