Option Explicit

' ============================================================
' MeshCollider.cls (v8 ECS BUFFER COMPONENT - FINAL CLEAN)
' ============================================================
' - ECS-safe ownership
' - No public arrays
' - BVH-ready AABB structure
' - Optimized ray traversal
' - VBA compile-safe lifecycle
' ============================================================

' =========================
' IDENTITY (ECS CORE)
' =========================

Public id As Long
Public entityID As Long

Public Active As Boolean
Public IsTrigger As Boolean

' =========================
' SOURCE LINK
' =========================

Public meshID As Long

' =========================
' INTERNAL BUFFERS (PRIVATE ECS MEMORY)
' =========================

Private mVertices() As Single   ' flat XYZ (x,y,z)
Private mIndices() As Long

Private mVertexCount As Long
Private mIndexCount As Long

' =========================
' BOUNDING BOX (BVH READY)
' =========================

Public BoundsMin As Vector3
Public BoundsMax As Vector3

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    id = -1
    entityID = -1
    meshID = -1

    Active = True
    IsTrigger = False

    mVertexCount = 0
    mIndexCount = 0

    Set BoundsMin = New Vector3
    Set BoundsMax = New Vector3

End Sub

' ============================================================
' LOAD FROM MESH SYSTEM
' ============================================================

Public Sub LoadFromMesh(ByRef mesh As MeshData)

    Dim i As Long

    If mesh Is Nothing Then Exit Sub
    If Not mesh.IsValid Then Exit Sub

    ' -------------------------
    ' COPY VERTEX BUFFER
    ' -------------------------
    mVertexCount = mesh.VertexCount

    If mVertexCount > 0 Then
        ReDim mVertices(0 To (mVertexCount * 3) - 1)

        For i = 0 To UBound(mVertices)
            mVertices(i) = mesh.vertices(i)
        Next i
    End If

    ' -------------------------
    ' COPY INDEX BUFFER
    ' -------------------------
    mIndexCount = mesh.indexCount

    If mIndexCount > 0 Then
        ReDim mIndices(0 To mIndexCount - 1)

        For i = 0 To UBound(mIndices)
            mIndices(i) = mesh.indices(i)
        Next i
    End If

    ' -------------------------
    ' BUILD AABB
    ' -------------------------
    ComputeBounds

End Sub

' ============================================================
' BOUNDS COMPUTATION (FAST PASS)
' ============================================================

Public Sub ComputeBounds()

    Dim i As Long
    Dim x As Single, y As Single, z As Single

    If mVertexCount = 0 Then Exit Sub

    BoundsMin.x = 1E+20
    BoundsMin.y = 1E+20
    BoundsMin.z = 1E+20

    BoundsMax.x = -1E+20
    BoundsMax.y = -1E+20
    BoundsMax.z = -1E+20

    For i = 0 To UBound(mVertices) Step 3

        x = mVertices(i)
        y = mVertices(i + 1)
        z = mVertices(i + 2)

        If x < BoundsMin.x Then BoundsMin.x = x
        If y < BoundsMin.y Then BoundsMin.y = y
        If z < BoundsMin.z Then BoundsMin.z = z

        If x > BoundsMax.x Then BoundsMax.x = x
        If y > BoundsMax.y Then BoundsMax.y = y
        If z > BoundsMax.z Then BoundsMax.z = z

    Next i

End Sub

' ============================================================
' AABB TEST (BVH / GRID HOOK)
' ============================================================

Public Function IntersectsAABB(ByRef minV As Vector3, ByRef maxV As Vector3) As Boolean

    If BoundsMax.x < minV.x Then Exit Function
    If BoundsMin.x > maxV.x Then Exit Function

    If BoundsMax.y < minV.y Then Exit Function
    If BoundsMin.y > maxV.y Then Exit Function

    If BoundsMax.z < minV.z Then Exit Function
    If BoundsMin.z > maxV.z Then Exit Function

    IntersectsAABB = True

End Function

' ============================================================
' FAST RAY INTERSECTION (TRIANGLE BRUTE FORCE)
' ============================================================

Public Function RayIntersect(ByRef Origin As Vector3, _
                             ByRef dir As Vector3, _
                             ByRef hitDistance As Single) As Boolean

    Dim i As Long
    Dim bestDist As Single
    Dim hit As Boolean

    Dim v0 As Vector3, V1 As Vector3, V2 As Vector3

    bestDist = 1E+20
    hit = False

    Set v0 = New Vector3
    Set V1 = New Vector3
    Set V2 = New Vector3

    For i = 0 To mIndexCount - 1 Step 3

        GetVertexFast mIndices(i), v0
        GetVertexFast mIndices(i + 1), V1
        GetVertexFast mIndices(i + 2), V2

        Dim dist As Single

        If RayTriangleIntersect(Origin, dir, v0, V1, V2, dist) Then
            If dist < bestDist Then
                bestDist = dist
                hit = True
            End If
        End If

    Next i

    hitDistance = bestDist
    RayIntersect = hit

End Function

' ============================================================
' MÖLLER–TRUMBORE INTERSECTION
' ============================================================

Private Function RayTriangleIntersect(ByRef orig As Vector3, _
                                      ByRef dir As Vector3, _
                                      ByRef v0 As Vector3, _
                                      ByRef V1 As Vector3, _
                                      ByRef V2 As Vector3, _
                                      ByRef t As Single) As Boolean

    Dim EPS As Single
    EPS = 0.000001

    Dim edge1 As Vector3, edge2 As Vector3
    Set edge1 = SubVec(V1, v0)
    Set edge2 = SubVec(V2, v0)

    Dim pvec As Vector3
    Set pvec = Cross(dir, edge2)

    Dim det As Single
    det = Dot(edge1, pvec)

    If det > -EPS And det < EPS Then Exit Function

    Dim invDet As Single
    invDet = 1 / det

    Dim tvec As Vector3
    Set tvec = SubVec(orig, v0)

    Dim u As Single
    u = Dot(tvec, pvec) * invDet

    If u < 0 Or u > 1 Then Exit Function

    Dim qvec As Vector3
    Set qvec = Cross(tvec, edge1)

    Dim v As Single
    v = Dot(dir, qvec) * invDet

    If v < 0 Or u + v > 1 Then Exit Function

    t = Dot(edge2, qvec) * invDet

    If t > EPS Then
        RayTriangleIntersect = True
    End If

End Function

' ============================================================
' FAST VERTEX ACCESS (NO ALLOCATION PER CALL)
' ============================================================

Public Sub GetVertexFast(ByVal index As Long, ByRef v As Vector3)

    Dim i As Long
    i = index * 3

    v.x = mVertices(i)
    v.y = mVertices(i + 1)
    v.z = mVertices(i + 2)

End Sub

' ============================================================
' VECTOR HELPERS
' ============================================================

Private Function SubVec(ByRef a As Vector3, ByRef b As Vector3) As Vector3

    Dim v As Vector3
    Set v = New Vector3

    v.x = a.x - b.x
    v.y = a.y - b.y
    v.z = a.z - b.z

    Set SubVec = v

End Function

Private Function Cross(ByRef a As Vector3, ByRef b As Vector3) As Vector3

    Dim v As Vector3
    Set v = New Vector3

    v.x = a.y * b.z - a.z * b.y
    v.y = a.z * b.x - a.x * b.z
    v.z = a.x * b.y - a.y * b.x

    Set Cross = v

End Function

Private Function Dot(ByRef a As Vector3, ByRef b As Vector3) As Single
    Dot = a.x * b.x + a.y * b.y + a.z * b.z
End Function

' ============================================================
' LIFECYCLE (ECS v8 STANDARD)
' ============================================================

Public Sub Enable()
    Active = True
End Sub

Public Sub Disable()
    Active = False
End Sub

Public Sub Destroy()

    Erase mVertices
    Erase mIndices

    mVertexCount = 0
    mIndexCount = 0

    Active = False
    meshID = -1
    entityID = -1

End Sub

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- MeshCollider v8 ----"
    Debug.Print "Entity: "; entityID
    Debug.Print "Vertices: "; mVertexCount
    Debug.Print "Indices: "; mIndexCount
    Debug.Print "Trigger: "; IsTrigger

End Sub




