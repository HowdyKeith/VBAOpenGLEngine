'If you want next upgrades, the REAL engine leap is:

'?? dynamic AABB tree (real-time physics like Box2D/Unity)
'?? fat AABB expansion system (continuous collision safety)
'?? frustum culling using camera planes
'?? spatial hashing hybrid system
'?? GPU-driven culling pipeline
Option Explicit

' ============================================================
' AABBBox.cls
' Axis-Aligned Bounding Box Component (VBA ECS v7 Core)
' ============================================================

' =========================
' ENTITY OWNER
' =========================

Public entityID As Long

' =========================
' BOUNDS
' =========================

Public Min As Vector3
Public Max As Vector3

' =========================
' DERIVED VALUES
' =========================

Public Center As Vector3
Public Extents As Vector3

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    entityID = -1

    Set Min = New Vector3
    Set Max = New Vector3

    Set Center = New Vector3
    Set Extents = New Vector3

End Sub

' ============================================================
' SET FROM MIN/MAX
' ============================================================

Public Sub SetMinMax(ByVal minX As Single, ByVal minY As Single, ByVal minZ As Single, _
                     ByVal maxX As Single, ByVal maxY As Single, ByVal maxZ As Single)

    Min.x = minX
    Min.y = minY
    Min.z = minZ

    Max.x = maxX
    Max.y = maxY
    Max.z = maxZ

    Call Recalculate

End Sub

' ============================================================
' SET FROM CENTER + EXTENTS
' ============================================================

Public Sub SetCenterExtents(ByVal cx As Single, ByVal cy As Single, ByVal cz As Single, _
                            ByVal ex As Single, ByVal ey As Single, ByVal ez As Single)

    Center.x = cx
    Center.y = cy
    Center.z = cz

    Extents.x = ex
    Extents.y = ey
    Extents.z = ez

    Min.x = cx - ex
    Min.y = cy - ey
    Min.z = cz - ez

    Max.x = cx + ex
    Max.y = cy + ey
    Max.z = cz + ez

End Sub

' ============================================================
' RECOMPUTE DERIVED VALUES
' ============================================================

Public Sub Recalculate()

    Center.x = (Min.x + Max.x) * 0.5
    Center.y = (Min.y + Max.y) * 0.5
    Center.z = (Min.z + Max.z) * 0.5

    Extents.x = (Max.x - Min.x) * 0.5
    Extents.y = (Max.y - Min.y) * 0.5
    Extents.z = (Max.z - Min.z) * 0.5

End Sub

' ============================================================
' EXPAND TO INCLUDE POINT
' ============================================================

Public Sub Encapsulate(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    If x < Min.x Then Min.x = x
    If y < Min.y Then Min.y = y
    If z < Min.z Then Min.z = z

    If x > Max.x Then Max.x = x
    If y > Max.y Then Max.y = y
    If z > Max.z Then Max.z = z

    Call Recalculate

End Sub

' ============================================================
' EXPAND TO INCLUDE OTHER BOX
' ============================================================

Public Sub EncapsulateBox(ByRef box As AABBBox)

    If box Is Nothing Then Exit Sub

    If box.Min.x < Min.x Then Min.x = box.Min.x
    If box.Min.y < Min.y Then Min.y = box.Min.y
    If box.Min.z < Min.z Then Min.z = box.Min.z

    If box.Max.x > Max.x Then Max.x = box.Max.x
    If box.Max.y > Max.y Then Max.y = box.Max.y
    If box.Max.z > Max.z Then Max.z = box.Max.z

    Call Recalculate

End Sub

' ============================================================
' INTERSECTION TEST (AABB vs AABB)
' ============================================================

Public Function Intersects(ByRef box As AABBBox) As Boolean

    If box Is Nothing Then Exit Function

    If Max.x < box.Min.x Or Min.x > box.Max.x Then Exit Function
    If Max.y < box.Min.y Or Min.y > box.Max.y Then Exit Function
    If Max.z < box.Min.z Or Min.z > box.Max.z Then Exit Function

    Intersects = True

End Function

' ============================================================
' POINT CONTAINMENT
' ============================================================

Public Function Contains(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Boolean

    If x < Min.x Or x > Max.x Then Exit Function
    If y < Min.y Or y > Max.y Then Exit Function
    If z < Min.z Or z > Max.z Then Exit Function

    Contains = True

End Function

' ============================================================
' RAY INTERSECTION (SLAB METHOD)
' ============================================================

Public Function RayIntersect(ByRef Origin As Vector3, _
                             ByRef dir As Vector3, _
                             ByRef distance As Single) As Boolean

    Dim tmin As Single
    Dim tmax As Single
    Dim tymin As Single
    Dim tymax As Single
    Dim tzmin As Single
    Dim tzmax As Single

    Dim invX As Single
    Dim invY As Single
    Dim invZ As Single

    invX = 1# / dir.x
    invY = 1# / dir.y
    invZ = 1# / dir.z

    If invX >= 0 Then
        tmin = (Min.x - Origin.x) * invX
        tmax = (Max.x - Origin.x) * invX
    Else
        tmin = (Max.x - Origin.x) * invX
        tmax = (Min.x - Origin.x) * invX
    End If

    If invY >= 0 Then
        tymin = (Min.y - Origin.y) * invY
        tymax = (Max.y - Origin.y) * invY
    Else
        tymin = (Max.y - Origin.y) * invY
        tymax = (Min.y - Origin.y) * invY
    End If

    If (tmin > tymax) Or (tymin > tmax) Then Exit Function

    If tymin > tmin Then tmin = tymin
    If tymax < tmax Then tmax = tymax

    If invZ >= 0 Then
        tzmin = (Min.z - Origin.z) * invZ
        tzmax = (Max.z - Origin.z) * invZ
    Else
        tzmin = (Max.z - Origin.z) * invZ
        tzmax = (Min.z - Origin.z) * invZ
    End If

    If (tmin > tzmax) Or (tzmin > tmax) Then Exit Function

    If tzmin > tmin Then tmin = tzmin
    If tzmax < tmax Then tmax = tzmax

    distance = tmin

    RayIntersect = True

End Function

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- AABB BOX ----"
    Debug.Print "Entity: "; entityID
    Debug.Print "Min: "; Min.x; Min.y; Min.z
    Debug.Print "Max: "; Max.x; Max.y; Max.z
    Debug.Print "Center: "; Center.x; Center.y; Center.z
    Debug.Print "Extents: "; Extents.x; Extents.y; Extents.z

End Sub





