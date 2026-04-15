Option Explicit

' ============================================================
' SpatialGridSystem.cls (ECS v8 BROADPHASE)
' ============================================================

Private mCellSize As Single

Private Type cell
    Entities As Collection
End Type

Private mGrid As Object ' Scripting.Dictionary (late-bound)

' =========================
' INIT
' =========================

Public Sub Initialize(ByVal cellSize As Single)

    Set mGrid = CreateObject("Scripting.Dictionary")
    mCellSize = cellSize

End Sub

' =========================
' HASH KEY
' =========================

Private Function HashKey(ByVal x As Long, ByVal y As Long, ByVal z As Long) As String
    HashKey = x & "_" & y & "_" & z
End Function

' =========================
' CLEAR GRID
' =========================

Public Sub Clear()
    mGrid.RemoveAll
End Sub

' =========================
' INSERT ENTITY AABB
' =========================

Public Sub InsertEntity(ByVal entityID As Long, ByRef minV As Vector3, ByRef maxV As Vector3)

    Dim x As Long, y As Long, z As Long
    Dim ix As Long, iy As Long, iz As Long

    Dim key As String
    Dim col As Collection

    ix = Int(minV.x / mCellSize)
    iy = Int(minV.y / mCellSize)
    iz = Int(minV.z / mCellSize)

    Dim ix2 As Long, iy2 As Long, iz2 As Long

    ix2 = Int(maxV.x / mCellSize)
    iy2 = Int(maxV.y / mCellSize)
    iz2 = Int(maxV.z / mCellSize)

    For x = ix To ix2
        For y = iy To iy2
            For z = iz To iz2

                key = HashKey(x, y, z)

                If Not mGrid.Exists(key) Then
                    Set col = New Collection
                    mGrid.Add key, col
                Else
                    Set col = mGrid(key)
                End If

                col.Add entityID

            Next z
        Next y
    Next x

End Sub

' =========================
' QUERY NEARBY ENTITIES
' =========================

Public Sub QueryAABB(ByRef minV As Vector3, ByRef maxV As Vector3, ByRef result As Collection)

    Dim x As Long, y As Long, z As Long
    Dim ix As Long, iy As Long, iz As Long

    Dim key As String
    Dim col As Collection
    Dim item As Variant

    Set result = New Collection

    ix = Int(minV.x / mCellSize)
    iy = Int(minV.y / mCellSize)
    iz = Int(minV.z / mCellSize)

    Dim ix2 As Long, iy2 As Long, iz2 As Long

    ix2 = Int(maxV.x / mCellSize)
    iy2 = Int(maxV.y / mCellSize)
    iz2 = Int(maxV.z / mCellSize)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    For x = ix To ix2
        For y = iy To iy2
            For z = iz To iz2

                key = HashKey(x, y, z)

                If mGrid.Exists(key) Then

                    Set col = mGrid(key)

                    For Each item In col
                        If Not seen.Exists(item) Then
                            seen.Add item, True
                            result.Add item
                        End If
                    Next item

                End If

            Next z
        Next y
    Next x

End Sub


