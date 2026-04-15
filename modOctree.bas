Option Explicit
' =========================================================
' Module: modOctree
' Version: v1.1.0 (FIXED - Vector3 class removed from UDT)
' Role: Spatial partitioning (AABB octree)
' NOTE: VBA UDTs cannot contain object references.
'       Center stored as flat Singles cx/cy/cz.
' =========================================================
Public Type OctreeNode
    cx          As Single       ' Center X (was: Center As Vector3)
    cy          As Single       ' Center Y
    cz          As Single       ' Center Z
    HalfSize    As Single
    Children(0 To 7) As Long
    Objects()   As Long
    ObjectCount As Long
    isLeaf      As Boolean
End Type

Public Type octree
    Nodes()     As OctreeNode
    Root        As Long
    NodeCount   As Long
End Type

' =========================================================
' HELPERS
' =========================================================
Public Function OctreeContains(ByRef n As OctreeNode, _
                                ByVal x As Single, _
                                ByVal y As Single, _
                                ByVal z As Single) As Boolean
    OctreeContains = (Abs(x - n.cx) <= n.HalfSize And _
                      Abs(y - n.cy) <= n.HalfSize And _
                      Abs(z - n.cz) <= n.HalfSize)
End Function


