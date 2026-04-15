Option Explicit

' =========================================================
' BVHBuilder.v8.cls
' COMPATIBLE WITH REFIT-ONLY BVH SYSTEM
' (NO InsertLeaf, NO HIERARCHY CONSTRUCTION)
' =========================================================

Public MaxLeafSize As Long
Public splitAxis As Long

Private Sub Class_Initialize()
    MaxLeafSize = 2
    splitAxis = 0
End Sub

' =========================================================
' BUILD ENTRY (SIMPLIFIED)
' =========================================================
Public Function Build(ByRef system As BVHSystem, ByRef Entities As Collection) As Long

    Dim i As Long
    Dim leafNodes() As Long

    ReDim leafNodes(0 To Entities.count - 1)

    ' -------------------------------------------------
    ' We DO NOT insert into BVH anymore
    ' BVH is already 1:1 mapped to ECS entities
    ' -------------------------------------------------

    For i = 1 To Entities.count
        leafNodes(i - 1) = i - 1
    Next i

    ' -------------------------------------------------
    ' Just force BVH refit instead of rebuild
    ' -------------------------------------------------

    system.Refit

    Build = 0

End Function


