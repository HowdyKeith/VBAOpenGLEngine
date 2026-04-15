Option Explicit

' ============================================================
' modTypes.bas  v1.1
' Project-Wide Shared Type Definitions
' ADDITIONS over v1.0:
'   - RayHitFlat: flat struct version safe for .bas (Raycast.cls)
'   - EventRecord: replaces undefined EventFIX in EventSystem.cls
' ============================================================

' Matrix4 (identical layout to GLMath.Mat4, usable in UDTs)
Public Type Matrix4
    m(0 To 15) As Single
End Type

' InstanceRecord / InstanceBuffer (SceneGraph.CollectRenderInstances)
Public Type InstanceRecord
    WorldMatrix As Matrix4
    meshID      As Long
End Type

Public Type InstanceBuffer
    data()  As InstanceRecord
    Count   As Long
End Type

' RayHitFlat - safe flat struct (no object refs - VBA .bas modules cannot
' store Object references inside a Public Type).
' Raycast.cls should use this, plus its own separate Vector3 fields for point/normal.
Public Type RayHitFlat
    hit       As Boolean
    distance  As Single
    entityID  As Long
    px        As Single
    py        As Single
    pz        As Single
    nx        As Single
    ny        As Single
    nz        As Single
End Type

' EventRecord - replaces "EventFIX" (was undefined) in EventSystem.cls
Public Type EventRecord
    name As String
    data As Variant
End Type

' Matrix4 <-> GLMath.Mat4 copy helpers
Public Sub CopyMat4ToMatrix4(ByRef src As GLMath.Mat4, ByRef dst As Matrix4)
    Dim i As Long
    For i = 0 To 15: dst.m(i) = src.m(i): Next i
End Sub

Public Sub CopyMatrix4ToMat4(ByRef src As Matrix4, ByRef dst As GLMath.Mat4)
    Dim i As Long
    For i = 0 To 15: dst.m(i) = src.m(i): Next i
End Sub
