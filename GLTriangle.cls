Option Explicit

' ============================================================
' Class      : GLTriangle
' Version    : v8.12
' Role       : Geometry primitive with Normal Calculation
' ============================================================

Public v0 As Vector3
Public V1 As Vector3
Public V2 As Vector3
Public normal As Vector3

Private Sub Class_Initialize()
    Set v0 = New Vector3
    Set V1 = New Vector3
    Set V2 = New Vector3
    Set normal = New Vector3
End Sub

Public Sub SetTriangle(ByRef a As Vector3, ByRef b As Vector3, ByRef c As Vector3)
    v0.SetValues a.x, a.y, a.z
    V1.SetValues b.x, b.y, b.z
    V2.SetValues c.x, c.y, c.z
    ComputeNormal
End Sub

Public Sub ComputeNormal()
    Dim ux As Single, uy As Single, uz As Single
    Dim vx As Single, vy As Single, vz As Single

    ux = V1.x - v0.x: uy = V1.y - v0.y: uz = V1.z - v0.z
    vx = V2.x - v0.x: vy = V2.y - v0.y: vz = V2.z - v0.z

    normal.x = uy * vz - uz * vy
    normal.y = uz * vx - ux * vz
    normal.z = ux * vy - uy * vx

    Dim l As Single: l = Sqr(normal.x * normal.x + normal.y * normal.y + normal.z * normal.z)
    If l <> 0 Then
        normal.x = normal.x / l: normal.y = normal.y / l: normal.z = normal.z / l
    End If
End Sub


