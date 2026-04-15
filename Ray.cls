Option Explicit

' ============================================================
' Ray.cls
' Mathematical Ray Primitive (VBA ECS v7 Engine Core)
' ============================================================

' =========================
' CORE DATA
' =========================

Public Origin As Vector3
Public direction As Vector3

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    Set Origin = New Vector3
    Set direction = New Vector3

    direction.x = 0
    direction.y = 0
    direction.z = 1

End Sub

' =========================
' SET RAY
' =========================

Public Sub Set(ByVal ox As Single, ByVal oy As Single, ByVal oz As Single, _
               ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)

    Origin.x = ox
    Origin.y = oy
    Origin.z = oz

    direction.x = dx
    direction.y = dy
    direction.z = dz

    Call Normalize(direction)

End Sub

' =========================
' NORMALIZE DIRECTION
' =========================

Public Sub NormalizeDirection()

    Call Normalize(direction)

End Sub

Private Sub Normalize(ByRef v As Vector3)

    Dim len As Single

    len = Sqr(v.X * v.X + v.Y * v.Y + v.Z * v.Z)

    If len = 0 Then Exit Sub

    v.X = v.X / len
    v.Y = v.Y / len
    v.Z = v.Z / len

End Sub

' =========================
' POINT ALONG RAY
' =========================

Public Function GetPoint(ByVal t As Single) As Vector3

    Dim p As Vector3
    Set p = New Vector3

    p.x = Origin.x + direction.x * t
    p.y = Origin.y + direction.y * t
    p.z = Origin.z + direction.z * t

    Set GetPoint = p

End Function

' =========================
' TRANSFORM RAY
' =========================

Public Sub Transform(ByRef mat As Matrix4)

    ' Transforms ray into another space (useful for local-space collision tests)

    Set Origin = modMathUtils.TransformVector3(mat, Origin)
    Set direction = modMathUtils.TransformDirection(mat, direction)

    Call Normalize(direction)

End Sub

' =========================
' COPY
' =========================

Public Function Clone() As Ray

    Dim r As Ray
    Set r = New Ray

    r.Origin.x = Origin.x
    r.Origin.y = Origin.y
    r.Origin.z = Origin.z

    r.direction.x = direction.x
    r.direction.y = direction.y
    r.direction.z = direction.z

    Set Clone = r

End Function

' =========================
' DEBUG
' =========================

Public Sub DebugPrint()

    Debug.Print "---- RAY ----"
    Debug.Print "Origin: "; Origin.x; Origin.y; Origin.z
    Debug.Print "Dir:    "; direction.x; direction.y; direction.z

End Sub





