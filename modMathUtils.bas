Option Explicit
' ============================================================
' modMathUtils.bas
' Shared Matrix + Vector math utilities (v7 FIXED)
' Works with GLMath.Mat4 (Public Type) and Vector3 class.
' ============================================================

' ============================================================
' MATRIX OPERATIONS (operate on GLMath.Mat4 UDT in-place)
' ============================================================

Public Sub IdentityMatrix(ByRef m As GLMath.Mat4)
    Dim i As Long
    For i = 0 To 15
        m.m(i) = 0
    Next i
    m.m(0) = 1
    m.m(5) = 1
    m.m(10) = 1
    m.m(15) = 1
End Sub

Public Sub TranslateMatrix(ByRef m As GLMath.Mat4, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    m.m(12) = m.m(0) * x + m.m(4) * y + m.m(8) * z + m.m(12)
    m.m(13) = m.m(1) * x + m.m(5) * y + m.m(9) * z + m.m(13)
    m.m(14) = m.m(2) * x + m.m(6) * y + m.m(10) * z + m.m(14)
    m.m(15) = m.m(3) * x + m.m(7) * y + m.m(11) * z + m.m(15)
End Sub

Public Sub ScaleMatrix(ByRef m As GLMath.Mat4, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    m.m(0) = m.m(0) * x
    m.m(1) = m.m(1) * x
    m.m(2) = m.m(2) * x
    m.m(3) = m.m(3) * x
    m.m(4) = m.m(4) * y
    m.m(5) = m.m(5) * y
    m.m(6) = m.m(6) * y
    m.m(7) = m.m(7) * y
    m.m(8) = m.m(8) * z
    m.m(9) = m.m(9) * z
    m.m(10) = m.m(10) * z
    m.m(11) = m.m(11) * z
End Sub

Public Sub RotateMatrixX(ByRef m As GLMath.Mat4, ByVal angleDeg As Single)
    Dim rad As Single
    rad = angleDeg * 3.14159265 / 180#
    Dim c As Single: c = Cos(rad)
    Dim s As Single: s = Sin(rad)
    Dim tmp As GLMath.Mat4
    tmp = m
    m.m(4) = tmp.m(4) * c + tmp.m(8) * s
    m.m(5) = tmp.m(5) * c + tmp.m(9) * s
    m.m(6) = tmp.m(6) * c + tmp.m(10) * s
    m.m(7) = tmp.m(7) * c + tmp.m(11) * s
    m.m(8) = tmp.m(4) * (-s) + tmp.m(8) * c
    m.m(9) = tmp.m(5) * (-s) + tmp.m(9) * c
    m.m(10) = tmp.m(6) * (-s) + tmp.m(10) * c
    m.m(11) = tmp.m(7) * (-s) + tmp.m(11) * c
End Sub

Public Sub RotateMatrixY(ByRef m As GLMath.Mat4, ByVal angleDeg As Single)
    Dim rad As Single
    rad = angleDeg * 3.14159265 / 180#
    Dim c As Single: c = Cos(rad)
    Dim s As Single: s = Sin(rad)
    Dim tmp As GLMath.Mat4
    tmp = m
    m.m(0) = tmp.m(0) * c + tmp.m(8) * (-s)
    m.m(1) = tmp.m(1) * c + tmp.m(9) * (-s)
    m.m(2) = tmp.m(2) * c + tmp.m(10) * (-s)
    m.m(3) = tmp.m(3) * c + tmp.m(11) * (-s)
    m.m(8) = tmp.m(0) * s + tmp.m(8) * c
    m.m(9) = tmp.m(1) * s + tmp.m(9) * c
    m.m(10) = tmp.m(2) * s + tmp.m(10) * c
    m.m(11) = tmp.m(3) * s + tmp.m(11) * c
End Sub

Public Sub RotateMatrixZ(ByRef m As GLMath.Mat4, ByVal angleDeg As Single)
    Dim rad As Single
    rad = angleDeg * 3.14159265 / 180#
    Dim c As Single: c = Cos(rad)
    Dim s As Single: s = Sin(rad)
    Dim tmp As GLMath.Mat4
    tmp = m
    m.m(0) = tmp.m(0) * c + tmp.m(4) * s
    m.m(1) = tmp.m(1) * c + tmp.m(5) * s
    m.m(2) = tmp.m(2) * c + tmp.m(6) * s
    m.m(3) = tmp.m(3) * c + tmp.m(7) * s
    m.m(4) = tmp.m(0) * (-s) + tmp.m(4) * c
    m.m(5) = tmp.m(1) * (-s) + tmp.m(5) * c
    m.m(6) = tmp.m(2) * (-s) + tmp.m(6) * c
    m.m(7) = tmp.m(3) * (-s) + tmp.m(7) * c
End Sub

Public Function MultiplyMatrix(ByRef a As GLMath.Mat4, ByRef b As GLMath.Mat4) As GLMath.Mat4
    Dim r As GLMath.Mat4
    Dim row As Long, col As Long, k As Long
    For row = 0 To 3
        For col = 0 To 3
            Dim s As Single: s = 0
            For k = 0 To 3
                s = s + a.m(row + k * 4) * b.m(k + col * 4)
            Next k
            r.m(row + col * 4) = s
        Next col
    Next row
    MultiplyMatrix = r
End Function

' ============================================================
' VECTOR TRANSFORM HELPERS
' ============================================================

Public Function TransformVector3(ByRef mat As GLMath.Mat4, ByRef v As Vector3) As Vector3
    Dim result As Vector3
    Set result = New Vector3
    result.x = mat.m(0) * v.x + mat.m(4) * v.y + mat.m(8) * v.z + mat.m(12)
    result.y = mat.m(1) * v.x + mat.m(5) * v.y + mat.m(9) * v.z + mat.m(13)
    result.z = mat.m(2) * v.x + mat.m(6) * v.y + mat.m(10) * v.z + mat.m(14)
    Set TransformVector3 = result
End Function

Public Function TransformNormal(ByRef mat As GLMath.Mat4, ByRef v As Vector3) As Vector3
    ' Normals transform by the inverse-transpose (for uniform scale, same as rotation)
    Dim result As Vector3
    Set result = New Vector3
    result.x = mat.m(0) * v.x + mat.m(4) * v.y + mat.m(8) * v.z
    result.y = mat.m(1) * v.x + mat.m(5) * v.y + mat.m(9) * v.z
    result.z = mat.m(2) * v.x + mat.m(6) * v.y + mat.m(10) * v.z
    ' re-normalize
    Dim len As Single
    len = Sqr(result.X * result.X + result.Y * result.Y + result.z * result.z)
    If len > 0 Then
        result.X = result.X / len
        result.Y = result.Y / len
        result.z = result.z / len
    End If
    Set TransformNormal = result
End Function

Public Function TransformDirection(ByRef mat As GLMath.Mat4, ByRef v As Vector3) As Vector3
    ' Direction vectors ignore translation
    Set TransformDirection = TransformNormal(mat, v)
End Function

' ============================================================
' VECTOR3 HELPERS (operate on Vector3 class objects)
' ============================================================

Public Function Vec3Add(ByRef a As Vector3, ByRef b As Vector3) As Vector3
    Dim r As Vector3: Set r = New Vector3
    r.x = a.x + b.x: r.y = a.y + b.y: r.z = a.z + b.z
    Set Vec3Add = r
End Function

Public Function Vec3Sub(ByRef a As Vector3, ByRef b As Vector3) As Vector3
    Dim r As Vector3: Set r = New Vector3
    r.x = a.x - b.x: r.y = a.y - b.y: r.z = a.z - b.z
    Set Vec3Sub = r
End Function

Public Function Vec3Dot(ByRef a As Vector3, ByRef b As Vector3) As Single
    Vec3Dot = a.x * b.x + a.y * b.y + a.z * b.z
End Function

Public Function Vec3Cross(ByRef a As Vector3, ByRef b As Vector3) As Vector3
    Dim r As Vector3: Set r = New Vector3
    r.x = a.y * b.z - a.z * b.y
    r.y = a.z * b.x - a.x * b.z
    r.z = a.x * b.y - a.y * b.x
    Set Vec3Cross = r
End Function

Public Function Vec3Length(ByRef v As Vector3) As Single
    Vec3Length = Sqr(v.x * v.x + v.y * v.y + v.z * v.z)
End Function

Public Sub Vec3Normalize(ByRef v As Vector3)
    Dim len As Single
    len = Vec3Length(v)
    If len = 0 Then Exit Sub
    v.X = v.X / len: v.Y = v.Y / len: v.z = v.z / len
End Sub

Public Function Vec3Normalized(ByRef v As Vector3) As Vector3
    Dim r As Vector3: Set r = New Vector3
    r.x = v.x: r.y = v.y: r.z = v.z
    Vec3Normalize r
    Set Vec3Normalized = r
End Function

' ============================================================
' ATAN2 (Full quadrant-correct implementation)
' ============================================================
Public Function Atan2(ByVal y As Single, ByVal x As Single) As Single
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 And y >= 0 Then
        Atan2 = Atn(y / x) + 3.14159265
    ElseIf x < 0 And y < 0 Then
        Atan2 = Atn(y / x) - 3.14159265
    ElseIf x = 0 And y > 0 Then
        Atan2 = 3.14159265 / 2
    ElseIf x = 0 And y < 0 Then
        Atan2 = -3.14159265 / 2
    Else
        Atan2 = 0
    End If
End Function

Public Function DegToRad(ByVal deg As Single) As Single
    DegToRad = deg * 3.14159265 / 180#
End Function

Public Function RadToDeg(ByVal rad As Single) As Single
    RadToDeg = rad * 180# / 3.14159265
End Function



