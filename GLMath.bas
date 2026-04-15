' =========================================================
' GLMath.bas v5.2
' Minimal 3D math for camera + projection
' Column-major matrices (OpenGL style)
' =========================================================

Option Explicit

Public Type Mat4
    m(0 To 15) As Single
End Type

Public Function Identity() As Mat4
    Dim r As Mat4
    r.m(0) = 1
    r.m(5) = 1
    r.m(10) = 1
    r.m(15) = 1
    Identity = r
End Function

Public Function Perspective(ByVal FOV As Single, ByVal aspect As Single, _
                             ByVal nearZ As Single, ByVal farZ As Single) As Mat4

    Dim r As Mat4
    Dim t As Single

    t = 1 / Tan(FOV * 0.5 * 3.14159265 / 180)

    r.m(0) = t / aspect
    r.m(5) = t
    r.m(10) = (farZ + nearZ) / (nearZ - farZ)
    r.m(11) = -1
    r.m(14) = (2 * farZ * nearZ) / (nearZ - farZ)

    Perspective = r

End Function

Public Function LookAt( _
    ByVal eyeX As Single, ByVal eyeY As Single, ByVal eyeZ As Single, _
    ByVal centerX As Single, ByVal centerY As Single, ByVal centerZ As Single, _
    ByVal upX As Single, ByVal upY As Single, ByVal upZ As Single) As Mat4

    Dim fx As Single, fy As Single, fz As Single
    Dim sx As Single, sy As Single, sz As Single
    Dim ux As Single, uy As Single, uz As Single

    Dim r As Mat4

    ' forward
    fx = centerX - eyeX
    fy = centerY - eyeY
    fz = centerZ - eyeZ

    ' normalize (simple)
    Dim fl As Single
    fl = Sqr(fx * fx + fy * fy + fz * fz)
    fx = fx / fl: fy = fy / fl: fz = fz / fl

    ' side = f × up
    sx = fy * upZ - fz * upY
    sy = fz * upX - fx * upZ
    sz = fx * upY - fy * upX

    Dim sl As Single
    sl = Sqr(sx * sx + sy * sy + sz * sz)
    sx = sx / sl: sy = sy / sl: sz = sz / sl

    ' up = s × f
    ux = sy * fz - sz * fy
    uy = sz * fx - sx * fz
    uz = sx * fy - sy * fx

    r.m(0) = sx
    r.m(1) = ux
    r.m(2) = -fx
    r.m(3) = 0

    r.m(4) = sy
    r.m(5) = uy
    r.m(6) = -fy
    r.m(7) = 0

    r.m(8) = sz
    r.m(9) = uz
    r.m(10) = -fz
    r.m(11) = 0

    r.m(12) = -(sx * eyeX + sy * eyeY + sz * eyeZ)
    r.m(13) = -(ux * eyeX + uy * eyeY + uz * eyeZ)
    r.m(14) = (fx * eyeX + fy * eyeY + fz * eyeZ)
    r.m(15) = 1

    LookAt = r

End Function





