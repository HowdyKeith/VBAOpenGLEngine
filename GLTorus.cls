Option Explicit

' ============================================================
' GLTorus.cls  v1.0  WEEK 2
' Procedural torus (donut) primitive.
' Equivalent to: glutSolidTorus / glutWireTorus
'
' Usage:
'   Dim t As New GLTorus
'   t.InnerRadius = 0.3 : t.OuterRadius = 1.0
'   t.Sides = 16 : t.Rings = 24
'   t.Draw
'
' C equivalent: glutSolidTorus(innerRadius, outerRadius, sides, rings)
' ============================================================

Public InnerRadius As Single   ' radius of the tube
Public OuterRadius As Single   ' distance from center of tube to center of torus
Public Sides       As Long     ' segments around the tube cross-section
Public Rings       As Long     ' segments around the torus ring

Public Active   As Boolean
Public Position As Vector3
Public Rotation As Vector3
Public ScaleVec As Vector3

Private Sub Class_Initialize()
    InnerRadius = 0.3
    OuterRadius = 1#
    Sides       = 16
    Rings       = 24
    Active      = True
    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3
    ScaleVec.SetValues 1, 1, 1
End Sub

Private Sub Class_Terminate()
    Set Position = Nothing: Set Rotation = Nothing: Set ScaleVec = Nothing
End Sub

Public Sub Draw()
    If Not Active Then Exit Sub
    GL.glPushMatrix
    GLPrimitive.ApplyTransform Position, Rotation, ScaleVec
    DrawTorusGeometry False
    GL.glPopMatrix
End Sub

Public Sub DrawWire()
    If Not Active Then Exit Sub
    GL.glPushMatrix
    GLPrimitive.ApplyTransform Position, Rotation, ScaleVec
    DrawTorusGeometry True
    GL.glPopMatrix
End Sub

' Static versions (modFreeGLUT shim calls these)
Public Sub DrawSolid(ByVal inner As Single, ByVal outer As Single, ByVal sides As Long, ByVal rings As Long)
    InnerRadius = inner: OuterRadius = outer: Me.Sides = sides: Me.Rings = rings
    DrawTorusGeometry False
End Sub

Public Sub DrawWireform(ByVal inner As Single, ByVal outer As Single, ByVal sides As Long, ByVal rings As Long)
    InnerRadius = inner: OuterRadius = outer: Me.Sides = sides: Me.Rings = rings
    DrawTorusGeometry True
End Sub

Private Sub DrawTorusGeometry(ByVal wire As Boolean)
    Dim i As Long, j As Long
    Dim ringAngle As Double, sideAngle As Double
    Dim dRing As Double, dSide As Double

    dRing = GLPrimitive.TWO_PI / Rings
    dSide = GLPrimitive.TWO_PI / Sides

    If wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE

    For i = 0 To Rings - 1
        GL.glBegin GL_TRIANGLE_STRIP

        For j = 0 To Sides
            Dim k As Long
            For k = 0 To 1   ' two vertices per side position to form the strip
                ringAngle = (i + k) * dRing
                sideAngle = j * dSide

                ' Position on the torus surface
                Dim cosR As Single, sinR As Single
                Dim cosS As Single, sinS As Single
                cosR = CSng(Cos(ringAngle))
                sinR = CSng(Sin(ringAngle))
                cosS = CSng(Cos(sideAngle))
                sinS = CSng(Sin(sideAngle))

                ' Centre of tube at this ring position
                Dim cx As Single, cy As Single, cz As Single
                cx = OuterRadius * cosR
                cy = OuterRadius * sinR
                cz = 0

                ' Vertex = tube centre + inner offset along normal
                Dim vx As Single, vy As Single, vz As Single
                vx = (OuterRadius + InnerRadius * cosS) * cosR
                vy = (OuterRadius + InnerRadius * cosS) * sinR
                vz = InnerRadius * sinS

                ' Normal = vertex - tube centre, normalised
                Dim nx As Single, ny As Single, nz As Single
                nx = vx - cx: ny = vy - cy: nz = vz - cz
                GLPrimitive.Normalise3 nx, ny, nz

                GL.glNormal3f nx, ny, nz
                GL.glVertex3f vx, vy, vz
            Next k
        Next j

        GL.glEnd
    Next i

    If wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
End Sub

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub
