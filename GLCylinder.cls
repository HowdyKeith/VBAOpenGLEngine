Option Explicit

' ============================================================
' GLCylinder.cls  v1.0  WEEK 2
' Procedural cylinder / cone primitive.
' Equivalent to: glutSolidCylinder / glutWireCylinder
'                glutSolidCone     / glutWireCone
'
' Usage:
'   Dim c As New GLCylinder
'   c.BaseRadius = 0.5 : c.TopRadius = 0.5 : c.Height = 2.0
'   c.Slices = 16 : c.Stacks = 4
'   c.Draw
'
' Set TopRadius = 0 for a cone (glutSolidCone equivalent).
' ============================================================

Public BaseRadius As Single
Public TopRadius  As Single    ' 0 = cone, BaseRadius = cylinder
Public Height     As Single
Public Slices     As Long      ' segments around circumference
Public Stacks     As Long      ' segments along height

Public Active   As Boolean
Public Position As Vector3
Public Rotation As Vector3
Public ScaleVec As Vector3

Private Sub Class_Initialize()
    BaseRadius = 0.5
    TopRadius  = 0.5
    Height     = 2#
    Slices     = 16
    Stacks     = 4
    Active     = True
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
    DrawCylinderGeometry False
    GL.glPopMatrix
End Sub

Public Sub DrawWire()
    If Not Active Then Exit Sub
    GL.glPushMatrix
    GLPrimitive.ApplyTransform Position, Rotation, ScaleVec
    DrawCylinderGeometry True
    GL.glPopMatrix
End Sub

' Static versions for modFreeGLUT shim
Public Sub DrawSolid(ByVal baseR As Single, ByVal topR As Single, ByVal h As Single, ByVal slc As Long, ByVal stk As Long)
    BaseRadius = baseR: TopRadius = topR: Height = h: Slices = slc: Stacks = stk
    DrawCylinderGeometry False
End Sub

Public Sub DrawWireform(ByVal baseR As Single, ByVal topR As Single, ByVal h As Single, ByVal slc As Long, ByVal stk As Long)
    BaseRadius = baseR: TopRadius = topR: Height = h: Slices = slc: Stacks = stk
    DrawCylinderGeometry True
End Sub

Private Sub DrawCylinderGeometry(ByVal wire As Boolean)
    Dim i As Long, j As Long
    Dim dPhi   As Double: dPhi   = GLPrimitive.TWO_PI / Slices
    Dim dStack As Single: dStack = Height / Stacks

    If wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE

    ' --- SIDE WALLS ---
    ' The slant normal for a truncated cone:
    ' nLen = sqrt((BaseRadius-TopRadius)^2 + Height^2) -- the slant height
    Dim slantLen As Single
    slantLen = Sqr((BaseRadius - TopRadius) ^ 2 + Height * Height)
    Dim nSide As Single   ' normal Y component (pointing outward along slant)
    Dim nOut  As Single   ' normal radial component
    If slantLen > 0 Then
        nSide = (BaseRadius - TopRadius) / slantLen
        nOut  = Height / slantLen
    Else
        nSide = 0: nOut = 1
    End If

    For i = 0 To Stacks - 1
        Dim y0 As Single: y0 = i       * dStack
        Dim y1 As Single: y1 = (i + 1) * dStack
        Dim r0 As Single: r0 = BaseRadius + (TopRadius - BaseRadius) * (y0 / Height)
        Dim r1 As Single: r1 = BaseRadius + (TopRadius - BaseRadius) * (y1 / Height)

        GL.glBegin GL_TRIANGLE_STRIP
        For j = 0 To Slices
            Dim phi As Double: phi = j * dPhi
            Dim cosP As Single: cosP = CSng(Cos(phi))
            Dim sinP As Single: sinP = CSng(Sin(phi))

            ' normal (same for both vertices in this strip row)
            GL.glNormal3f nOut * cosP, nSide, nOut * sinP
            GL.glVertex3f r1 * cosP, y1, r1 * sinP
            GL.glNormal3f nOut * cosP, nSide, nOut * sinP
            GL.glVertex3f r0 * cosP, y0, r0 * sinP
        Next j
        GL.glEnd
    Next i

    ' --- BASE CAP (if BaseRadius > 0) ---
    If BaseRadius > 0.001 Then
        GL.glBegin GL_TRIANGLE_FAN
        GL.glNormal3f 0, -1, 0
        GL.glVertex3f 0, 0, 0
        For j = Slices To 0 Step -1   ' reversed winding for outward-facing normal
            Dim phiB As Double: phiB = j * dPhi
            GL.glVertex3f BaseRadius * CSng(Cos(phiB)), 0, BaseRadius * CSng(Sin(phiB))
        Next j
        GL.glEnd
    End If

    ' --- TOP CAP (if TopRadius > 0) ---
    If TopRadius > 0.001 Then
        GL.glBegin GL_TRIANGLE_FAN
        GL.glNormal3f 0, 1, 0
        GL.glVertex3f 0, Height, 0
        For j = 0 To Slices
            Dim phiT As Double: phiT = j * dPhi
            GL.glVertex3f TopRadius * CSng(Cos(phiT)), Height, TopRadius * CSng(Sin(phiT))
        Next j
        GL.glEnd
    End If

    If wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
End Sub

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub
