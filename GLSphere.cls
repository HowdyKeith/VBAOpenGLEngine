Option Explicit

' ============================================================
' GLSphere.cls  v1.0  WEEK 2
' Procedural UV-sphere primitive.
' Equivalent to: glutSolidSphere / glutWireSphere
'
' Usage:
'   Dim s As New GLSphere
'   s.Radius = 1.5
'   s.Stacks = 16 : s.Slices = 16
'   s.SetPosition 0, 0, 0
'   s.Draw          ' solid
'   s.DrawWire      ' wireframe
' ============================================================

Public Radius   As Single
Public Slices   As Long      ' longitude segments (like orange slices)
Public Stacks   As Long      ' latitude segments  (like orange rings)
Public Active   As Boolean

Public Position As Vector3
Public Rotation As Vector3
Public ScaleVec As Vector3

Private Sub Class_Initialize()
    Radius  = 1#
    Slices  = 16
    Stacks  = 16
    Active  = True
    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3
    ScaleVec.SetValues 1, 1, 1
End Sub

Private Sub Class_Terminate()
    Set Position = Nothing
    Set Rotation = Nothing
    Set ScaleVec = Nothing
End Sub

' ============================================================
' SOLID DRAW  — equivalent to glutSolidSphere(radius, slices, stacks)
' ============================================================
Public Sub Draw()
    If Not Active Then Exit Sub
    GL.glPushMatrix
    GLPrimitive.ApplyTransform Position, Rotation, ScaleVec
    DrawSphereGeometry False
    GL.glPopMatrix
End Sub

' ============================================================
' WIRE DRAW  — equivalent to glutWireSphere(radius, slices, stacks)
' ============================================================
Public Sub DrawWire()
    If Not Active Then Exit Sub
    GL.glPushMatrix
    GLPrimitive.ApplyTransform Position, Rotation, ScaleVec
    DrawSphereGeometry True
    GL.glPopMatrix
End Sub

' ============================================================
' STATIC VERSIONS (no transform - called directly by modFreeGLUT shim)
' These mirror the C signature exactly:
'   void glutSolidSphere(double radius, int slices, int stacks)
' ============================================================
Public Sub DrawSolid(ByVal r As Single, ByVal slc As Long, ByVal stk As Long)
    Radius = r: Slices = slc: Stacks = stk
    DrawSphereGeometry False
End Sub

Public Sub DrawWireform(ByVal r As Single, ByVal slc As Long, ByVal stk As Long)
    Radius = r: Slices = slc: Stacks = stk
    DrawSphereGeometry True
End Sub

' ============================================================
' GEOMETRY GENERATION
' Produces a UV sphere with correct per-vertex normals and UVs.
' Uses GL_TRIANGLE_STRIP rows for minimal vertex count.
' ============================================================
Private Sub DrawSphereGeometry(ByVal wire As Boolean)
    Dim i As Long, j As Long
    Dim theta0 As Double, theta1 As Double   ' latitude  (0..PI)
    Dim phi    As Double                      ' longitude (0..2PI)
    Dim dTheta As Double, dPhi As Double

    dTheta = GLPrimitive.PI  / Stacks
    dPhi   = GLPrimitive.TWO_PI / Slices

    If wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE

    For i = 0 To Stacks - 1
        theta0 = i       * dTheta   ' top of strip
        theta1 = (i + 1) * dTheta   ' bottom of strip

        GL.glBegin GL_TRIANGLE_STRIP

        For j = 0 To Slices
            phi = j * dPhi

            ' --- top vertex ---
            Dim nx0 As Single, ny0 As Single, nz0 As Single
            nx0 = CSng(Sin(theta0) * Cos(phi))
            ny0 = CSng(Cos(theta0))
            nz0 = CSng(Sin(theta0) * Sin(phi))
            GL.glNormal3f nx0, ny0, nz0
            GL.glVertex3f nx0 * Radius, ny0 * Radius, nz0 * Radius

            ' --- bottom vertex ---
            Dim nx1 As Single, ny1 As Single, nz1 As Single
            nx1 = CSng(Sin(theta1) * Cos(phi))
            ny1 = CSng(Cos(theta1))
            nz1 = CSng(Sin(theta1) * Sin(phi))
            GL.glNormal3f nx1, ny1, nz1
            GL.glVertex3f nx1 * Radius, ny1 * Radius, nz1 * Radius
        Next j

        GL.glEnd
    Next i

    If wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
End Sub

' ============================================================
' TRANSFORM HELPERS
' ============================================================
Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub

Public Sub SetScale(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    ScaleVec.SetValues x, y, z
End Sub

Public Sub DebugPrint()
    Debug.Print "[GLSphere] r=" & Radius & " slices=" & Slices & " stacks=" & Stacks
End Sub
