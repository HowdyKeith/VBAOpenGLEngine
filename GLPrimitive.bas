Option Explicit

' ============================================================
' GLPrimitive.bas  v1.0  WEEK 2
' Shared math utilities for procedural GL primitive generation.
' Used by GLSphere, GLTorus, GLCylinder.
' ============================================================

Public Const PI     As Double = 3.14159265358979
Public Const TWO_PI As Double = 6.28318530717959
Public Const HALF_PI As Double = 1.5707963267949

' ============================================================
' COLOUR HELPERS (used by wire vs solid mode)
' ============================================================
Public Sub SetColor3f(ByVal r As Single, ByVal g As Single, ByVal b As Single)
    ' Route through modGL direct call (glColor3f is in opengl32 1.x)
    ' We declare it here to keep primitives self-contained
    modGL.apiSetColor3f r, g, b
End Sub

' ============================================================
' DEGREE / RADIAN CONVERSION
' ============================================================
Public Function DegToRad(ByVal deg As Double) As Double
    DegToRad = deg * PI / 180#
End Function

Public Function RadToDeg(ByVal rad As Double) As Double
    RadToDeg = rad * 180# / PI
End Function

' ============================================================
' NORMALISE A 3-COMPONENT VECTOR (in-place via 3 Singles)
' ============================================================
Public Sub Normalise3(ByRef x As Single, ByRef y As Single, ByRef z As Single)
    Dim len As Single
    len = Sqr(x * x + y * y + z * z)
    If len > 0.000001 Then
        x = x / len: y = y / len: z = z / len
    End If
End Sub

' ============================================================
' DRAW TRANSFORM STACK HELPERS
' (so primitives can position themselves like glutSolid* does)
' ============================================================
Public Sub ApplyTransform(ByRef pos As Vector3, ByRef rot As Vector3, ByRef scl As Vector3)
    GL.glTranslatef pos.x, pos.y, pos.z
    GL.glRotatef rot.x, 1, 0, 0
    GL.glRotatef rot.y, 0, 1, 0
    GL.glRotatef rot.z, 0, 0, 1
    GL.glScalef scl.x, scl.y, scl.z
End Sub

' ============================================================
' POLYGON MODE HELPERS
' ============================================================
Public Sub BeginSolid()
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
End Sub

Public Sub BeginWire()
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE
End Sub

Public Sub EndPrimitive()
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL  ' restore
End Sub
