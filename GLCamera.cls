Option Explicit

' ============================================================
' GLCamera.v8.12.cls  FIXED
' FIXES:
'   - RaycastDown Sub was declared but had no body and no End Sub
'     (syntax error that would prevent the entire project compiling)
'   - ApplyProjection and ApplyView stubs now contain real legacy GL calls
'     matching Camera.cls implementation
' ============================================================

Public Position As Vector3
Public Rotation As Vector3  ' X=Pitch, Y=Yaw

Public FOV         As Single
Public AspectRatio As Single
Public NearPlane   As Single
Public FarPlane    As Single

Private Sub Class_Initialize()
    Set Position = New Vector3
    Set Rotation = New Vector3
    FOV = 60
    AspectRatio = 800 / 600
    NearPlane = 0.1
    FarPlane = 500
End Sub

Private Sub Class_Terminate()
    Set Position = Nothing
    Set Rotation = Nothing
End Sub

' --- SYSTEM HOOK ---
Public Sub Update(ByRef inputMgr As Object, ByVal dt As Single)
    ' Logic handled by FPSController or CameraSystem
End Sub

' --- DIRECTION MATH ---
Public Function GetForward() As Vector3
    Dim v As New Vector3
    Dim pRad As Single, yRad As Single
    pRad = Rotation.x * (3.14159 / 180)
    yRad = Rotation.y * (3.14159 / 180)
    v.x = Sin(yRad) * Cos(pRad)
    v.y = -Sin(pRad)
    v.z = -Cos(yRad) * Cos(pRad)
    Set GetForward = v.Normalize()
End Function

Public Function GetRight() As Vector3
    Dim fwd As Vector3: Set fwd = Me.GetForward()
    Dim worldUp As New Vector3: worldUp.SetValues 0, 1, 0
    Set GetRight = fwd.Cross(worldUp).Normalize()
End Function

' --- PROJECTION (FIXED: was empty stub) ---
Public Sub ApplyProjection()
    GL.glMatrixMode GL.GL_PROJECTION
    GL.glLoadIdentity
    modGL.apiPerspective CDbl(FOV), CDbl(AspectRatio), CDbl(NearPlane), CDbl(FarPlane)
    GL.glMatrixMode GL.GL_MODELVIEW
End Sub

' --- VIEW (FIXED: was empty stub) ---
Public Sub ApplyView()
    GL.glMatrixMode GL.GL_MODELVIEW
    GL.glLoadIdentity
    GL.glRotatef -Rotation.x, 1, 0, 0
    GL.glRotatef -Rotation.y, 0, 1, 0
    GL.glTranslatef -Position.x, -Position.y, -Position.z
End Sub

' --- FIXED: RaycastDown had no End Sub - was a compile-breaking syntax error ---
Public Sub RaycastDown(ByRef Origin As Vector3, ByRef OutResult As IntersectionResult)
    ' Delegate to physics system if available, otherwise this is a no-op stub.
    ' FPSController calls physics.RaycastDown directly; this stub exists for
    ' CameraSystem compatibility.
    OutResult.Clear
End Sub
