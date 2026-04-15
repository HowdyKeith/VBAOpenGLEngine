Option Explicit
' ============================================================
' Class      : CameraSystem
' Version    : v8.1
' Description: ECS Camera System - manages active camera,
'              applies projection and view matrices each frame.
'              Works with GLCamera.cls and GL.bas facade.
' ============================================================

' =========================
' ACTIVE CAMERA REFERENCE
' =========================
Public ActiveCamera As GLCamera

' =========================
' PROJECTION SETTINGS
' =========================
Public FOV         As Single
Public NearPlane   As Single
Public FarPlane    As Single
Public AspectRatio As Single

' =========================
' INIT
' =========================
Private Sub Class_Initialize()
    Set ActiveCamera = New GLCamera
    FOV = 60#
    NearPlane = 0.1
    FarPlane = 1000#
    AspectRatio = 1.7777   ' 16:9 default
End Sub

' ============================================================
' BIND CAMERA (SWAP ACTIVE)
' ============================================================
Public Sub SetCamera(ByRef cam As GLCamera)
    If cam Is Nothing Then Exit Sub
    Set ActiveCamera = cam
End Sub

' ============================================================
' APPLY PROJECTION MATRIX
' ============================================================
Public Sub ApplyProjection()
    If ActiveCamera Is Nothing Then Exit Sub
    GL.glMatrixMode GL.GL_PROJECTION
    GL.glLoadIdentity
    GL.gluPerspective FOV, AspectRatio, NearPlane, FarPlane
    GL.glMatrixMode GL.GL_MODELVIEW
End Sub

' ============================================================
' APPLY VIEW MATRIX  (Rotation then Translation, inverted for camera space)
' ============================================================
Public Sub ApplyView()
    If ActiveCamera Is Nothing Then Exit Sub
    GL.glMatrixMode GL.GL_MODELVIEW
    GL.glLoadIdentity
    ' Pitch (X), Yaw (Y) - invert for camera space
    GL.glRotatef -ActiveCamera.Rotation.x, 1, 0, 0
    GL.glRotatef -ActiveCamera.Rotation.y, 0, 1, 0
    GL.glTranslatef -ActiveCamera.Position.x, _
                    -ActiveCamera.Position.y, _
                    -ActiveCamera.Position.z
End Sub

' ============================================================
' FULL FRAME UPDATE (call once per frame before rendering)
' ============================================================
Public Sub Update(ByRef inputMgr As Object, ByVal dt As Single)
    If ActiveCamera Is Nothing Then Exit Sub
    ' Delegate per-frame logic to the camera object itself
    On Error Resume Next
    ActiveCamera.Update inputMgr, dt
    On Error GoTo 0
    Call ApplyProjection
    Call ApplyView
End Sub

' ============================================================
' HELPERS
' ============================================================
Public Function GetPosition() As Vector3
    If ActiveCamera Is Nothing Then
        Set GetPosition = New Vector3
    Else
        Set GetPosition = ActiveCamera.Position
    End If
End Function

Public Function GetForward() As Vector3
    If ActiveCamera Is Nothing Then
        Set GetForward = New Vector3
    Else
        Set GetForward = ActiveCamera.GetForward()
    End If
End Function

' ============================================================
' CLEANUP
' ============================================================
Private Sub Class_Terminate()
    Set ActiveCamera = Nothing
End Sub


