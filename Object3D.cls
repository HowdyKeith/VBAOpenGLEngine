Option Explicit
' ============================================================
' Class      : Object3D
' Version    : v8.1
' Description: Base 3D transform container. Lightweight
'              alternative to Model when you need transform +
'              bounds without submesh/material management.
'              Use Model.cls for full renderable assets.
' ============================================================

' =========================
' IDENTITY
' =========================
Public id     As Long
Public name   As String
Public Active As Boolean

' =========================
' TRANSFORM
' =========================
Public Position As Vector3
Public Rotation As Vector3
Public ScaleVec As Vector3

' =========================
' BOUNDS (WORLD SPACE)
' =========================
Public BoundsMin As Vector3
Public BoundsMax As Vector3

' =========================
' VISIBILITY
' =========================
Public visible As Boolean

' =========================
' INIT
' =========================
Private Sub Class_Initialize()
    id = -1
    name = ""
    Active = True
    visible = True

    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3
    ScaleVec.SetValues 1, 1, 1

    Set BoundsMin = New Vector3
    Set BoundsMax = New Vector3
End Sub

' =========================
' TRANSFORM HELPERS
' =========================
Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub

Public Sub SetRotation(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Rotation.SetValues x, y, z
End Sub

Public Sub SetScale(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    ScaleVec.SetValues x, y, z
End Sub

' =========================
' BOUNDS UPDATE
' =========================
Public Sub UpdateBounds(ByRef minV As Vector3, ByRef maxV As Vector3)
    BoundsMin.SetValues minV.x, minV.y, minV.z
    BoundsMax.SetValues maxV.x, maxV.y, maxV.z
End Sub

' =========================
' APPLY TRANSFORM TO GL STACK
' =========================
Public Sub PushTransform()
    GL.glPushMatrix
    GL.glTranslatef Position.x, Position.y, Position.z
    GL.glRotatef Rotation.x, 1, 0, 0
    GL.glRotatef Rotation.y, 0, 1, 0
    GL.glRotatef Rotation.z, 0, 0, 1
    GL.glScalef ScaleVec.x, ScaleVec.y, ScaleVec.z
End Sub

Public Sub PopTransform()
    GL.glPopMatrix
End Sub

' =========================
' ENABLE / DISABLE
' =========================
Public Sub Enable():  Active = True:  End Sub
Public Sub Disable(): Active = False: End Sub

Public Sub SetVisible(ByVal value As Boolean)
    visible = value
End Sub

' =========================
' CLEANUP
' =========================
Private Sub Class_Terminate()
    Set Position = Nothing
    Set Rotation = Nothing
    Set ScaleVec = Nothing
    Set BoundsMin = Nothing
    Set BoundsMax = Nothing
End Sub

' =========================
' DEBUG
' =========================
Public Sub DebugPrint()
    Debug.Print "---- OBJECT3D ----"
    Debug.Print "ID: "; id; " Name: "; name
    Debug.Print "Active: "; Active; " Visible: "; visible
    Debug.Print "Pos: "; Position.x; Position.y; Position.z
    Debug.Print "Rot: "; Rotation.x; Rotation.y; Rotation.z
    Debug.Print "Scale: "; ScaleVec.x; ScaleVec.y; ScaleVec.z
End Sub

