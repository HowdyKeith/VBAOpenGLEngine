Option Explicit

' ============================================================
' Sphere.cls
' Basic Geometric Primitive (VBA OpenGL ECS v7)
' =========================

' =========================
' IDENTITY & SETTINGS
' =========================
Public id As Long
Public name As String

' =========================
' TRANSFORM
' =========================
' Assumes you have a Vector3 class in your project
Public Position As Vector3
Public Radius As Single

' =========================
' INIT
' =========================
Private Sub Class_Initialize()
    id = -1
    name = "Sphere"
    
    Set Position = New Vector3
    Radius = 1#
End Sub

' ============================================================
' UTILITIES
' ============================================================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.x = x
    Position.y = y
    Position.z = z
End Sub

' ============================================================
' DEBUG
' ============================================================
Public Sub DebugPrint()
    Debug.Print "---- SPHERE ----"
    Debug.Print "Name: "; name
    Debug.Print "Radius: "; Radius
    Debug.Print "Pos: "; Position.x; Position.y; Position.z
End Sub


