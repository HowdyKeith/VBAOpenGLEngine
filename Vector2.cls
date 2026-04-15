Option Explicit

' ============================================================
' Class      : Vector2
' Version    : v1.0
' Description: 2D vector for UV coordinates.
'              Was used by SubMesh.cls but never defined.
' ============================================================

Public x As Single
Public y As Single

Public Sub SetValues(ByVal px As Single, ByVal py As Single)
    x = px: y = py
End Sub

Public Sub CopyFrom(ByRef v As Vector2)
    x = v.x: y = v.y
End Sub

Public Function Clone() As Vector2
    Dim r As New Vector2
    r.x = x: r.y = y
    Set Clone = r
End Function

Public Sub DebugPrint(Optional ByVal label As String = "Vector2")
    Debug.Print label & ": [" & x & ", " & y & "]"
End Sub
