Option Explicit

' ============================================================
' Class      : Vector3
' Version    : v8.08
' Description: 3D vector math primitive (Performance Optimized)
' ============================================================

Public x As Single
Public y As Single
Public z As Single

' --- CORE SETTERS ---

' Direct value assignment
Public Sub SetValues(ByVal px As Single, ByVal py As Single, ByVal pz As Single)
    x = px: y = py: z = pz
End Sub

' Fast copy from another vector
Public Sub CopyFrom(ByRef v As Vector3)
    x = v.x: y = v.y: z = v.z
End Sub

' --- IN-PLACE MATH (Engine Loop Standard) ---
' Use these inside your Draw/Culling loops to avoid New object overhead

Public Sub Add(ByRef v As Vector3)
    x = x + v.x: y = y + v.y: z = z + v.z
End Sub

Public Sub Subtract(ByRef v As Vector3)
    x = x - v.x: y = y - v.y: z = z - v.z
End Sub

Public Sub ScaleInPlace(ByVal s As Single)
    x = x * s: y = y * s: z = z * s
End Sub

' --- FUNCTIONAL MATH (Returns New Object) ---
' Use these for one-off calculations where you don't want to change the original

Public Function GetAdded(ByRef v As Vector3) As Vector3
    Set GetAdded = New Vector3
    GetAdded.SetValues x + v.x, y + v.y, z + v.z
End Function

Public Function GetSubtracted(ByRef v As Vector3) As Vector3
    Set GetSubtracted = New Vector3
    GetSubtracted.SetValues x - v.x, y - v.y, z - v.z
End Function

' --- VECTOR OPERATIONS ---

' Modifies Me and returns Me for chaining (e.g., v.Normalize.ScaleInPlace 5)
Public Function Normalize() As Vector3
    Dim m As Single: m = Sqr(x * x + y * y + z * z)
    If m > 0 Then
        x = x / m: y = y / m: z = z / m
    End If
    Set Normalize = Me
End Function

' Returns the Dot Product scalar
Public Function Dot(ByRef v As Vector3) As Single
    Dot = (x * v.x) + (y * v.y) + (z * v.z)
End Function

' Returns a New Vector3 representing the cross product
Public Function Cross(ByRef v As Vector3) As Vector3
    Dim res As New Vector3
    res.x = (y * v.z) - (z * v.y)
    res.y = (z * v.x) - (x * v.z)
    res.z = (x * v.y) - (y * v.x)
    Set Cross = res
End Function

' --- DEBUG ---

Public Sub DebugPrint(Optional label As String = "Vector3")
    Debug.Print label & ": [" & x & ", " & y & ", " & z & "]"
End Sub


