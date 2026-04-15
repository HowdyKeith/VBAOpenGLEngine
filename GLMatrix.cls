Option Explicit

' =========================================================
' Class      : GLMatrix
' Version    : v8.05
' Description: 4x4 Row-Major Matrix for v8 Render Engine.
'              Fixed-size array workaround for Object Modules.
' =========================================================

' Private backing array to ensure memory is contiguous for OpenGL
Private p_m(0 To 15) As Single

' =========================================================
' INITIALIZATION
' =========================================================

Private Sub Class_Initialize()
    Identity
End Sub

' =========================================================
' CORE MATH OPERATIONS
' =========================================================

' Resets the matrix to an Identity state
Public Sub Identity()
    Dim i As Long
    For i = 0 To 15: p_m(i) = 0: Next i
    p_m(0) = 1: p_m(5) = 1: p_m(10) = 1: p_m(15) = 1
End Sub

' Property accessors to mimic Public Array syntax: Matrix.m(i)
Public Property Get m(ByVal index As Long) As Single
    If index >= 0 And index <= 15 Then
        m = p_m(index)
    Else
        Err.Raise 9, "GLMatrix", "Matrix index out of bounds (0-15)"
    End If
End Property

Public Property Let m(ByVal index As Long, ByVal value As Single)
    If index >= 0 And index <= 15 Then
        p_m(index) = value
    Else
        Err.Raise 9, "GLMatrix", "Matrix index out of bounds (0-15)"
    End If
End Property

' =========================================================
' ENGINE BRIDGE
' =========================================================

' Returns the raw memory pointer to the first element.
' Required for glUniformMatrix4fv in the GL facade.
Public Function GetPtr() As LongPtr
    GetPtr = VarPtr(p_m(0))
End Function

' Returns a copy of the raw data array
Public Function GetData() As Single()
    GetData = p_m
End Function

' =========================================================
' DEBUGGING
' =========================================================

Public Sub DebugPrint(Optional label As String = "GLMatrix")
    Dim i As Long
    Debug.Print "--- " & label & " ---"
    For i = 0 To 3
        Debug.Print format(p_m(i * 4), "0.000"); Tab(10); _
                    format(p_m(i * 4 + 1), "0.000"); Tab(20); _
                    format(p_m(i * 4 + 2), "0.000"); Tab(30); _
                    format(p_m(i * 4 + 3), "0.000")
    Next i
End Sub

' =========================================================
' CLEANUP
' =========================================================

Private Sub Class_Terminate()
    ' Array is automatically reclaimed
End Sub




