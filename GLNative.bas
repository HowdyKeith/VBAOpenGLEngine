Option Explicit

' ============================================================
' Module     : GLNative
' Version    : v8.08
' Description: Low-level API Dispatcher
' ============================================================

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, _
    ByVal vtReturn As Integer, ByVal cActuals As Long, _
    ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

' --- DISPATCH ENGINE ---

Private Function Dispatch(ByVal fn As LongPtr, ByVal vtRet As Integer, _
    ByRef vt() As Integer, ByRef pv() As LongPtr) As Variant
    
    Dim r As Variant
    If fn = 0 Then Exit Function
    
    Dim hr As Long, cActuals As Long
    On Error Resume Next
    cActuals = UBound(vt) + 1
    On Error GoTo 0
    
    ' Call STDCALL (4)
    If cActuals = 0 Then
        hr = DispCallFunc(0, fn, 4, vtRet, 0, 0, 0, r)
    Else
        hr = DispCallFunc(0, fn, 4, vtRet, cActuals, vt(0), pv(0), r)
    End If
    
    Dispatch = r
End Function

' --- TYPE-SAFE HELPERS ---

Private Function CLngPtrSafe(ByVal v As Variant) As LongPtr
    If IsEmpty(v) Or IsNull(v) Then CLngPtrSafe = 0 Else CLngPtrSafe = CLngPtr(v)
End Function

Private Function CLngSafe(ByVal v As Variant) As Long
    If IsEmpty(v) Or IsNull(v) Then CLngSafe = 0 Else CLngSafe = CLng(v)
End Function

' --- PUBLIC CALL SIGNATURES ---

' Returns LongPtr (used for IDs/Pointers) - 0 Args
Public Function CallR0LP(ByVal fn As LongPtr) As LongPtr
    Dim vt() As Integer, pv() As LongPtr
    CallR0LP = CLngPtrSafe(Dispatch(fn, 20, vt, pv))
End Function

' Returns Long - 1 Int Arg (glCreateShader)
Public Function CallR1I_I(ByVal fn As LongPtr, ByVal a0 As Long) As Long
    Dim vt(0) As Integer, pv(0) As LongPtr
    vt(0) = 3: pv(0) = VarPtr(a0)
    CallR1I_I = CLngSafe(Dispatch(fn, 3, vt, pv))
End Function

' Returns Long - 2 Pointers (glGetUniformLocation)
Public Function CallR2PP_I(ByVal fn As LongPtr, ByVal a0 As LongPtr, ByVal a1 As LongPtr) As Long
    Dim vt(1) As Integer, pv(1) As LongPtr
    vt(0) = 20: pv(0) = VarPtr(a0)
    vt(1) = 20: pv(1) = VarPtr(a1)
    CallR2PP_I = CLngSafe(Dispatch(fn, 3, vt, pv))
End Function

' Sub - 1 Int Arg (glUseProgram, glCompileShader, etc.)
Public Sub Call1I(ByVal fn As LongPtr, ByVal a0 As Long)
    Dim vt(0) As Integer, pv(0) As LongPtr
    vt(0) = 3: pv(0) = VarPtr(a0)
    Dispatch fn, 0, vt, pv
End Sub

' Sub - 2 Int Args (glAttachShader)
Public Sub Call2II(ByVal fn As LongPtr, ByVal a0 As Long, ByVal a1 As Long)
    Dim vt(1) As Integer, pv(1) As LongPtr
    vt(0) = 3: pv(0) = VarPtr(a0)
    vt(1) = 3: pv(1) = VarPtr(a1)
    Dispatch fn, 0, vt, pv
End Sub

' Sub - 4 Args: 3 Int, 1 Pointer (glUniformMatrix4fv)
Public Sub Call4IIIP(ByVal fn As LongPtr, ByVal a0 As Long, ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As LongPtr)
    Dim vt(3) As Integer, pv(3) As LongPtr
    vt(0) = 3: pv(0) = VarPtr(a0)
    vt(1) = 3: pv(1) = VarPtr(a1)
    vt(2) = 3: pv(2) = VarPtr(a2)
    vt(3) = 20: pv(3) = a3 ' Address already passed as LongPtr
    Dispatch fn, 0, vt, pv
End Sub

' Sub - 1 Int, 1 Pointer (for glGenVertexArrays / glGenBuffers)
Public Sub Call2IP(ByVal fn As LongPtr, ByVal a0 As Long, ByVal a1 As LongPtr)
    Dim vt(1) As Integer, pv(1) As LongPtr
    vt(0) = 3: pv(0) = VarPtr(a0)
    vt(1) = 20: pv(1) = VarPtr(a1)
    Dispatch fn, 0, vt, pv
End Sub

' Sub - 4 Args (for glBufferData)
Public Sub Call4PIPI(ByVal fn As LongPtr, ByVal a0 As LongPtr, ByVal a1 As LongPtr, ByVal a2 As LongPtr, ByVal a3 As LongPtr)
    Dim vt(3) As Integer, pv(3) As LongPtr
    vt(0) = 20: pv(0) = VarPtr(a0)
    vt(1) = 20: pv(1) = VarPtr(a1)
    vt(2) = 20: pv(2) = VarPtr(a2)
    vt(3) = 20: pv(3) = VarPtr(a3)
    Dispatch fn, 0, vt, pv
End Sub

' Sub - 6 Args (for glVertexAttribPointer)
Public Sub Call6IIIIIP(ByVal fn As LongPtr, ByVal a0 As Long, ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As Long, ByVal a4 As Long, ByVal a5 As LongPtr)
    Dim vt(5) As Integer, pv(5) As LongPtr
    vt(0) = 3: pv(0) = VarPtr(a0)
    vt(1) = 3: pv(1) = VarPtr(a1)
    vt(2) = 3: pv(2) = VarPtr(a2)
    vt(3) = 3: pv(3) = VarPtr(a3)
    vt(4) = 3: pv(4) = VarPtr(a4)
    vt(5) = 20: pv(5) = VarPtr(a5)
    Dispatch fn, 0, vt, pv
End Sub



