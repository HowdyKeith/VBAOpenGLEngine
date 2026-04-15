Option Explicit

' ============================================================
' Module     : modGLWrapper
' Version    : v8.41
' ============================================================

#If Win64 Then
    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cArgs As Long, ByRef rgvt As Integer, ByRef rgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
#Else
    Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cArgs As Long, ByRef rgvt As Integer, ByRef rgpvarg As Long, ByRef pvargResult As Variant) As Long
#End If

Private Function glCall(ByVal pFunc As LongPtr, ByVal argCount As Long, ParamArray args()) As Variant
    If pFunc = 0 Then Exit Function
    Dim vTypes() As Integer, vPtrs() As LongPtr, i As Long, res As Variant
    If argCount > 0 Then
        ReDim vTypes(argCount - 1): ReDim vPtrs(argCount - 1)
        For i = 0 To argCount - 1: vTypes(i) = VarType(args(i)): vPtrs(i) = VarPtr(args(i)): Next i
    End If
    DispCallFunc 0, pFunc, 4, vbLong, argCount, vTypes(0), vPtrs(0), res
    glCall = res
End Function

' --- VAO / VBO ---
Public Sub wGenVertexArrays(ByVal n As Long, ByRef a As Long): glCall GLLoader.p_glGenVertexArrays, 2, n, VarPtr(a): End Sub
Public Sub wBindVertexArray(ByVal id As Long): glCall GLLoader.p_glBindVertexArray, 1, id: End Sub
Public Sub wDeleteVertexArrays(ByVal n As Long, ByRef a As Long): glCall GLLoader.p_glDeleteVertexArrays, 2, n, VarPtr(a): End Sub
Public Sub wGenBuffers(ByVal n As Long, ByRef b As Long): glCall GLLoader.p_glGenBuffers, 2, n, VarPtr(b): End Sub
Public Sub wBindBuffer(ByVal t As Long, ByVal id As Long): glCall GLLoader.p_glBindBuffer, 2, t, id: End Sub
Public Sub wBufferData(ByVal t As Long, ByVal s As LongPtr, ByVal p As LongPtr, ByVal u As Long): glCall GLLoader.p_glBufferData, 4, t, s, p, u: End Sub
Public Sub wDeleteBuffers(ByVal n As Long, ByRef b As Long): glCall GLLoader.p_glDeleteBuffers, 2, n, VarPtr(b): End Sub

' --- SHADERS ---
Public Function wCreateShader(ByVal t As Long) As Long: wCreateShader = CLng(glCall(GLLoader.p_glCreateShader, 1, t)): End Function
Public Function wCreateProgram() As Long: wCreateProgram = CLng(glCall(GLLoader.p_glCreateProgram, 0)): End Function
Public Sub wShaderSource(ByVal s As Long, ByVal src As String): glCall GLLoader.p_glShaderSource, 4, s, 1, VarPtr(src), 0: End Sub
Public Sub wCompileShader(ByVal s As Long): glCall GLLoader.p_glCompileShader, 1, s: End Sub
Public Sub wAttachShader(ByVal p As Long, ByVal s As Long): glCall GLLoader.p_glAttachShader, 2, p, s: End Sub
Public Sub wLinkProgram(ByVal p As Long): glCall GLLoader.p_glLinkProgram, 1, p: End Sub
Public Sub wUseProgram(ByVal p As Long): glCall GLLoader.p_glUseProgram, 1, p: End Sub
Public Sub wDeleteShader(ByVal s As Long): glCall GLLoader.p_glDeleteShader, 1, s: End Sub
Public Sub wDeleteProgram(ByVal p As Long): glCall GLLoader.p_glDeleteProgram, 1, p: End Sub

' --- DRAWING ---
Public Sub wEnableVertexAttribArray(ByVal i As Long): glCall GLLoader.p_glEnableVertexAttribArray, 1, i: End Sub
Public Sub wVertexAttribPointer(ByVal i As Long, ByVal s As Long, ByVal t As Long, ByVal n As Byte, ByVal st As Long, ByVal o As LongPtr)
    glCall GLLoader.p_glVertexAttribPointer, 6, i, s, t, CLng(n), st, o
End Sub
Public Sub wDrawArrays(ByVal m As Long, ByVal f As Long, ByVal c As Long): glCall GLLoader.p_glDrawArrays, 3, m, f, c: End Sub
Public Sub wDrawElements(ByVal m As Long, ByVal c As Long, ByVal t As Long, ByVal i As LongPtr): glCall GLLoader.p_glDrawElements, 4, m, c, t, i: End Sub

' --- UNIFORMS ---
Public Sub wUniform1i(ByVal loc As Long, ByVal v0 As Long)
    glCall GL.p_glUniform1i, 2, loc, v0
End Sub

Public Sub wUniform1f(ByVal loc As Long, ByVal v0 As Single)
    ' Force vbSingle (4 bytes) type for DispCallFunc
    If GL.p_glUniform1f = 0 Then Exit Sub
    
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr
    Dim res As Variant
    
    vTypes(0) = vbLong:   vPtrs(0) = VarPtr(loc)
    vTypes(1) = vbSingle: vPtrs(1) = VarPtr(v0)
    
    ' Call CC_STDCALL (4) with 2 arguments
    DispCallFunc 0, GL.p_glUniform1f, 4, vbLong, 2, vTypes(0), vPtrs(0), res
End Sub


