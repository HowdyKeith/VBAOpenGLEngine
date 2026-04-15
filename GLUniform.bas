Option Explicit

' =========================================================
' Module     : GLUniform
' Version    : v8.16
' Description: Bridge for uploading VBA data to GLSL Shaders.
'              Requires GL.bas v8.47+ and GLMatrix.cls v8.05+.
' =========================================================

''' <summary>
''' Uploads a 4x4 Matrix to a named uniform in the shader program.
''' </summary>
Public Sub SetMat4(ByVal program As Long, ByVal name As String, ByRef m As GLMatrix)
    Dim loc As Long
    Dim nameAnsi() As Byte
    
    ' 1. Convert VBA String to Null-Terminated ANSI (Required for C-APIs)
    nameAnsi = StrConv(name & vbNullChar, vbFromUnicode)
    
    ' 2. Fetch the location ID from the shader program
    loc = GL.glGetUniformLocation(program, VarPtr(nameAnsi(0)))

    ' 3. If the uniform exists, upload the 16 floats from the matrix pointer
    If loc <> -1 Then
        ' Uses the raw memory address provided by GLMatrix.GetPtr
        GL.glUniformMatrix4fv loc, 1, 0, m.GetPtr
    Else
        ' Optional: Debug.Print "Uniform '" & Name & "' not found in program " & program
    End If
End Sub

''' <summary>
''' Uploads an Integer or Texture Sampler ID to a named uniform.
''' </summary>
Public Sub SetInt(ByVal program As Long, ByVal name As String, ByVal value As Long)
    Dim loc As Long
    Dim nameAnsi() As Byte
    
    nameAnsi = StrConv(name & vbNullChar, vbFromUnicode)
    loc = GL.glGetUniformLocation(program, VarPtr(nameAnsi(0)))
    
    If loc <> -1 Then
        GL.glUniform1i loc, value
    End If
End Sub

''' <summary>
''' Uploads a Float to a named uniform.
''' </summary>
Public Sub SetFloat(ByVal program As Long, ByVal name As String, ByVal value As Single)
    Dim loc As Long
    Dim nameAnsi() As Byte
    
    nameAnsi = StrConv(name & vbNullChar, vbFromUnicode)
    loc = GL.glGetUniformLocation(program, VarPtr(nameAnsi(0)))
    
    If loc <> -1 Then
        GL.glUniform1f loc, value
    End If
End Sub


