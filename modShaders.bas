Option Explicit

' ============================================================
' modShaders.bas - v6.4 FIXED
' FIXES:
'   - Called modGL.glCreateShader / modGL.glShaderSource etc which didn't exist in modGL
'     All calls now route to GL.bas (the authoritative facade)
'   - GL_VERTEX_SHADER / GL_FRAGMENT_SHADER now referenced from GL module
'   - glShaderSource previously passed raw StrPtr + count args matching old stub;
'     now uses GL.glShaderSource which handles ANSI conversion internally
' ============================================================

Public Function CompileShader(ByVal src As String, ByVal shaderType As Long) As Long
    Dim shader As Long

    shader = GL.glCreateShader(shaderType)
    If shader = 0 Then
        Debug.Print "[modShaders] glCreateShader returned 0 for type " & shaderType
        Exit Function
    End If

    ' GL.glShaderSource handles null-termination and ANSI conversion internally
    GL.glShaderSource shader, src
    GL.glCompileShader shader

    Debug.Print "[modShaders] Compiled shader ID=" & shader & " type=" & shaderType
    CompileShader = shader
End Function

Public Function CreateProgram(ByVal vs As String, ByVal fs As String) As Long
    Dim vShader As Long, fShader As Long, prog As Long

    vShader = CompileShader(vs, GL.GL_VERTEX_SHADER)
    If vShader = 0 Then Exit Function

    fShader = CompileShader(fs, GL.GL_FRAGMENT_SHADER)
    If fShader = 0 Then
        GL.glDeleteShader vShader
        Exit Function
    End If

    prog = GL.glCreateProgram()
    If prog = 0 Then
        GL.glDeleteShader vShader
        GL.glDeleteShader fShader
        Exit Function
    End If

    GL.glAttachShader prog, vShader
    GL.glAttachShader prog, fShader
    GL.glLinkProgram prog

    ' Clean up shader objects (they're now part of the program)
    GL.glDeleteShader vShader
    GL.glDeleteShader fShader

    Debug.Print "[modShaders] Program created ID=" & prog
    CreateProgram = prog
End Function
