Option Explicit

' ============================================================
' Module     : ShaderLoader
' Version    : v1.1 FIXED
' FIXES:
'   - GL.glShaderSource was called as (shaderID, 1, codePtr, 0) - old 4-arg
'     C-API style. GL.bas v9.4 now exposes glShaderSource(shader, srcString)
'     which handles ANSI conversion internally. Updated to use new signature.
'   - GL_COMPILE_STATUS, GL_VERTEX_SHADER, GL_FRAGMENT_SHADER were unqualified
'     constants - now GL.GL_COMPILE_STATUS etc.
'   - glGetShaderiv result checked properly
' ============================================================

Public Function LoadShader(ByVal filePath As String, ByVal shaderType As Long) As Long
    Dim shaderCode As String
    shaderCode = ReadTextFile(filePath)

    If shaderCode = "" Then
        Debug.Print "[ShaderLoader] ERROR: Could not read file: " & filePath
        LoadShader = 0
        Exit Function
    End If

    Dim shaderID As Long
    shaderID = GL.glCreateShader(shaderType)
    If shaderID = 0 Then
        Debug.Print "[ShaderLoader] ERROR: glCreateShader returned 0 for: " & filePath
        Exit Function
    End If

    ' FIXED: Use GL.glShaderSource(shader, string) - handles ANSI conversion internally
    GL.glShaderSource shaderID, shaderCode
    GL.glCompileShader shaderID

    ' Check compile status
    Dim success As Long
    GL.glGetShaderiv shaderID, GL.GL_COMPILE_STATUS, success
    If success = 0 Then
        Debug.Print "[ShaderLoader] ERROR: Compile failed for: " & filePath
        GL.glDeleteShader shaderID
        LoadShader = 0
        Exit Function
    End If

    Debug.Print "[ShaderLoader] Compiled: " & filePath & " (ID=" & shaderID & ")"
    LoadShader = shaderID
End Function

Public Function CreateProgram(ByVal vertexPath As String, ByVal fragmentPath As String) As Long
    Dim vert As Long, frag As Long, prog As Long

    ' FIXED: GL.GL_VERTEX_SHADER / GL.GL_FRAGMENT_SHADER (qualified)
    vert = LoadShader(vertexPath, GL.GL_VERTEX_SHADER)
    If vert = 0 Then
        Debug.Print "[ShaderLoader] ERROR: Vertex shader failed: " & vertexPath
        Exit Function
    End If

    frag = LoadShader(fragmentPath, GL.GL_FRAGMENT_SHADER)
    If frag = 0 Then
        Debug.Print "[ShaderLoader] ERROR: Fragment shader failed: " & fragmentPath
        GL.glDeleteShader vert
        Exit Function
    End If

    prog = GL.glCreateProgram()
    If prog = 0 Then
        GL.glDeleteShader vert
        GL.glDeleteShader frag
        Exit Function
    End If

    GL.glAttachShader prog, vert
    GL.glAttachShader prog, frag
    GL.glLinkProgram prog

    GL.glDeleteShader vert
    GL.glDeleteShader frag

    Debug.Print "[ShaderLoader] Program linked: " & prog
    CreateProgram = prog
End Function

' Helper: read a text file via FSO
Private Function ReadTextFile(ByVal path As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1)
        ReadTextFile = ts.ReadAll
        ts.Close
    Else
        Debug.Print "[ShaderLoader] File not found: " & path
        ReadTextFile = ""
    End If
End Function
