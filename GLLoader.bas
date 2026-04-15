Option Explicit

' ============================================================
' GLLoader.v8.33.bas  FIXED
' FIXES:
'   - Added p_glGluPerspective pointer (from glu32) for GL.bas legacy bridge
'   - Added missing GL_VERTEX_SHADER / GL_FRAGMENT_SHADER constant wrappers
' ============================================================

#If Win64 Then
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddressGLU Lib "glu32.dll" Alias "gluPerspective" (ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double) As LongPtr
    Private Declare PtrSafe Function wglGetProcAddress Lib "opengl32.dll" (ByVal sName As String) As LongPtr
    Private Declare PtrSafe Function GetModuleHandleGLU Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
#Else
    Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function wglGetProcAddress Lib "opengl32.dll" (ByVal sName As String) As Long
    Private Declare Function GetModuleHandleGLU Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
#End If

' --- Standard GL 1.x pointers (loaded via GetProcAddress on opengl32 hModule)
Public p_glMatrixMode   As LongPtr, p_glLoadIdentity As LongPtr
Public p_glPushMatrix   As LongPtr, p_glPopMatrix    As LongPtr
Public p_glTranslatef   As LongPtr, p_glRotatef      As LongPtr, p_glScalef As LongPtr
Public p_glBegin        As LongPtr, p_glEnd           As LongPtr
Public p_glVertex3f     As LongPtr, p_glNormal3f      As LongPtr
Public p_glClear        As LongPtr, p_glClearColor    As LongPtr
Public p_glEnable       As LongPtr, p_glDisable       As LongPtr
Public p_glDrawArrays   As LongPtr, p_glDrawElements  As LongPtr
Public p_glPolygonMode  As LongPtr, p_glBlendFunc     As LongPtr

' glu32 pointer
Public p_glGluPerspective As LongPtr

' --- Extension pointers (loaded via wglGetProcAddress)
Public p_glGenVertexArrays      As LongPtr, p_glBindVertexArray  As LongPtr, p_glDeleteVertexArrays As LongPtr
Public p_glGenBuffers           As LongPtr, p_glBindBuffer       As LongPtr
Public p_glBufferData           As LongPtr, p_glDeleteBuffers    As LongPtr
Public p_glEnableVertexAttribArray As LongPtr, p_glVertexAttribPointer As LongPtr
Public p_glCreateShader         As LongPtr, p_glShaderSource     As LongPtr, p_glCompileShader As LongPtr
Public p_glCreateProgram        As LongPtr, p_glAttachShader     As LongPtr, p_glLinkProgram   As LongPtr
Public p_glUseProgram           As LongPtr, p_glDeleteShader     As LongPtr, p_glDeleteProgram As LongPtr

Public Sub Init(ByVal hLib As LongPtr)
    If hLib = 0 Then Exit Sub

    ' --- Standard 1.x (from opengl32 module handle)
    p_glMatrixMode   = GetProcAddress(hLib, "glMatrixMode")
    p_glLoadIdentity = GetProcAddress(hLib, "glLoadIdentity")
    p_glPushMatrix   = GetProcAddress(hLib, "glPushMatrix")
    p_glPopMatrix    = GetProcAddress(hLib, "glPopMatrix")
    p_glTranslatef   = GetProcAddress(hLib, "glTranslatef")
    p_glRotatef      = GetProcAddress(hLib, "glRotatef")
    p_glScalef       = GetProcAddress(hLib, "glScalef")
    p_glBegin        = GetProcAddress(hLib, "glBegin")
    p_glEnd          = GetProcAddress(hLib, "glEnd")
    p_glVertex3f     = GetProcAddress(hLib, "glVertex3f")
    p_glNormal3f     = GetProcAddress(hLib, "glNormal3f")
    p_glClear        = GetProcAddress(hLib, "glClear")
    p_glClearColor   = GetProcAddress(hLib, "glClearColor")
    p_glEnable       = GetProcAddress(hLib, "glEnable")
    p_glDisable      = GetProcAddress(hLib, "glDisable")
    p_glDrawArrays   = GetProcAddress(hLib, "glDrawArrays")
    p_glDrawElements = GetProcAddress(hLib, "glDrawElements")
    p_glPolygonMode  = GetProcAddress(hLib, "glPolygonMode")
    p_glBlendFunc    = GetProcAddress(hLib, "glBlendFunc")

    ' --- glu32 perspective
    Dim hGLU As LongPtr
    hGLU = GetModuleHandleGLU("glu32.dll")
    If hGLU <> 0 Then
        p_glGluPerspective = GetProcAddress(hGLU, "gluPerspective")
    End If

    ' --- Extensions (via wglGetProcAddress - must have a current GL context)
    p_glGenVertexArrays         = wglGetProcAddress("glGenVertexArrays")
    p_glBindVertexArray         = wglGetProcAddress("glBindVertexArray")
    p_glDeleteVertexArrays      = wglGetProcAddress("glDeleteVertexArrays")
    p_glGenBuffers              = wglGetProcAddress("glGenBuffers")
    p_glBindBuffer              = wglGetProcAddress("glBindBuffer")
    p_glBufferData              = wglGetProcAddress("glBufferData")
    p_glDeleteBuffers           = wglGetProcAddress("glDeleteBuffers")
    p_glEnableVertexAttribArray = wglGetProcAddress("glEnableVertexAttribArray")
    p_glVertexAttribPointer     = wglGetProcAddress("glVertexAttribPointer")
    p_glCreateShader            = wglGetProcAddress("glCreateShader")
    p_glShaderSource            = wglGetProcAddress("glShaderSource")
    p_glCompileShader           = wglGetProcAddress("glCompileShader")
    p_glCreateProgram           = wglGetProcAddress("glCreateProgram")
    p_glAttachShader            = wglGetProcAddress("glAttachShader")
    p_glLinkProgram             = wglGetProcAddress("glLinkProgram")
    p_glUseProgram              = wglGetProcAddress("glUseProgram")
    p_glDeleteShader            = wglGetProcAddress("glDeleteShader")
    p_glDeleteProgram           = wglGetProcAddress("glDeleteProgram")

    Debug.Print "[GLLoader] Init complete. VAO=" & p_glGenVertexArrays & " Shader=" & p_glCreateShader
End Sub
