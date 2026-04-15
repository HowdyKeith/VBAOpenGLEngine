Option Explicit

' =========================================================
' Module     : GL
' Version    : v9.4 FIXED FINAL
' ADDITIONS over v9.3:
'   - glBufferSubData, glGenerateMipmap, glGetShaderiv, glGetProgramiv
'   - glGenQueries, glDeleteQueries, glBeginQuery, glEndQuery, glGetQueryObjectiv
'   - glGetUniformLocationS() string convenience overload
'   - All extension pointer declarations added
' =========================================================

#If Win64 Then
    Private Declare PtrSafe Function wglGetProcAddress Lib "opengl32.dll" (ByVal s As String) As LongPtr
    Private Declare PtrSafe Sub apiClear Lib "opengl32.dll" Alias "glClear" (ByVal mask As Long)
    Private Declare PtrSafe Sub apiClearColor Lib "opengl32.dll" Alias "glClearColor" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare PtrSafe Sub apiEnable Lib "opengl32.dll" Alias "glEnable" (ByVal cap As Long)
    Private Declare PtrSafe Sub apiDisable Lib "opengl32.dll" Alias "glDisable" (ByVal cap As Long)
    Private Declare PtrSafe Sub apiDrawElements Lib "opengl32.dll" Alias "glDrawElements" (ByVal mode As Long, ByVal count As Long, ByVal type_ As Long, ByVal indices As LongPtr)
    Private Declare PtrSafe Sub apiDrawArrays Lib "opengl32.dll" Alias "glDrawArrays" (ByVal mode As Long, ByVal first As Long, ByVal count As Long)
    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cArgs As Long, ByRef rgvt As Integer, ByRef rgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
#Else
    Private Declare Function wglGetProcAddress Lib "opengl32.dll" (ByVal s As String) As Long
    Private Declare Sub apiClear Lib "opengl32.dll" Alias "glClear" (ByVal mask As Long)
    Private Declare Sub apiClearColor Lib "opengl32.dll" Alias "glClearColor" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare Sub apiEnable Lib "opengl32.dll" Alias "glEnable" (ByVal cap As Long)
    Private Declare Sub apiDisable Lib "opengl32.dll" Alias "glDisable" (ByVal cap As Long)
    Private Declare Sub apiDrawElements Lib "opengl32.dll" Alias "glDrawElements" (ByVal mode As Long, ByVal count As Long, ByVal type_ As Long, ByVal indices As Long)
    Private Declare Sub apiDrawArrays Lib "opengl32.dll" Alias "glDrawArrays" (ByVal mode As Long, ByVal first As Long, ByVal count As Long)
    Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cArgs As Long, ByRef rgvt As Integer, ByRef rgpvarg As Long, ByRef pvargResult As Variant) As Long
#End If

' FUNCTION POINTERS
Public p_glGenVertexArrays As LongPtr:   Public p_glBindVertexArray As LongPtr:  Public p_glDeleteVertexArrays As LongPtr
Public p_glGenBuffers As LongPtr:        Public p_glBindBuffer As LongPtr:       Public p_glBufferData As LongPtr
Public p_glBufferSubData As LongPtr:     Public p_glDeleteBuffers As LongPtr
Public p_glEnableVertexAttribArray As LongPtr: Public p_glVertexAttribPointer As LongPtr: Public p_glVertexAttribDivisor As LongPtr
Public p_glDrawElementsIndirect As LongPtr:    Public p_glMultiDrawElementsIndirect As LongPtr
Public p_glDrawElementsInstanced As LongPtr:   Public p_glDrawArraysInstanced As LongPtr
Public p_glBindBufferBase As LongPtr:    Public p_glDispatchCompute As LongPtr:  Public p_glMemoryBarrier As LongPtr
Public p_glUseProgram As LongPtr:        Public p_glCreateProgram As LongPtr:    Public p_glGetUniformLocation As LongPtr
Public p_glUniformMatrix4fv As LongPtr:  Public p_glUniform1i As LongPtr:        Public p_glUniform1f As LongPtr: Public p_glUniform3f As LongPtr
Public p_glActiveTexture As LongPtr:     Public p_glBindTexture As LongPtr:      Public p_glGenerateMipmap As LongPtr
Public p_glCreateShader As LongPtr:      Public p_glShaderSource As LongPtr:     Public p_glCompileShader As LongPtr
Public p_glAttachShader As LongPtr:      Public p_glLinkProgram As LongPtr:      Public p_glDeleteShader As LongPtr: Public p_glDeleteProgram As LongPtr
Public p_glGetShaderiv As LongPtr:       Public p_glGetProgramiv As LongPtr:     Public p_glGetShaderInfoLog As LongPtr
Public p_glCullFace As LongPtr:          Public p_glPolygonMode As LongPtr:      Public p_glBlendFunc As LongPtr
Public p_glViewport As LongPtr:          Public p_glGetError As LongPtr
Public p_glGenQueries As LongPtr:        Public p_glDeleteQueries As LongPtr
Public p_glBeginQuery As LongPtr:        Public p_glEndQuery As LongPtr:         Public p_glGetQueryObjectiv As LongPtr

' CONSTANTS
Public Const GL_COLOR_BUFFER_BIT As Long = &H4000&:  Public Const GL_DEPTH_BUFFER_BIT As Long = &H100&
Public Const GL_TRIANGLES As Long = &H4&:  Public Const GL_QUADS As Long = &H7
Public Const GL_POINTS As Long = 0:       Public Const GL_LINES As Long = 1
Public Const GL_ARRAY_BUFFER As Long = &H8892&:        Public Const GL_ELEMENT_ARRAY_BUFFER As Long = &H8893&
Public Const GL_SHADER_STORAGE_BUFFER As Long = &H90D2&: Public Const GL_DRAW_INDIRECT_BUFFER As Long = &H90EE&
Public Const GL_STATIC_DRAW As Long = &H88E4&:        Public Const GL_DYNAMIC_DRAW As Long = &H88E8&
Public Const GL_FLOAT As Long = &H1406:               Public Const GL_UNSIGNED_INT As Long = &H1405
Public Const GL_UNSIGNED_BYTE As Long = &H1401:       Public Const GL_UNSIGNED_SHORT As Long = &H1403
Public Const GL_FALSE As Long = 0:                    Public Const GL_TRUE As Long = 1
Public Const GL_DEPTH_TEST As Long = &HB71:           Public Const GL_BLEND As Long = &HBE2
Public Const GL_CULL_FACE As Long = &HB44:            Public Const GL_TEXTURE_2D As Long = &HDE1
Public Const GL_SRC_ALPHA As Long = &H302:            Public Const GL_ONE_MINUS_SRC_ALPHA As Long = &H303
Public Const GL_ONE As Long = 1:                      Public Const GL_TEXTURE0 As Long = &H84C0
Public Const GL_RGBA As Long = &H1908:                Public Const GL_RGB As Long = &H1907
Public Const GL_NEAREST As Long = &H2600:             Public Const GL_LINEAR As Long = &H2601
Public Const GL_TEXTURE_MIN_FILTER As Long = &H2801:  Public Const GL_TEXTURE_MAG_FILTER As Long = &H2800
Public Const GL_TEXTURE_WRAP_S As Long = &H2802:      Public Const GL_TEXTURE_WRAP_T As Long = &H2803
Public Const GL_REPEAT As Long = &H2901:              Public Const GL_CLAMP_TO_EDGE As Long = &H812F
Public Const GL_FRONT_AND_BACK As Long = &H408:       Public Const GL_FRONT As Long = &H404: Public Const GL_BACK As Long = &H405
Public Const GL_LINE As Long = &H1B01:                Public Const GL_FILL As Long = &H1B02
Public Const GL_MODELVIEW As Long = &H1700:           Public Const GL_PROJECTION As Long = &H1701
Public Const GL_VERTEX_SHADER As Long = &H8B31:       Public Const GL_FRAGMENT_SHADER As Long = &H8B30
Public Const GL_GEOMETRY_SHADER As Long = &H8DD9:     Public Const GL_COMPUTE_SHADER As Long = &H91B9
Public Const GL_COMPILE_STATUS As Long = &H8B81:      Public Const GL_LINK_STATUS As Long = &H8B82
Public Const GL_ALL_BARRIER_BITS As Long = &HFFFFFFFF: Public Const GL_SHADER_STORAGE_BARRIER_BIT As Long = &H2000
Public Const GL_TIME_ELAPSED As Long = &H88BF:        Public Const GL_QUERY_RESULT As Long = &H8866

Private Function CallGL(ByVal pFunc As LongPtr, ParamArray args() As Variant) As Variant
    If pFunc = 0 Then Exit Function
    Dim i As Long, count As Long: count = UBound(args) - LBound(args) + 1
    Dim vTypes() As Integer, vPtrs() As LongPtr
    If count > 0 Then
        ReDim vTypes(count - 1): ReDim vPtrs(count - 1)
        For i = 0 To count - 1: vTypes(i) = VarType(args(i)): vPtrs(i) = VarPtr(args(i)): Next
    End If
    DispCallFunc 0, pFunc, 4, vbLong, count, vTypes(0), vPtrs(0), CallGL
End Function

' DIRECT WRAPPERS
Public Sub glClear(ByVal mask As Long): apiClear mask: End Sub
Public Sub glClearColor(r As Single, g As Single, b As Single, a As Single): apiClearColor r, g, b, a: End Sub
Public Sub glEnable(cap As Long): apiEnable cap: End Sub
Public Sub glDisable(cap As Long): apiDisable cap: End Sub
Public Sub glDrawArrays(mode As Long, first As Long, count As Long): apiDrawArrays mode, first, count: End Sub
Public Sub glDrawElements(mode As Long, count As Long, type_ As Long, indices As LongPtr): apiDrawElements mode, count, type_, indices: End Sub

' Legacy matrix pipeline
Public Sub glMatrixMode(mode As Long):  CallGL p_glMatrixMode, mode: End Sub
Public Sub glLoadIdentity():            CallGL p_glLoadIdentity: End Sub
Public Sub glPushMatrix():              CallGL p_glPushMatrix: End Sub
Public Sub glPopMatrix():               CallGL p_glPopMatrix: End Sub
Public Sub glTranslatef(x As Single, y As Single, z As Single): CallGL p_glTranslatef, x, y, z: End Sub
Public Sub glRotatef(a As Single, x As Single, y As Single, z As Single): CallGL p_glRotatef, a, x, y, z: End Sub
Public Sub glScalef(x As Single, y As Single, z As Single): CallGL p_glScalef, x, y, z: End Sub
Public Sub glBegin(mode As Long): CallGL p_glBegin, mode: End Sub
Public Sub glEnd(): CallGL p_glEnd: End Sub
Public Sub glVertex3f(x As Single, y As Single, z As Single): CallGL p_glVertex3f, x, y, z: End Sub
Public Sub glNormal3f(x As Single, y As Single, z As Single): CallGL p_glNormal3f, x, y, z: End Sub
Public Sub gluPerspective(fovy As Double, aspect As Double, zNear As Double, zFar As Double): CallGL p_gluPerspective, fovy, aspect, zNear, zFar: End Sub

' VAO/VBO
Public Sub glGenVertexArrays(n As Long, ByRef arrays As Long):    CallGL p_glGenVertexArrays, n, VarPtr(arrays): End Sub
Public Sub glBindVertexArray(va As Long):                          CallGL p_glBindVertexArray, va: End Sub
Public Sub glDeleteVertexArrays(n As Long, ByRef arrays As Long): CallGL p_glDeleteVertexArrays, n, VarPtr(arrays): End Sub
Public Sub glGenBuffers(n As Long, ByRef buffers As Long):        CallGL p_glGenBuffers, n, VarPtr(buffers): End Sub
Public Sub glBindBuffer(target As Long, buffer As Long):           CallGL p_glBindBuffer, target, buffer: End Sub
Public Sub glBufferData(target As Long, size As LongPtr, data As LongPtr, usage As Long): CallGL p_glBufferData, target, size, data, usage: End Sub
Public Sub glBufferSubData(target As Long, offset As LongPtr, size As LongPtr, data As LongPtr): CallGL p_glBufferSubData, target, offset, size, data: End Sub
Public Sub glDeleteBuffers(n As Long, ByRef buffers As Long):     CallGL p_glDeleteBuffers, n, VarPtr(buffers): End Sub
Public Sub glEnableVertexAttribArray(index As Long): CallGL p_glEnableVertexAttribArray, index: End Sub
Public Sub glVertexAttribPointer(index As Long, size As Long, type_ As Long, normalized As Long, stride As Long, pointer As LongPtr)
    CallGL p_glVertexAttribPointer, index, size, type_, normalized, stride, pointer
End Sub
Public Sub glVertexAttribDivisor(index As Long, divisor As Long): CallGL p_glVertexAttribDivisor, index, divisor: End Sub

' Shaders
Public Function glCreateShader(shaderType As Long) As Long: glCreateShader = CLng(CallGL(p_glCreateShader, shaderType)): End Function
Public Sub glShaderSource(shader As Long, src As String)
    Dim b() As Byte: b = StrConv(src & vbNullChar, vbFromUnicode)
    Dim bPtr As LongPtr: bPtr = VarPtr(b(0))
    Dim one As Long: one = 1
    CallGL p_glShaderSource, shader, one, VarPtr(bPtr), 0
End Sub
Public Sub glCompileShader(shader As Long):  CallGL p_glCompileShader, shader: End Sub
Public Function glCreateProgram() As Long:   glCreateProgram = CLng(CallGL(p_glCreateProgram)): End Function
Public Sub glAttachShader(p As Long, s As Long): CallGL p_glAttachShader, p, s: End Sub
Public Sub glLinkProgram(p As Long):         CallGL p_glLinkProgram, p: End Sub
Public Sub glUseProgram(p As Long):          CallGL p_glUseProgram, p: End Sub
Public Sub glDeleteShader(s As Long):        CallGL p_glDeleteShader, s: End Sub
Public Sub glDeleteProgram(p As Long):       CallGL p_glDeleteProgram, p: End Sub
Public Sub glGetShaderiv(shader As Long, pname As Long, ByRef params As Long): CallGL p_glGetShaderiv, shader, pname, VarPtr(params): End Sub
Public Sub glGetProgramiv(p As Long, pname As Long, ByRef params As Long): CallGL p_glGetProgramiv, p, pname, VarPtr(params): End Sub

' GetUniformLocation - pointer version (used by GLUniform, ShaderProgram)
Public Function glGetUniformLocation(prog As Long, namePtr As LongPtr) As Long
    glGetUniformLocation = CLng(CallGL(p_glGetUniformLocation, prog, namePtr))
End Function
' String convenience overload (used by ComputeParticleSystem)
Public Function glGetUniformLocationS(prog As Long, name As String) As Long
    Dim b() As Byte: b = StrConv(name & vbNullChar, vbFromUnicode)
    glGetUniformLocationS = CLng(CallGL(p_glGetUniformLocation, prog, VarPtr(b(0))))
End Function

Public Sub glUniformMatrix4fv(location As Long, count As Long, transpose As Long, value As LongPtr): CallGL p_glUniformMatrix4fv, location, count, transpose, value: End Sub
Public Sub glUniform1i(location As Long, v0 As Long): CallGL p_glUniform1i, location, v0: End Sub
Public Sub glUniform1f(location As Long, v0 As Single): CallGL p_glUniform1f, location, v0: End Sub
Public Sub glUniform3f(location As Long, x As Single, y As Single, z As Single): CallGL p_glUniform3f, location, x, y, z: End Sub

' Texture (extension-backed for modern GL, but in opengl32 for GL 1.x compat)
Public Sub glActiveTexture(unit As Long): CallGL p_glActiveTexture, unit: End Sub
Public Sub glBindTexture(target As Long, tex As Long): CallGL p_glBindTexture, target, tex: End Sub
Public Sub glGenerateMipmap(target As Long): CallGL p_glGenerateMipmap, target: End Sub

' Render state
Public Sub glCullFace(mode As Long): CallGL p_glCullFace, mode: End Sub
Public Sub glPolygonMode(face As Long, mode As Long): CallGL p_glPolygonMode, face, mode: End Sub
Public Sub glBlendFunc(s As Long, d As Long): CallGL p_glBlendFunc, s, d: End Sub
Public Sub glViewport(x As Long, y As Long, w As Long, h As Long): CallGL p_glViewport, x, y, w, h: End Sub
Public Function glGetError() As Long: glGetError = CLng(CallGL(p_glGetError)): End Function

' Instanced/indirect
Public Sub glDrawElementsInstanced(mode As Long, count As Long, type_ As Long, indices As LongPtr, instanceCount As Long)
    CallGL p_glDrawElementsInstanced, mode, count, type_, indices, instanceCount
End Sub
Public Sub glDrawElementsIndirect(mode As Long, type_ As Long, indirect As LongPtr): CallGL p_glDrawElementsIndirect, mode, type_, indirect: End Sub
Public Sub glMultiDrawElementsIndirect(mode As Long, type_ As Long, indirect As LongPtr, drawcount As Long, stride As Long)
    CallGL p_glMultiDrawElementsIndirect, mode, type_, indirect, drawcount, stride
End Sub

' SSBO/Compute
Public Sub glBindBufferBase(target As Long, index As Long, buffer As Long): CallGL p_glBindBufferBase, target, index, buffer: End Sub
Public Sub glDispatchCompute(x As Long, y As Long, z As Long): CallGL p_glDispatchCompute, x, y, z: End Sub
Public Sub glMemoryBarrier(barriers As Long): CallGL p_glMemoryBarrier, barriers: End Sub

' GPU Timer Queries
Public Sub glGenQueries(n As Long, ByRef ids As Long):    CallGL p_glGenQueries, n, VarPtr(ids): End Sub
Public Sub glDeleteQueries(n As Long, ByRef ids As Long): CallGL p_glDeleteQueries, n, VarPtr(ids): End Sub
Public Sub glBeginQuery(target As Long, id As Long):      CallGL p_glBeginQuery, target, id: End Sub
Public Sub glEndQuery(target As Long):                    CallGL p_glEndQuery, target: End Sub
Public Sub glGetQueryObjectiv(id As Long, pname As Long, ByRef params As Long): CallGL p_glGetQueryObjectiv, id, pname, VarPtr(params): End Sub

' INITIALIZATION
Public Sub GL_Init()
    LoadExtensions
    LoadLegacyPointers
    Debug.Print "[GL] v9.4 Final - All extensions loaded."
End Sub

Public Sub LoadShaderExtensions()
    If p_glCreateShader <> 0 Then Exit Sub
    p_glCreateShader       = wglGetProcAddress("glCreateShader")
    p_glShaderSource       = wglGetProcAddress("glShaderSource")
    p_glCompileShader      = wglGetProcAddress("glCompileShader")
    p_glCreateProgram      = wglGetProcAddress("glCreateProgram")
    p_glAttachShader       = wglGetProcAddress("glAttachShader")
    p_glLinkProgram        = wglGetProcAddress("glLinkProgram")
    p_glUseProgram         = wglGetProcAddress("glUseProgram")
    p_glDeleteShader       = wglGetProcAddress("glDeleteShader")
    p_glDeleteProgram      = wglGetProcAddress("glDeleteProgram")
    p_glGetUniformLocation = wglGetProcAddress("glGetUniformLocation")
    p_glUniformMatrix4fv   = wglGetProcAddress("glUniformMatrix4fv")
    p_glUniform1i          = wglGetProcAddress("glUniform1i")
    p_glUniform1f          = wglGetProcAddress("glUniform1f")
    p_glUniform3f          = wglGetProcAddress("glUniform3f")
    p_glGetShaderiv        = wglGetProcAddress("glGetShaderiv")
    p_glGetProgramiv       = wglGetProcAddress("glGetProgramiv")
    p_glGetShaderInfoLog   = wglGetProcAddress("glGetShaderInfoLog")
End Sub

Private Sub LoadExtensions()
    p_glGenVertexArrays          = wglGetProcAddress("glGenVertexArrays")
    p_glBindVertexArray          = wglGetProcAddress("glBindVertexArray")
    p_glDeleteVertexArrays       = wglGetProcAddress("glDeleteVertexArrays")
    p_glGenBuffers               = wglGetProcAddress("glGenBuffers")
    p_glBindBuffer               = wglGetProcAddress("glBindBuffer")
    p_glBufferData               = wglGetProcAddress("glBufferData")
    p_glBufferSubData            = wglGetProcAddress("glBufferSubData")
    p_glDeleteBuffers            = wglGetProcAddress("glDeleteBuffers")
    p_glEnableVertexAttribArray  = wglGetProcAddress("glEnableVertexAttribArray")
    p_glVertexAttribPointer      = wglGetProcAddress("glVertexAttribPointer")
    p_glVertexAttribDivisor      = wglGetProcAddress("glVertexAttribDivisor")
    p_glActiveTexture            = wglGetProcAddress("glActiveTexture")
    p_glBindTexture              = wglGetProcAddress("glBindTexture")
    p_glGenerateMipmap           = wglGetProcAddress("glGenerateMipmap")
    p_glCullFace                 = wglGetProcAddress("glCullFace")
    p_glPolygonMode              = wglGetProcAddress("glPolygonMode")
    p_glBlendFunc                = wglGetProcAddress("glBlendFunc")
    p_glViewport                 = wglGetProcAddress("glViewport")
    p_glGetError                 = wglGetProcAddress("glGetError")
    p_glDrawElementsInstanced    = wglGetProcAddress("glDrawElementsInstanced")
    p_glDrawElementsIndirect     = wglGetProcAddress("glDrawElementsIndirect")
    p_glMultiDrawElementsIndirect = wglGetProcAddress("glMultiDrawElementsIndirect")
    p_glDrawArraysInstanced      = wglGetProcAddress("glDrawArraysInstanced")
    p_glBindBufferBase           = wglGetProcAddress("glBindBufferBase")
    p_glDispatchCompute          = wglGetProcAddress("glDispatchCompute")
    p_glMemoryBarrier            = wglGetProcAddress("glMemoryBarrier")
    p_glGenQueries               = wglGetProcAddress("glGenQueries")
    p_glDeleteQueries            = wglGetProcAddress("glDeleteQueries")
    p_glBeginQuery               = wglGetProcAddress("glBeginQuery")
    p_glEndQuery                 = wglGetProcAddress("glEndQuery")
    p_glGetQueryObjectiv         = wglGetProcAddress("glGetQueryObjectiv")
    LoadShaderExtensions
End Sub

Private p_glMatrixMode As LongPtr:  Private p_glLoadIdentity As LongPtr
Private p_glPushMatrix As LongPtr:  Private p_glPopMatrix As LongPtr
Private p_glTranslatef As LongPtr:  Private p_glRotatef As LongPtr:  Private p_glScalef As LongPtr
Private p_glBegin As LongPtr:       Private p_glEnd As LongPtr
Private p_glVertex3f As LongPtr:    Private p_glNormal3f As LongPtr
Private p_gluPerspective As LongPtr

Private Sub LoadLegacyPointers()
    p_glMatrixMode   = GLLoader.p_glMatrixMode
    p_glLoadIdentity = GLLoader.p_glLoadIdentity
    p_glPushMatrix   = GLLoader.p_glPushMatrix
    p_glPopMatrix    = GLLoader.p_glPopMatrix
    p_glTranslatef   = GLLoader.p_glTranslatef
    p_glRotatef      = GLLoader.p_glRotatef
    p_glScalef       = GLLoader.p_glScalef
    p_glBegin        = GLLoader.p_glBegin
    p_glEnd          = GLLoader.p_glEnd
    p_glVertex3f     = GLLoader.p_glVertex3f
    p_glNormal3f     = GLLoader.p_glNormal3f
    p_gluPerspective = GLLoader.p_glGluPerspective
End Sub

' ============================================================
' POINT SIZE (added Week 3)
' ============================================================
#If Win64 Then
    Private Declare PtrSafe Sub glPointSize_Lib Lib "opengl32.dll" Alias "glPointSize" (ByVal size As Single)
#Else
    Private Declare Sub glPointSize_Lib Lib "opengl32.dll" Alias "glPointSize" (ByVal size As Single)
#End If
Public Sub glPointSize(ByVal size As Single): glPointSize_Lib size: End Sub
Public Const GL_POINT_SIZE         As Long = &HB11
Public Const GL_PROGRAM_POINT_SIZE As Long = &H8642
