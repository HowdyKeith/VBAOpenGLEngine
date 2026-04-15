Option Explicit

' ============================================================
' Module     : modGL
' Version    : v8.43 FIXED FINAL
' CHANGES from v8.42:
'   - Added texture functions: glGenTextures, glBindTexture, glDeleteTextures,
'     glTexParameteri, glTexImage2D, glTexSubImage2D, glGenerateMipmap
'   - Added glBufferSubData (MeshBatcher FlushDynamic/FlushInstanced)
'   - Added GPU timer queries: glGenQueries, glDeleteQueries,
'     glBeginQuery, glEndQuery, glGetQueryObjectiv (RenderPipeline)
'   - Added glGetShaderiv (ShaderLoader)
'   - CopyMemory declared here for MeshBatcher (was used but never declared)
'   - Added GL constants: GL_TEXTURE_MIN/MAG_FILTER, GL_WRAP_S/T, GL_REPEAT,
'     GL_CLAMP_TO_EDGE, GL_TIME_ELAPSED, GL_QUERY_RESULT,
'     GL_SHADER_STORAGE_BARRIER_BIT, GL_ELEMENT_ARRAY_BUFFER
' ============================================================

' ============================================================
' SYSTEM MEMORY COPY (required by MeshBatcher CopyMemory calls)
' ============================================================
#If Win64 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal length As LongPtr)
    Private Declare PtrSafe Sub glEnable_Lib Lib "opengl32.dll" Alias "glEnable" (ByVal c As Long)
    Private Declare PtrSafe Sub glDisable_Lib Lib "opengl32.dll" Alias "glDisable" (ByVal c As Long)
    Private Declare PtrSafe Sub glClear_Lib Lib "opengl32.dll" Alias "glClear" (ByVal m As Long)
    Private Declare PtrSafe Sub glClearColor_Lib Lib "opengl32.dll" Alias "glClearColor" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare PtrSafe Sub glMatrixMode_Lib Lib "opengl32.dll" Alias "glMatrixMode" (ByVal m As Long)
    Private Declare PtrSafe Sub glLoadIdentity_Lib Lib "opengl32.dll" Alias "glLoadIdentity" ()
    Private Declare PtrSafe Sub glPushMatrix_Lib Lib "opengl32.dll" Alias "glPushMatrix" ()
    Private Declare PtrSafe Sub glPopMatrix_Lib Lib "opengl32.dll" Alias "glPopMatrix" ()
    Private Declare PtrSafe Sub glTranslatef_Lib Lib "opengl32.dll" Alias "glTranslatef" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glRotatef_Lib Lib "opengl32.dll" Alias "glRotatef" (ByVal a As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glScalef_Lib Lib "opengl32.dll" Alias "glScalef" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub gluPerspective_Lib Lib "glu32.dll" Alias "gluPerspective" (ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double)
    Private Declare PtrSafe Sub glBegin_Lib Lib "opengl32.dll" Alias "glBegin" (ByVal m As Long)
    Private Declare PtrSafe Sub glEnd_Lib Lib "opengl32.dll" Alias "glEnd" ()
    Private Declare PtrSafe Sub glVertex3f_Lib Lib "opengl32.dll" Alias "glVertex3f" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glNormal3f_Lib Lib "opengl32.dll" Alias "glNormal3f" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glPolygonMode_Lib Lib "opengl32.dll" Alias "glPolygonMode" (ByVal face As Long, ByVal mode As Long)
    Private Declare PtrSafe Sub glBlendFunc_Lib Lib "opengl32.dll" Alias "glBlendFunc" (ByVal sfactor As Long, ByVal dfactor As Long)
    Private Declare PtrSafe Sub glDrawArrays_Lib Lib "opengl32.dll" Alias "glDrawArrays" (ByVal mode As Long, ByVal first As Long, ByVal count As Long)
    Private Declare PtrSafe Sub glDrawElements_Lib Lib "opengl32.dll" Alias "glDrawElements" (ByVal mode As Long, ByVal count As Long, ByVal type_ As Long, ByVal indices As LongPtr)
    Private Declare PtrSafe Sub glGenTextures_Lib Lib "opengl32.dll" Alias "glGenTextures" (ByVal n As Long, ByRef textures As Long)
    Private Declare PtrSafe Sub glBindTexture_Lib Lib "opengl32.dll" Alias "glBindTexture" (ByVal target As Long, ByVal texture As Long)
    Private Declare PtrSafe Sub glDeleteTextures_Lib Lib "opengl32.dll" Alias "glDeleteTextures" (ByVal n As Long, ByRef textures As Long)
    Private Declare PtrSafe Sub glTexParameteri_Lib Lib "opengl32.dll" Alias "glTexParameteri" (ByVal target As Long, ByVal pname As Long, ByVal param As Long)
    Private Declare PtrSafe Sub glTexImage2D_Lib Lib "opengl32.dll" Alias "glTexImage2D" (ByVal target As Long, ByVal level As Long, ByVal internalfmt As Long, ByVal w As Long, ByVal h As Long, ByVal border As Long, ByVal format As Long, ByVal type_ As Long, ByVal pixels As LongPtr)
    Private Declare PtrSafe Sub glTexSubImage2D_Lib Lib "opengl32.dll" Alias "glTexSubImage2D" (ByVal target As Long, ByVal level As Long, ByVal xoff As Long, ByVal yoff As Long, ByVal w As Long, ByVal h As Long, ByVal format As Long, ByVal type_ As Long, ByVal pixels As LongPtr)
    Private Declare PtrSafe Function wglGetProcAddress Lib "opengl32.dll" (ByVal s As String) As LongPtr
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As Long, ByVal src As Long, ByVal length As Long)
    Private Declare Sub glEnable_Lib Lib "opengl32.dll" Alias "glEnable" (ByVal c As Long)
    Private Declare Sub glDisable_Lib Lib "opengl32.dll" Alias "glDisable" (ByVal c As Long)
    Private Declare Sub glClear_Lib Lib "opengl32.dll" Alias "glClear" (ByVal m As Long)
    Private Declare Sub glClearColor_Lib Lib "opengl32.dll" Alias "glClearColor" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare Sub glMatrixMode_Lib Lib "opengl32.dll" Alias "glMatrixMode" (ByVal m As Long)
    Private Declare Sub glLoadIdentity_Lib Lib "opengl32.dll" Alias "glLoadIdentity" ()
    Private Declare Sub glPushMatrix_Lib Lib "opengl32.dll" Alias "glPushMatrix" ()
    Private Declare Sub glPopMatrix_Lib Lib "opengl32.dll" Alias "glPopMatrix" ()
    Private Declare Sub glTranslatef_Lib Lib "opengl32.dll" Alias "glTranslatef" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glRotatef_Lib Lib "opengl32.dll" Alias "glRotatef" (ByVal a As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glScalef_Lib Lib "opengl32.dll" Alias "glScalef" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub gluPerspective_Lib Lib "glu32.dll" Alias "gluPerspective" (ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double)
    Private Declare Sub glBegin_Lib Lib "opengl32.dll" Alias "glBegin" (ByVal m As Long)
    Private Declare Sub glEnd_Lib Lib "opengl32.dll" Alias "glEnd" ()
    Private Declare Sub glVertex3f_Lib Lib "opengl32.dll" Alias "glVertex3f" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glNormal3f_Lib Lib "opengl32.dll" Alias "glNormal3f" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glPolygonMode_Lib Lib "opengl32.dll" Alias "glPolygonMode" (ByVal face As Long, ByVal mode As Long)
    Private Declare Sub glBlendFunc_Lib Lib "opengl32.dll" Alias "glBlendFunc" (ByVal sfactor As Long, ByVal dfactor As Long)
    Private Declare Sub glDrawArrays_Lib Lib "opengl32.dll" Alias "glDrawArrays" (ByVal mode As Long, ByVal first As Long, ByVal count As Long)
    Private Declare Sub glDrawElements_Lib Lib "opengl32.dll" Alias "glDrawElements" (ByVal mode As Long, ByVal count As Long, ByVal type_ As Long, ByVal indices As Long)
    Private Declare Sub glGenTextures_Lib Lib "opengl32.dll" Alias "glGenTextures" (ByVal n As Long, ByRef textures As Long)
    Private Declare Sub glBindTexture_Lib Lib "opengl32.dll" Alias "glBindTexture" (ByVal target As Long, ByVal texture As Long)
    Private Declare Sub glDeleteTextures_Lib Lib "opengl32.dll" Alias "glDeleteTextures" (ByVal n As Long, ByRef textures As Long)
    Private Declare Sub glTexParameteri_Lib Lib "opengl32.dll" Alias "glTexParameteri" (ByVal target As Long, ByVal pname As Long, ByVal param As Long)
    Private Declare Sub glTexImage2D_Lib Lib "opengl32.dll" Alias "glTexImage2D" (ByVal target As Long, ByVal level As Long, ByVal internalfmt As Long, ByVal w As Long, ByVal h As Long, ByVal border As Long, ByVal format As Long, ByVal type_ As Long, ByVal pixels As Long)
    Private Declare Sub glTexSubImage2D_Lib Lib "opengl32.dll" Alias "glTexSubImage2D" (ByVal target As Long, ByVal level As Long, ByVal xoff As Long, ByVal yoff As Long, ByVal w As Long, ByVal h As Long, ByVal format As Long, ByVal type_ As Long, ByVal pixels As Long)
    Private Declare Function wglGetProcAddress Lib "opengl32.dll" (ByVal s As String) As Long
#End If

' ============================================================
' GL CONSTANTS (texture, query, barrier)
' ============================================================
Public Const GL_TEXTURE_MIN_FILTER         As Long = &H2801
Public Const GL_TEXTURE_MAG_FILTER         As Long = &H2800
Public Const GL_TEXTURE_WRAP_S             As Long = &H2802
Public Const GL_TEXTURE_WRAP_T             As Long = &H2803
Public Const GL_REPEAT                     As Long = &H2901
Public Const GL_CLAMP_TO_EDGE              As Long = &H812F
Public Const GL_TIME_ELAPSED               As Long = &H88BF
Public Const GL_QUERY_RESULT               As Long = &H8866
Public Const GL_SHADER_STORAGE_BARRIER_BIT As Long = &H2000
Public Const GL_ELEMENT_ARRAY_BUFFER       As Long = &H8893

' ============================================================
' LEGACY api* BRIDGES
' ============================================================
Public Sub apiPolygonMode(ByVal f As Long, ByVal m As Long): glPolygonMode_Lib f, m: End Sub
Public Sub apiBlendFunc(ByVal s As Long, ByVal d As Long):   glBlendFunc_Lib s, d:   End Sub
Public Sub apiBegin(ByVal m As Long):     glBegin_Lib m:     End Sub
Public Sub apiEnd():                      glEnd_Lib:          End Sub
Public Sub apiVertex3f(ByVal x As Single, ByVal y As Single, ByVal z As Single): glVertex3f_Lib x, y, z: End Sub
Public Sub apiNormal3f(ByVal x As Single, ByVal y As Single, ByVal z As Single): glNormal3f_Lib x, y, z: End Sub
Public Sub apiEnable(ByVal c As Long):    glEnable_Lib c:     End Sub
Public Sub apiDisable(ByVal c As Long):   glDisable_Lib c:    End Sub
Public Sub apiClear(ByVal m As Long):     glClear_Lib m:      End Sub
Public Sub apiClearColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single): glClearColor_Lib r, g, b, a: End Sub
Public Sub apiMatrixMode(ByVal m As Long): glMatrixMode_Lib m: End Sub
Public Sub apiLoadIdentity():             glLoadIdentity_Lib: End Sub
Public Sub apiPushMatrix():               glPushMatrix_Lib:   End Sub
Public Sub apiPopMatrix():                glPopMatrix_Lib:    End Sub
Public Sub apiTranslatef(ByVal x As Single, ByVal y As Single, ByVal z As Single): glTranslatef_Lib x, y, z: End Sub
Public Sub apiRotatef(ByVal a As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single): glRotatef_Lib a, x, y, z: End Sub
Public Sub apiScalef(ByVal x As Single, ByVal y As Single, ByVal z As Single): glScalef_Lib x, y, z: End Sub
Public Sub apiPerspective(ByVal fovy As Double, ByVal asp As Double, ByVal zn As Double, ByVal zf As Double): gluPerspective_Lib fovy, asp, zn, zf: End Sub

' ============================================================
' STANDARD GL (direct opengl32 Lib calls)
' ============================================================
Public Sub glEnable(ByVal cap As Long):  glEnable_Lib cap:  End Sub
Public Sub glDisable(ByVal cap As Long): glDisable_Lib cap: End Sub
Public Sub glClear(ByVal mask As Long):  glClear_Lib mask:  End Sub
Public Sub glClearColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single): glClearColor_Lib r, g, b, a: End Sub
Public Sub glBlendFunc(ByVal s As Long, ByVal d As Long):  glBlendFunc_Lib s, d: End Sub
Public Sub glPolygonMode(ByVal face As Long, ByVal mode As Long): glPolygonMode_Lib face, mode: End Sub
Public Sub glDrawArrays(ByVal mode As Long, ByVal first As Long, ByVal count As Long): glDrawArrays_Lib mode, first, count: End Sub
Public Sub glDrawElements(ByVal mode As Long, ByVal count As Long, ByVal type_ As Long, ByVal indices As LongPtr): glDrawElements_Lib mode, count, type_, indices: End Sub

' --- Textures (in opengl32.dll directly, no wgl needed) ---
Public Sub glGenTextures(ByVal n As Long, ByRef textures As Long):    glGenTextures_Lib n, textures:    End Sub
Public Sub glBindTexture(ByVal target As Long, ByVal tex As Long):    glBindTexture_Lib target, tex:    End Sub
Public Sub glDeleteTextures(ByVal n As Long, ByRef textures As Long): glDeleteTextures_Lib n, textures: End Sub
Public Sub glTexParameteri(ByVal target As Long, ByVal pname As Long, ByVal param As Long): glTexParameteri_Lib target, pname, param: End Sub
Public Sub glTexImage2D(ByVal target As Long, ByVal level As Long, ByVal internalfmt As Long, _
                        ByVal w As Long, ByVal h As Long, ByVal border As Long, _
                        ByVal format As Long, ByVal type_ As Long, ByVal pixels As LongPtr)
    glTexImage2D_Lib target, level, internalfmt, w, h, border, format, type_, pixels
End Sub
Public Sub glTexSubImage2D(ByVal target As Long, ByVal level As Long, _
                           ByVal xoff As Long, ByVal yoff As Long, _
                           ByVal w As Long, ByVal h As Long, _
                           ByVal format As Long, ByVal type_ As Long, ByVal pixels As LongPtr)
    glTexSubImage2D_Lib target, level, xoff, yoff, w, h, format, type_, pixels
End Sub

' ============================================================
' EXTENSION-BACKED (delegate to GL.bas which owns the pointers)
' ============================================================
Public Sub glUseProgram(ByVal prog As Long):     GL.glUseProgram prog:        End Sub
Public Sub glBindVertexArray(ByVal vao As Long): GL.glBindVertexArray vao:    End Sub
Public Sub glActiveTexture(ByVal unit As Long):  GL.glActiveTexture unit:     End Sub
Public Sub glCullFace(ByVal mode As Long):       GL.glCullFace mode:          End Sub
Public Sub glViewport(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long): GL.glViewport x, y, w, h: End Sub
Public Function glGetError() As Long:            glGetError = GL.glGetError:  End Function
Public Sub glGenerateMipmap(ByVal target As Long): GL.glGenerateMipmap target: End Sub
Public Sub glGetShaderiv(ByVal shader As Long, ByVal pname As Long, ByRef params As Long): GL.glGetShaderiv shader, pname, params: End Sub
Public Sub glBufferSubData(ByVal target As Long, ByVal offset As LongPtr, ByVal size As LongPtr, ByVal data As LongPtr): GL.glBufferSubData target, offset, size, data: End Sub

Public Sub glDrawElementsInstanced(ByVal mode As Long, ByVal count As Long, ByVal type_ As Long, _
                                   ByVal indices As LongPtr, ByVal instanceCount As Long)
    GL.glDrawElementsInstanced mode, count, type_, indices, instanceCount
End Sub

' --- Shaders ---
Public Function glCreateShader(ByVal shaderType As Long) As Long: glCreateShader = GL.glCreateShader(shaderType): End Function
Public Sub glShaderSource(ByVal shader As Long, ByVal src As String): GL.glShaderSource shader, src: End Sub
Public Sub glCompileShader(ByVal shader As Long):  GL.glCompileShader shader:  End Sub
Public Function glCreateProgram() As Long:         glCreateProgram = GL.glCreateProgram(): End Function
Public Sub glAttachShader(ByVal p As Long, ByVal s As Long): GL.glAttachShader p, s: End Sub
Public Sub glLinkProgram(ByVal p As Long):         GL.glLinkProgram p:         End Sub
Public Sub glDeleteShader(ByVal s As Long):        GL.glDeleteShader s:        End Sub
Public Sub glDeleteProgram(ByVal p As Long):       GL.glDeleteProgram p:       End Sub

' --- GPU Timer Queries ---
Public Sub glGenQueries(ByVal n As Long, ByRef ids As Long):    GL.glGenQueries n, ids:          End Sub
Public Sub glDeleteQueries(ByVal n As Long, ByRef ids As Long): GL.glDeleteQueries n, ids:       End Sub
Public Sub glBeginQuery(ByVal target As Long, ByVal id As Long): GL.glBeginQuery target, id:     End Sub
Public Sub glEndQuery(ByVal target As Long):                     GL.glEndQuery target:            End Sub
Public Sub glGetQueryObjectiv(ByVal id As Long, ByVal pname As Long, ByRef params As Long): GL.glGetQueryObjectiv id, pname, params: End Sub

' ============================================================
' COLOUR (added Week 2 for primitive rendering)
' ============================================================
#If Win64 Then
    Private Declare PtrSafe Sub glColor3f_Lib Lib "opengl32.dll" Alias "glColor3f" (ByVal r As Single, ByVal g As Single, ByVal b As Single)
    Private Declare PtrSafe Sub glColor4f_Lib Lib "opengl32.dll" Alias "glColor4f" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
#Else
    Private Declare Sub glColor3f_Lib Lib "opengl32.dll" Alias "glColor3f" (ByVal r As Single, ByVal g As Single, ByVal b As Single)
    Private Declare Sub glColor4f_Lib Lib "opengl32.dll" Alias "glColor4f" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
#End If
Public Sub apiSetColor3f(ByVal r As Single, ByVal g As Single, ByVal b As Single): glColor3f_Lib r, g, b: End Sub
Public Sub apiSetColor4f(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single): glColor4f_Lib r, g, b, a: End Sub
Public Sub glColor3f(ByVal r As Single, ByVal g As Single, ByVal b As Single): glColor3f_Lib r, g, b: End Sub
Public Sub glColor4f(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single): glColor4f_Lib r, g, b, a: End Sub

' ============================================================
' POINT SIZE (added Week 3 for star map rendering)
' ============================================================
#If Win64 Then
    Private Declare PtrSafe Sub glPointSize_Lib Lib "opengl32.dll" Alias "glPointSize" (ByVal size As Single)
#Else
    Private Declare Sub glPointSize_Lib Lib "opengl32.dll" Alias "glPointSize" (ByVal size As Single)
#End If
Public Sub glPointSize(ByVal size As Single): glPointSize_Lib size: End Sub

Public Const GL_POINT_SIZE            As Long = &HB11
Public Const GL_PROGRAM_POINT_SIZE    As Long = &H8642   ' GL 3.2+ - lets shader set gl_PointSize
Public Const GL_POINT_SMOOTH          As Long = &HB10
