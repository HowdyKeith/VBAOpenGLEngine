Option Explicit

' ============================================================
' GLConstants.bas  FIXED v1.2
' Standard OpenGL Hex Values
' FIXES:
'   - Added all constants that were missing but used elsewhere:
'     GL_DEPTH_TEST, GL_BLEND, GL_UNSIGNED_INT, GL_FLOAT,
'     GL_VERTEX_SHADER, GL_FRAGMENT_SHADER, GL_MODELVIEW, GL_PROJECTION,
'     GL_FRONT_AND_BACK, GL_LINE, GL_FILL, GL_QUADS, GL_SRC_ALPHA etc.
' NOTE: GL.bas also defines many of these. Having them here too is safe in VBA
'       as long as both modules are in the same project - VBA resolves unqualified
'       references module by module. Use GL.GL_* for explicit qualification.
' ============================================================

' --- Clear bits ---
Public Const GL_STENCIL_BUFFER_BIT As Long = &H400

' --- Primitives ---
Public Const GL_POINTS         As Long = &H0
Public Const GL_LINES          As Long = &H1
Public Const GL_TRIANGLES      As Long = &H4
Public Const GL_TRIANGLE_STRIP As Long = &H5
Public Const GL_TRIANGLE_FAN   As Long = &H6
Public Const GL_QUADS          As Long = &H7

' --- Data Types ---
Public Const GL_BYTE           As Long = &H1400
Public Const GL_UNSIGNED_BYTE  As Long = &H1401
Public Const GL_UNSIGNED_SHORT As Long = &H1403
Public Const GL_UNSIGNED_INT   As Long = &H1405
Public Const GL_FLOAT          As Long = &H1406
Public Const GL_FALSE          As Long = 0
Public Const GL_TRUE           As Long = 1

' --- Capabilities ---
Public Const GL_DEPTH_TEST     As Long = &HB71
Public Const GL_BLEND          As Long = &HBE2
Public Const GL_CULL_FACE      As Long = &HB44
Public Const GL_LIGHTING       As Long = &HB50
Public Const GL_TEXTURE_2D     As Long = &HDE1

' --- Blend factors ---
Public Const GL_SRC_ALPHA           As Long = &H302
Public Const GL_ONE_MINUS_SRC_ALPHA As Long = &H303
Public Const GL_ONE                 As Long = 1

' --- Shaders ---
Public Const GL_VERTEX_SHADER   As Long = &H8B31
Public Const GL_FRAGMENT_SHADER As Long = &H8B30
Public Const GL_GEOMETRY_SHADER As Long = &H8DD9
Public Const GL_COMPILE_STATUS  As Long = &H8B81
Public Const GL_LINK_STATUS     As Long = &H8B82
Public Const GL_INFO_LOG_LENGTH As Long = &H8B84

' --- Matrix modes (legacy pipeline) ---
Public Const GL_MODELVIEW  As Long = &H1700
Public Const GL_PROJECTION As Long = &H1701

' --- Buffers ---
Public Const GL_DYNAMIC_DRAW As Long = &H88E8

' --- Textures ---
Public Const GL_TEXTURE0 As Long = &H84C0
Public Const GL_RGBA     As Long = &H1908
Public Const GL_RGB      As Long = &H1907
Public Const GL_NEAREST  As Long = &H2600
Public Const GL_LINEAR   As Long = &H2601

' --- Polygon / face ---
Public Const GL_FRONT_AND_BACK As Long = &H408
Public Const GL_FRONT          As Long = &H404
Public Const GL_BACK           As Long = &H405
Public Const GL_LINE           As Long = &H1B01
Public Const GL_FILL           As Long = &H1B02

' --- Misc ---
Public Const GL_VERSION As Long = &H1F02
