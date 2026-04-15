Option Explicit

' ============================================================
' Class      : AtlasRegion
' Version    : v1.0
' Description: A named UV region within a TextureAtlas.
'              Was referenced by TextureAtlasSystem.cls but never defined.
' ============================================================

' Normalised UV coordinates (0..1)
Public x As Single   ' left
Public y As Single   ' top
Public w As Single   ' width
Public h As Single   ' height

' Optional metadata
Public key    As String
Public width  As Long    ' pixel width of the source texture
Public Height As Long    ' pixel height of the source texture

Public Sub SetUV(ByVal px As Single, ByVal py As Single, _
                 ByVal pw As Single, ByVal ph As Single)
    x = px: y = py: w = pw: h = ph
End Sub

Public Sub DebugPrint()
    Debug.Print "AtlasRegion [" & key & "] UV=(" & x & "," & y & ") Size=(" & w & "," & h & ")"
End Sub
