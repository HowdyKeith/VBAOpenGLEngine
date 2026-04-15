Option Explicit

' ============================================================
' Texture2D.cls
' VBA OpenGL Texture Wrapper (GDI+ Loader + GL Upload)
' Engine Target: v7 ECS / RenderSystem pipeline
' ============================================================

' =========================
' OPENGL HANDLE
' =========================

Public textureID As LongPtr
Private m_IsLoaded As Boolean

' =========================
' METADATA
' =========================

Public width As Long
Public Height As Long
Public Channels As Long

Public filePath As String

' Raw pixel buffer (RGBA)
Private m_Pixels() As Byte

' =========================
' INITIALIZATION
' =========================

Private Sub Class_Initialize()
    textureID = 0
    m_IsLoaded = False
    width = 0
    Height = 0
    Channels = 4
End Sub

' =========================
' LOAD FROM FILE (GDI+)
' =========================

Public Function LoadFromFile(ByVal path As String) As Boolean

    On Error GoTo fail

    filePath = path

    Dim hBitmap As LongPtr
    Dim bmpData() As Byte

    ' --- Load image via GDI+ helper ---
    If Not LoadImageToRGBA(path, bmpData, width, Height) Then
        LoadFromFile = False
        Exit Function
    End If

    m_Pixels = bmpData
    Channels = 4

    Call UploadToGPU

    LoadFromFile = True
    Exit Function

fail:
    LoadFromFile = False
End Function

' =========================
' GPU UPLOAD
' =========================

Public Sub UploadToGPU()

    If m_IsLoaded Then Exit Sub
    If width = 0 Or Height = 0 Then Exit Sub

    modGL.glGenTextures 1, textureID
    modGL.glBindTexture GL_TEXTURE_2D, textureID

    ' --- Filtering ---
    modGL.glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    modGL.glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR

    ' --- Wrapping ---
    modGL.glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
    modGL.glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT

    ' --- Upload pixel data ---
    modGL.glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, _
                 width, Height, 0, _
                 GL_RGBA, GL_UNSIGNED_BYTE, m_Pixels(0)

    modGL.glGenerateMipmap GL_TEXTURE_2D

    modGL.glBindTexture GL_TEXTURE_2D, 0

    m_IsLoaded = True

End Sub

' =========================
' BIND / UNBIND
' =========================

Public Sub Bind(Optional ByVal slot As Long = 0)

    If Not m_IsLoaded Then Exit Sub

    modGL.glActiveTexture GL_TEXTURE0 + slot
    modGL.glBindTexture GL_TEXTURE_2D, textureID

End Sub

Public Sub Unbind()

    modGL.glBindTexture GL_TEXTURE_2D, 0

End Sub

' =========================
' CREATE EMPTY TEXTURE
' =========================

Public Sub CreateEmpty(ByVal w As Long, ByVal h As Long)

    width = w
    Height = h
    Channels = 4

    ReDim m_Pixels(w * h * 4 - 1)

    Dim i As Long
    For i = 0 To UBound(m_Pixels)
        m_Pixels(i) = 255
    Next i

    Call UploadToGPU

End Sub

' =========================
' UPDATE TEXTURE DATA
' =========================

Public Sub UpdatePixels(ByRef newPixels() As Byte)

    If Not m_IsLoaded Then Exit Sub

    m_Pixels = newPixels

    modGL.glBindTexture GL_TEXTURE_2D, textureID

    modGL.glTexSubImage2D GL_TEXTURE_2D, 0, _
                    0, 0, width, Height, _
                    GL_RGBA, GL_UNSIGNED_BYTE, m_Pixels(0)

    modGL.glBindTexture GL_TEXTURE_2D, 0

End Sub

' =========================
' CLEANUP
' =========================

Public Sub Destroy()

    If m_IsLoaded Then
        modGL.glDeleteTextures 1, textureID
    End If

    textureID = 0
    m_IsLoaded = False

End Sub

' =========================
' GDI+ IMAGE LOADER (RGBA)
' =========================

Private Function LoadImageToRGBA(ByVal filePath As String, _
                                ByRef outPixels() As Byte, _
                                ByRef outW As Long, _
                                ByRef outH As Long) As Boolean

    On Error GoTo fail

    ' NOTE:
    ' This assumes you already have a GDI+ wrapper module like:
    ' - GDIPlus_Init
    ' - GDIPlus_LoadBitmap
    ' - GDIPlus_GetBitmapData

    Dim hBitmap As LongPtr

    If Not GDIPlus_LoadBitmap(filePath, hBitmap, outW, outH) Then
        LoadImageToRGBA = False
        Exit Function
    End If

    If Not GDIPlus_GetBitmapData(hBitmap, outPixels, outW, outH) Then
        LoadImageToRGBA = False
        Exit Function
    End If

    Call GDIPlus_DisposeBitmap(hBitmap)

    LoadImageToRGBA = True
    Exit Function

fail:
    LoadImageToRGBA = False
End Function





