Option Explicit

' ============================================================
' TextureSystem.cls
' ECS Texture Manager + GPU Cache Layer
' Engine Target: VBA OpenGL v6/v7
' ============================================================

' =========================
' INTERNAL STORAGE
' =========================

Private m_Textures As Collection          ' key -> Texture2D
Private m_PathToKey As Collection         ' filePath -> key
Private m_NextID As Long

' =========================
' INITIALIZATION
' =========================

Private Sub Class_Initialize()
    Set m_Textures = New Collection
    Set m_PathToKey = New Collection
    m_NextID = 1
End Sub

' =========================
' CREATE / GET TEXTURE
' =========================

Public Function LoadTexture(ByVal filePath As String) As Texture2D

    Dim key As String
    Dim tex As Texture2D

    key = LCase$(filePath)

    ' --- Return cached texture if exists ---
    If TextureExists(filePath) Then
        Set LoadTexture = GetTexture(filePath)
        Exit Function
    End If

    ' --- Create new texture ---
    Set tex = New Texture2D

    If Not tex.LoadFromFile(filePath) Then
        Set LoadTexture = Nothing
        Exit Function
    End If

    ' --- Store in cache ---
    Call AddTextureInternal(filePath, tex)

    Set LoadTexture = tex

End Function

' =========================
' GET TEXTURE (NO LOAD)
' =========================

Public Function GetTexture(ByVal filePath As String) As Texture2D

    Dim key As String
    key = LCase$(filePath)

    On Error GoTo fail
    Set GetTexture = m_Textures(key)
    Exit Function

fail:
    Set GetTexture = Nothing
End Function

' =========================
' CHECK EXISTENCE
' =========================

Public Function TextureExists(ByVal filePath As String) As Boolean

    Dim key As String
    key = LCase$(filePath)

    On Error GoTo fail
    Dim t As Texture2D
    Set t = m_Textures(key)

    TextureExists = True
    Exit Function

fail:
    TextureExists = False
End Function

' =========================
' INTERNAL ADD
' =========================

Private Sub AddTextureInternal(ByVal filePath As String, ByRef tex As Texture2D)

    Dim key As String
    key = LCase$(filePath)

    On Error Resume Next

    m_Textures.Add tex, key
    m_PathToKey.Add key, key

End Sub

' =========================
' REMOVE TEXTURE
' =========================

Public Sub RemoveTexture(ByVal filePath As String)

    Dim key As String
    Dim tex As Texture2D

    key = LCase$(filePath)

    On Error GoTo done

    Set tex = m_Textures(key)
    tex.Destroy

    m_Textures.Remove key
    m_PathToKey.Remove key

done:
End Sub

' =========================
' BIND TEXTURE (SHADER SLOT)
' =========================

Public Sub BindTexture(ByVal filePath As String, Optional ByVal slot As Long = 0)

    Dim tex As Texture2D

    Set tex = GetTexture(filePath)

    If tex Is Nothing Then Exit Sub

    tex.Bind slot

End Sub

' =========================
' PRELOAD FOLDER
' =========================

Public Sub PreloadTextures(ByVal filePaths As Variant)

    Dim i As Long
    For i = LBound(filePaths) To UBound(filePaths)

        Call LoadTexture(CStr(filePaths(i)))

    Next i

End Sub

' =========================
' BULK GPU UPLOAD (SAFETY PASS)
' =========================

Public Sub UploadAllToGPU()

    Dim tex As Texture2D
    Dim key As Variant

    For Each key In m_Textures

        Set tex = m_Textures(key)
        tex.UploadToGPU

    Next key

End Sub

' =========================
' MEMORY STATS
' =========================

Public Function GetTextureCount() As Long
    GetTextureCount = m_Textures.count
End Function

' =========================
' DEBUG LIST
' =========================

Public Sub PrintTextureList()

    Dim key As Variant

    Debug.Print "==== TEXTURE SYSTEM ===="

    For Each key In m_Textures
        Debug.Print "Texture: " & key
    Next key

End Sub

' =========================
' CLEAR ALL
' =========================

Public Sub ClearAll()

    Dim tex As Texture2D
    Dim key As Variant

    For Each key In m_Textures

        Set tex = m_Textures(key)
        tex.Destroy

    Next key

    Set m_Textures = New Collection
    Set m_PathToKey = New Collection

End Sub

' =========================
' CLEANUP
' =========================

Private Sub Class_Terminate()
    Call ClearAll
End Sub



