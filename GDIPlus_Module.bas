Option Explicit

' ============================================================
' GDIPlus_Module.bas
' Safe PNG/JPG Loader -> RGBA Byte Array
' VBA OpenGL Engine v6/v7 Dependency
' ============================================================

' =========================
' WINDOWS / GDI+ DECLARES
' =========================

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus.dll" ( _
    ByRef token As LongPtr, _
    ByRef inputbuf As GdiplusStartupInput, _
    ByVal outputbuf As LongPtr) As Long

Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus.dll" ( _
    ByVal token As LongPtr)

Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus.dll" ( _
    ByVal fileName As LongPtr, _
    ByRef image As LongPtr) As Long

Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus.dll" ( _
    ByVal image As LongPtr) As Long

Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus.dll" ( _
    ByVal image As LongPtr, _
    ByRef width As Long) As Long

Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus.dll" ( _
    ByVal image As LongPtr, _
    ByRef Height As Long) As Long

Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus.dll" ( _
    ByVal bitmap As LongPtr, _
    ByRef rect As RectL, _
    ByVal flags As Long, _
    ByVal format As Long, _
    ByRef lockedBitmapData As BitmapData) As Long

Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus.dll" ( _
    ByVal bitmap As LongPtr, _
    ByRef lockedBitmapData As BitmapData) As Long

' =========================
' STRUCTURES
' =========================

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type RectL
    x As Long
    y As Long
    width As Long
    Height As Long
End Type

Private Type BitmapData
    width As Long
    Height As Long
    stride As Long
    pixelFormat As Long
    Scan0 As LongPtr
    Reserved As LongPtr
End Type

' =========================
' GLOBAL STATE
' =========================

Private m_Token As LongPtr
Private m_Initialized As Boolean

' =========================
' INIT / SHUTDOWN
' =========================

Public Function GDIPlus_Init() As Boolean

    Dim si As GdiplusStartupInput
    si.GdiplusVersion = 1

    If m_Initialized Then
        GDIPlus_Init = True
        Exit Function
    End If

    If GdiplusStartup(m_Token, si, 0) = 0 Then
        m_Initialized = True
        GDIPlus_Init = True
    Else
        GDIPlus_Init = False
    End If

End Function

Public Sub GDIPlus_Shutdown()

    If m_Initialized Then
        GdiplusShutdown m_Token
    End If

    m_Initialized = False

End Sub

' =========================
' LOAD IMAGE -> RGBA BYTE ARRAY
' =========================

Public Function GDIPlus_GetBitmapData( _
    ByVal filePath As String, _
    ByRef outPixels() As Byte, _
    ByRef outW As Long, _
    ByRef outH As Long) As Boolean

    On Error GoTo fail

    Dim img As LongPtr
    Dim bmpData As BitmapData
    Dim rect As RectL

    If Not m_Initialized Then Call GDIPlus_Init

    If GdipLoadImageFromFile(StrPtr(filePath), img) <> 0 Then
        GDIPlus_GetBitmapData = False
        Exit Function
    End If

    Call GdipGetImageWidth(img, outW)
    Call GdipGetImageHeight(img, outH)

    rect.x = 0
    rect.y = 0
    rect.width = outW
    rect.Height = outH

    bmpData.stride = 0

    ' Lock bits (32bpp ARGB)
    If GdipBitmapLockBits(img, rect, 3, &H26200A, bmpData) <> 0 Then
        GdipDisposeImage img
        GDIPlus_GetBitmapData = False
        Exit Function
    End If

    Dim size As Long
    size = outW * outH * 4

    ReDim outPixels(size - 1)

    Dim i As Long
    Dim srcPtr As LongPtr
    srcPtr = bmpData.Scan0

    CopyMemory outPixels(0), ByVal srcPtr, size

    Call GdipBitmapUnlockBits(img, bmpData)
    Call GdipDisposeImage(img)

    GDIPlus_GetBitmapData = True
    Exit Function

fail:
    GDIPlus_GetBitmapData = False
End Function

' =========================
' LOAD BITMAP WRAPPER
' =========================

Public Function GDIPlus_LoadBitmap( _
    ByVal filePath As String, _
    ByRef outBitmap As LongPtr, _
    ByRef outW As Long, _
    ByRef outH As Long) As Boolean

    On Error GoTo fail

    If Not m_Initialized Then Call GDIPlus_Init

    If GdipLoadImageFromFile(StrPtr(filePath), outBitmap) <> 0 Then
        GDIPlus_LoadBitmap = False
        Exit Function
    End If

    Call GdipGetImageWidth(outBitmap, outW)
    Call GdipGetImageHeight(outBitmap, outH)

    GDIPlus_LoadBitmap = True
    Exit Function

fail:
    GDIPlus_LoadBitmap = False
End Function

' =========================
' DISPOSE BITMAP
' =========================

Public Sub GDIPlus_DisposeBitmap(ByVal bitmap As LongPtr)

    If bitmap <> 0 Then
        GdipDisposeImage bitmap
    End If

End Sub





