Option Explicit

' =========================================================
' InputSystem.v8.1.cls
' Keyboard & Mouse Polling System (VBA ECS v8.1)
' =========================================================

' =========================
' WIN32 DECLARES
' =========================
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
#End If

Private Type POINTAPI
    x As Long
    y As Long
End Type

' =========================
' STORAGE
' =========================
Private m_Current(0 To 255) As Boolean
Private m_Previous(0 To 255) As Boolean

Private m_MouseX As Long
Private m_MouseY As Long
Private m_PrevMouseX As Long
Private m_PrevMouseY As Long

' =========================
' UPDATE (CALL EVERY FRAME)
' =========================
Public Sub Update()
    Dim i As Long
    Dim state As Integer
    Dim pt As POINTAPI

    ' 1. Update Keyboard
    For i = 0 To 255
        m_Previous(i) = m_Current(i)
        state = GetAsyncKeyState(i)
        m_Current(i) = ((state And &H8000) <> 0)
    Next i

    ' 2. Update Mouse Position
    m_PrevMouseX = m_MouseX
    m_PrevMouseY = m_MouseY
    
    GetCursorPos pt
    ' Note: To get local coords, we'd need the HWND.
    ' For now, we store screen/global coords.
    m_MouseX = pt.x
    m_MouseY = pt.y
End Sub

' =========================
' WINDOW HANDLE (for client-space mouse coords)
' =========================
Public Sub SetWindowHandle(ByVal hWnd As LongPtr)
    m_hWnd = hWnd
End Sub
' =========================
' KEY ACCESSORS
' =========================
Public Function GetKey(ByVal keyCode As Long) As Boolean
    If keyCode >= 0 And keyCode <= 255 Then GetKey = m_Current(keyCode)
End Function

Public Function GetKeyDown(ByVal keyCode As Long) As Boolean
    If keyCode >= 0 And keyCode <= 255 Then GetKeyDown = m_Current(keyCode) And Not m_Previous(keyCode)
End Function

Public Function GetKeyUp(ByVal keyCode As Long) As Boolean
    If keyCode >= 0 And keyCode <= 255 Then GetKeyUp = Not m_Current(keyCode) And m_Previous(keyCode)
End Function

' =========================
' MOUSE ACCESSORS
' =========================
Public Function GetMouseX() As Long: GetMouseX = m_MouseX: End Function
Public Function GetMouseY() As Long: GetMouseY = m_MouseY: End Function
Public Function GetMouseDeltaX() As Long: GetMouseDeltaX = m_MouseX - m_PrevMouseX: End Function
Public Function GetMouseDeltaY() As Long: GetMouseDeltaY = m_MouseY - m_PrevMouseY: End Function

' Stub methods to prevent "Method Not Found" if called by Windows Events
Public Sub KeyDown(ByVal keyCode As Long): m_Current(keyCode) = True: End Sub
Public Sub KeyUp(ByVal keyCode As Long): m_Current(keyCode) = False: End Sub
Public Sub MouseMove(ByVal x As Long, ByVal y As Long): m_MouseX = x: m_MouseY = y: End Sub


