Option Explicit

' ============================================================
' OpenGLWindow.v8.8.cls
' BRIDGE BETWEEN FACADE AND WIN32 WINDOW MANAGER
' ============================================================

' --- Win32 API for Module Loading ---
#If Win64 Then
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As LongPtr
#Else
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
#End If

Private m_hWnd As LongPtr
Private m_Running As Boolean

' =========================
' CORE LIFECYCLE
' =========================

Public Function Create(ByVal Title As String, ByVal width As Long, ByVal Height As Long) As Boolean
    ' Call the module-level manager to create the Win32 window and GL context
    m_hWnd = modWindowManager.CreateGLWindow(Title, width, Height)
    
    If m_hWnd <> 0 Then
        ' 1. Get the handle to the OpenGL library directly using Win32 API
        Dim hLib As LongPtr
        hLib = GetModuleHandle("opengl32.dll")
        
        ' 2. Initialize the pointers in GLLoader
        If hLib <> 0 Then
            GLLoader.Init hLib
            m_Running = True
            Create = True
            Debug.Print "[OpenGLWindow] Window Created and GL Pointers Initialized."
        Else
            Debug.Print "[OpenGLWindow] Critical Error: Could not find opengl32.dll handle."
            Create = False
        End If
    Else
        Debug.Print "[OpenGLWindow] Critical Error: Failed to create window handle."
        Create = False
    End If
End Function

' =========================
' ENGINE LOOP METHODS
' =========================

Public Function IsRunning() As Boolean
    If m_Running Then
        m_Running = modWindowManager.ProcessMessages()
    End If
    IsRunning = m_Running
End Function

Public Sub Update()
    ' Hook for physics/logic
End Sub

Public Sub Render()
    ' Hook for frame-start logic
End Sub

Public Sub Present()
    modWindowManager.PageFlip
End Sub

' =========================
' CLEANUP
' =========================

Public Sub CloseWindow()
    Me.Destroy
End Sub

Public Sub Destroy()
    m_Running = False
    modWindowManager.CloseGLWindow
    m_hWnd = 0
    Debug.Print "[OpenGLWindow] Resources Released."
End Sub


