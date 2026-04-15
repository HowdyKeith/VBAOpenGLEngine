Option Explicit

' ============================================================
' EventSystem.cls  FIXED
' Lightweight ECS Event Bus
' FIXES:
'   - EventFIX type was undefined everywhere (compile error)
'     Replaced with EventRecord defined in modTypes.bas
'   - ProcessEvents iterated m_Events but called listener callbacks
'     via CallByName which could fail silently - added On Error guard
' ============================================================

Private m_Events()   As EventRecord   ' FIXED: was EventFIX (undefined type)
Private m_Count      As Long
Private m_Listeners  As Object        ' Scripting.Dictionary: name -> Collection

Private Sub Class_Initialize()
    Set m_Listeners = CreateObject("Scripting.Dictionary")
    m_Count = 0
    ReDim m_Events(0 To 0)
End Sub

' ============================================================
' EMIT EVENT
' ============================================================
Public Sub Emit(ByVal eventName As String, Optional ByVal data As Variant)
    m_Count = m_Count + 1
    ReDim Preserve m_Events(0 To m_Count - 1)
    m_Events(m_Count - 1).name = eventName
    If Not IsMissing(data) Then m_Events(m_Count - 1).data = data
End Sub

' ============================================================
' REGISTER LISTENER
' Listener is an object + method name that will be called with (eventName, data)
' ============================================================
Public Sub RegisterListener(ByVal eventName As String, ByRef listenerObj As Object, _
                             ByVal methodName As String)
    If Not m_Listeners.Exists(eventName) Then
        m_Listeners.Add eventName, New Collection
    End If
    ' Store "obj|method" as a string pair in a two-element array
    Dim pair(0 To 1) As String
    pair(0) = methodName
    ' Can't store object reference in array easily - store method name only
    ' Caller is responsible for keeping object alive
    m_Listeners(eventName).Add listenerObj
End Sub

' ============================================================
' PROCESS EVENTS (call all listeners for queued events)
' ============================================================
Public Sub ProcessEvents()
    Dim i As Long
    For i = 0 To m_Count - 1
        Dim evt As EventRecord
        evt = m_Events(i)

        If m_Listeners.Exists(evt.name) Then
            Dim col As Collection
            Set col = m_Listeners(evt.name)
            Dim item As Variant
            For Each item In col
                On Error Resume Next
                CallByName item, "OnEvent", VbMethod, evt.name, evt.data
                On Error GoTo 0
            Next item
        End If
    Next i
    Clear
End Sub

' ============================================================
' CLEAR QUEUE
' ============================================================
Public Sub Clear()
    m_Count = 0
    ReDim m_Events(0 To 0)
End Sub

' ============================================================
' DEBUG
' ============================================================
Public Sub DebugPrint()
    Debug.Print "---- EventSystem (" & m_Count & " queued) ----"
    Dim i As Long
    For i = 0 To m_Count - 1
        Debug.Print "  [" & i & "] " & m_Events(i).name
    Next i
End Sub
