Option Explicit

' ============================================================
' CollisionQuery.cls  FIXED
' Physics Query Descriptor
' FIXES:
'   - "Public TypeFIX As QueryType" - the field was named TypeFIX because
'     "Type" is a VBA reserved word and can't be used as a field name.
'     Renamed to QueryKind for clarity. All internal usages updated.
' ============================================================

Public Enum QueryType
    qtRaycast = 0
    qtSphere  = 1
    qtBox     = 2
    qtCapsule = 3
    qtOverlap = 4
End Enum

' FIXED: was "TypeFIX" - renamed to QueryKind
Public QueryKind As QueryType

Public Origin      As Vector3
Public direction   As Vector3
Public distance    As Single
Public Radius      As Single
Public HalfExtents As Vector3

Public IgnoreEntity    As Long
Public LayerMask       As Long
Public IncludeTriggers As Boolean

Private Sub Class_Initialize()
    Set Origin      = New Vector3
    Set direction   = New Vector3
    Set HalfExtents = New Vector3
    QueryKind       = qtRaycast
    distance        = 100
    Radius          = 1
    IgnoreEntity    = -1
    LayerMask       = -1
    IncludeTriggers = False
End Sub

Private Sub Class_Terminate()
    Set Origin      = Nothing
    Set direction   = Nothing
    Set HalfExtents = Nothing
End Sub

Public Sub SetRaycast(ByVal ox As Single, ByVal oy As Single, ByVal oz As Single, _
                      ByVal dx As Single, ByVal dy As Single, ByVal dz As Single, _
                      Optional ByVal maxDist As Single = 100)
    QueryKind = qtRaycast
    Origin.SetValues ox, oy, oz
    direction.SetValues dx, dy, dz
    distance = maxDist
End Sub

Public Sub SetSphere(ByVal cx As Single, ByVal cy As Single, ByVal cz As Single, _
                     ByVal r As Single)
    QueryKind = qtSphere
    Origin.SetValues cx, cy, cz
    Radius = r
End Sub

Public Sub SetBox(ByVal cx As Single, ByVal cy As Single, ByVal cz As Single, _
                  ByVal hx As Single, ByVal hy As Single, ByVal hz As Single)
    QueryKind = qtBox
    Origin.SetValues cx, cy, cz
    HalfExtents.SetValues hx, hy, hz
End Sub

Public Sub DebugPrint()
    Debug.Print "---- CollisionQuery [" & QueryKind & "] ----"
    Debug.Print "  Origin: "; Origin.x; Origin.y; Origin.z
End Sub
