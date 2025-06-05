VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegexMatchCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLRRegexMatchCollection.cls
Option Explicit

Private m_cobjMatchCollection As mscorlib.Object
Private m_host As CLRHost
Private m_matchCache As Collection ' Caches CLRRegexMatch VBA objects
Private m_isInitialized As Boolean
Private m_countCache As Long

Public Sub Initialize(ByVal cobjMCInstance As mscorlib.Object, ByVal hostInstance As CLRHost)
    Set m_host = hostInstance ' Set host first, as it might be needed even if cobjMCInstance is Nothing
    Set m_cobjMatchCollection = cobjMCInstance
    Set m_matchCache = New Collection ' Always create a new VBA collection
    
    If m_cobjMatchCollection Is Nothing Or m_host Is Nothing Then
        m_isInitialized = False
        m_countCache = 0
        Exit Sub
    End If

    On Error Resume Next
    m_countCache = m_host.CLRProperty(m_cobjMatchCollection, "Count")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexMatchCollection: Error getting Count: " & err.Description
        m_countCache = 0
        m_isInitialized = False
        err.Clear
        Exit Sub
    End If
    err.Clear
    On Error GoTo 0
    m_isInitialized = True
End Sub

Public Property Get Count() As Long
    Count = m_countCache
End Property

Public Function Item(ByVal idx As Long) As CLRRegexMatch ' idx is 0-based
    Dim vbaMatch As CLRRegexMatch
    Set vbaMatch = New CLRRegexMatch ' Create upfront

    If Not m_isInitialized Or idx < 0 Or idx >= m_countCache Then
        Call vbaMatch.Initialize(Nothing, m_host) ' Return an unsuccessful match
        Set Item = vbaMatch
        Exit Function
    End If

    On Error Resume Next ' Check cache
    Set vbaMatch = m_matchCache(CStr(idx + 1)) ' VBA Collection is 1-based key
    If err.Number <> 0 Or vbaMatch Is Nothing Then
        err.Clear
        On Error GoTo 0 ' Reset error handling for CLR call

        Dim cobjMatch As mscorlib.Object
        On Error Resume Next
        Set cobjMatch = m_host.CLRInvokeMethod(m_cobjMatchCollection, "get_Item", idx)
        
        Set vbaMatch = New CLRRegexMatch ' Re-New in case previous cache check failed weirdly
        If err.Number = 0 And Not cobjMatch Is Nothing Then
            Call vbaMatch.Initialize(cobjMatch, m_host)
            ' Only add to cache if successfully initialized.
            ' This assumes vbaMatch.Initialize sets some internal state to check validity.
            ' Or more simply, just cache it. Let's cache it.
            On Error Resume Next ' In case CStr(idx+1) is already a key (shouldn't happen if logic is right)
             m_matchCache.Add vbaMatch, CStr(idx + 1)
            err.Clear
        Else
            Debug.Print "CLRRegexMatchCollection: Error getting item at index " & idx & ": " & err.Description
            Call vbaMatch.Initialize(Nothing, m_host) ' Initialize as unsuccessful
            err.Clear
        End If
        On Error GoTo 0
    End If
    Set Item = vbaMatch
End Function

'Attribute NewEnum.VB_UserMemId = -4
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    ' Ensure the cache is fully populated before allowing enumeration
    ' This is crucial for For Each loops to work correctly.
    If m_isInitialized And m_matchCache.Count <> m_countCache And m_countCache > 0 Then
        Dim i As Long
        For i = 0 To m_countCache - 1
            ' Call Item property to fetch/cache if not present.
            ' The Item property itself handles adding to m_matchCache.
            Dim tempMatch As CLRRegexMatch
            Set tempMatch = Me.Item(i)
        Next i
    End If
    ' Always return the enumerator of m_matchCache.
    ' If m_matchCache is empty (e.g., due to initialization failure or no matches),
    ' this will correctly return an enumerator that yields no items.
    Set NewEnum = m_matchCache.[_NewEnum]
End Function

Private Sub Class_Terminate()
    Set m_cobjMatchCollection = Nothing
    Set m_host = Nothing
    Set m_matchCache = Nothing
End Sub

