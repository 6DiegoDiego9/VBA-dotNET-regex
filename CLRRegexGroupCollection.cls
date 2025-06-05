VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegexGroupCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLRRegexGroupCollection.cls
Option Explicit

Private m_cobjGroupCollection As mscorlib.Object
Private m_host As CLRHost
Private m_groupCache As Collection ' Caches CLRRegexGroup VBA objects
Private m_isInitialized As Boolean
Private m_countCache As Long

Public Sub Initialize(ByVal cobjGCInstance As mscorlib.Object, ByVal hostInstance As CLRHost)
    Set m_host = hostInstance
    Set m_cobjGroupCollection = cobjGCInstance
    Set m_groupCache = New Collection
    
    If m_cobjGroupCollection Is Nothing Or m_host Is Nothing Then
        m_isInitialized = False
        m_countCache = 0
        Exit Sub
    End If
    
    On Error Resume Next
    m_countCache = m_host.CLRProperty(m_cobjGroupCollection, "Count")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexGroupCollection: Error getting Count: " & err.Description
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

Public Function Item(ByVal Key As Variant) As CLRRegexGroup
    Dim vbaGroup As CLRRegexGroup
    Set vbaGroup = New CLRRegexGroup ' Create upfront

    If Not m_isInitialized Then
        Call vbaGroup.Initialize(Nothing, m_host)
        Set Item = vbaGroup
        Exit Function
    End If

    If IsNumeric(Key) Then
        If CLng(Key) < 0 Or CLng(Key) >= m_countCache Then
            Call vbaGroup.Initialize(Nothing, m_host)
            Set Item = vbaGroup
            Exit Function
        End If
    End If
    
    On Error Resume Next ' Check cache
    Set vbaGroup = m_groupCache(CStr(Key))
    If err.Number <> 0 Or vbaGroup Is Nothing Then
        err.Clear
        On Error GoTo 0

        Dim cobjGroup As mscorlib.Object
        On Error Resume Next
        Set cobjGroup = m_host.CLRInvokeMethod(m_cobjGroupCollection, "get_Item", Key)
        
        Set vbaGroup = New CLRRegexGroup ' Re-New
        If err.Number = 0 And Not cobjGroup Is Nothing Then
            Call vbaGroup.Initialize(cobjGroup, m_host)
            On Error Resume Next ' Add to cache
             m_groupCache.Add vbaGroup, CStr(Key)
            err.Clear
        Else
            Debug.Print "CLRRegexGroupCollection: Error getting group by key '" & Key & "': " & err.Description
            Call vbaGroup.Initialize(Nothing, m_host)
            err.Clear
        End If
        On Error GoTo 0
    End If
    Set Item = vbaGroup
End Function

'Attribute NewEnum.VB_UserMemId = -4
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    If m_isInitialized And m_groupCache.Count <> m_countCache And m_countCache > 0 Then
        Dim i As Long
        For i = 0 To m_countCache - 1
            Dim tempGroup As CLRRegexGroup
            Set tempGroup = Me.Item(i) ' Populate cache by numeric index; string keys handled by Item property
        Next i
    End If
    Set NewEnum = m_groupCache.[_NewEnum]
End Function

Private Sub Class_Terminate()
    Set m_cobjGroupCollection = Nothing
    Set m_host = Nothing
    Set m_groupCache = Nothing
End Sub

