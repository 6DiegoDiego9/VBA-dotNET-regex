VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegexMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLRRegexMatch.cls
Option Explicit

Private m_cobjMatch As mscorlib.Object
Private m_host As CLRHost
Private m_groupsCache As CLRRegexGroupCollection

' Cached properties
Private m_successCache As Boolean
Private m_valueCache As String
Private m_indexCache As Long
Private m_lengthCache As Long
Private m_isInitialized As Boolean

Public Sub Initialize(ByVal cobjMatchInstance As mscorlib.Object, ByVal hostInstance As CLRHost)
    Set m_cobjMatch = cobjMatchInstance
    Set m_host = hostInstance
    
    If m_cobjMatch Is Nothing Or m_host Is Nothing Then
        m_isInitialized = False
        m_successCache = False
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
        Exit Sub
    End If
    
    On Error Resume Next
    m_successCache = m_host.CLRProperty(m_cobjMatch, "Success")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexMatch: Error getting Success property: " & err.Description
        m_successCache = False
        m_isInitialized = False
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
        err.Clear
        Exit Sub
    End If
    err.Clear
    On Error GoTo 0
    
    If m_successCache Then
        On Error Resume Next
        m_valueCache = m_host.CLRProperty(m_cobjMatch, "Value")
        m_indexCache = m_host.CLRProperty(m_cobjMatch, "Index")
        m_lengthCache = m_host.CLRProperty(m_cobjMatch, "Length")
        If err.Number <> 0 Then
             Debug.Print "CLRRegexMatch: Error getting Value/Index/Length properties: " & err.Description
             ' Potentially invalidate success if core props fail
             m_successCache = False ' Or handle more gracefully
             m_valueCache = ""
             m_indexCache = -1
             m_lengthCache = 0
             err.Clear
        End If
        err.Clear
        On Error GoTo 0
    Else
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
    End If
    m_isInitialized = True
End Sub

Public Property Get Success() As Boolean
    Success = m_successCache
End Property

Public Property Get Value() As String
    Value = m_valueCache
End Property

Public Property Get Index() As Long
    Index = m_indexCache
End Property

Public Property Get Length() As Long
    Length = m_lengthCache
End Property

Public Property Get Groups() As CLRRegexGroupCollection
    If m_groupsCache Is Nothing Then
        Set m_groupsCache = New CLRRegexGroupCollection

        If Not m_isInitialized Or Not m_successCache Or m_cobjMatch Is Nothing Or m_host Is Nothing Then
            Call m_groupsCache.Initialize(Nothing, m_host)
        Else
            Dim cobjNetGroups As mscorlib.Object
            On Error Resume Next
            Set cobjNetGroups = m_host.CLRProperty(m_cobjMatch, "Groups")
            
            If err.Number = 0 And Not cobjNetGroups Is Nothing Then
                Call m_groupsCache.Initialize(cobjNetGroups, m_host)
            Else
                Debug.Print "CLRRegexMatch.Groups: Error getting .NET Groups property or object is null. Initializing as empty."
                err.Clear
                Call m_groupsCache.Initialize(Nothing, m_host)
            End If
            err.Clear
            On Error GoTo 0
        End If
    End If
    Set Groups = m_groupsCache
End Property

Public Function NextMatch() As CLRRegexMatch
    Dim vbaNextMatch As New CLRRegexMatch
    If Not m_isInitialized Or Not m_successCache Or m_cobjMatch Is Nothing Or m_host Is Nothing Then
        Call vbaNextMatch.Initialize(Nothing, m_host)
        Set NextMatch = vbaNextMatch
        Exit Function
    End If
    
    Dim cobjNextMatch As mscorlib.Object
    On Error Resume Next
    Set cobjNextMatch = m_host.CLRInvokeMethod(m_cobjMatch, "NextMatch")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexMatch: Error invoking NextMatch: " & err.Description
        Call vbaNextMatch.Initialize(Nothing, m_host)
    Else
        Call vbaNextMatch.Initialize(cobjNextMatch, m_host)
    End If
    err.Clear
    On Error GoTo 0
    
    Set NextMatch = vbaNextMatch ' Always return the match object, user checks .Success
End Function

Private Sub Class_Terminate()
    Set m_cobjMatch = Nothing
    Set m_host = Nothing
    Set m_groupsCache = Nothing
End Sub

