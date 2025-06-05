VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegexGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLRRegexGroup.cls
Option Explicit

Private m_cobjGroup As mscorlib.Object
Private m_host As CLRHost
Private m_capturesCache As CLRRegexCaptureCollection

' Cached properties
Private m_successCache As Boolean
Private m_nameCache As String
Private m_valueCache As String
Private m_indexCache As Long
Private m_lengthCache As Long
Private m_isInitialized As Boolean

Public Sub Initialize(ByVal cobjGroupInstance As mscorlib.Object, ByVal hostInstance As CLRHost)
    Set m_cobjGroup = cobjGroupInstance
    Set m_host = hostInstance

    If m_cobjGroup Is Nothing Or m_host Is Nothing Then
        m_isInitialized = False
        m_successCache = False
        m_nameCache = ""
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
        Exit Sub
    End If

    On Error Resume Next
    m_successCache = m_host.CLRProperty(m_cobjGroup, "Success")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexGroup: Error getting Success: " & err.Description
        m_successCache = False
        m_isInitialized = False
        err.Clear
        ' Fall through to set default values for non-success
    End If
    err.Clear ' Clear potential error from CLRProperty
    On Error GoTo 0

    If m_successCache Then
        On Error Resume Next
        m_nameCache = m_host.CLRProperty(m_cobjGroup, "Name")
        m_valueCache = m_host.CLRProperty(m_cobjGroup, "Value")
        m_indexCache = m_host.CLRProperty(m_cobjGroup, "Index")
        m_lengthCache = m_host.CLRProperty(m_cobjGroup, "Length")
        If err.Number <> 0 Then
            Debug.Print "CLRRegexGroup: Error getting Name/Value/Index/Length: " & err.Description
            ' Invalidate if core properties fail, even if Success was true
            m_successCache = False ' Or handle as partially valid
            m_nameCache = ""
            m_valueCache = ""
            m_indexCache = -1
            m_lengthCache = 0
            err.Clear
        End If
        err.Clear
        On Error GoTo 0
    Else
        m_nameCache = ""
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
    End If
    m_isInitialized = True ' Mark as initialized even if not successful, to indicate init was attempted
End Sub

Public Property Get Success() As Boolean
    Success = m_successCache
End Property

Public Property Get Name() As String
    Name = m_nameCache
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

Public Property Get Captures() As CLRRegexCaptureCollection
    If m_capturesCache Is Nothing Then
        Set m_capturesCache = New CLRRegexCaptureCollection

        If Not m_isInitialized Or Not m_successCache Or m_cobjGroup Is Nothing Or m_host Is Nothing Then
            Call m_capturesCache.Initialize(Nothing, m_host)
        Else
            Dim cobjNetCaptures As mscorlib.Object
            On Error Resume Next
            Set cobjNetCaptures = m_host.CLRProperty(m_cobjGroup, "Captures")
            
            If err.Number = 0 And Not cobjNetCaptures Is Nothing Then
                Call m_capturesCache.Initialize(cobjNetCaptures, m_host)
            Else
                Debug.Print "CLRRegexGroup.Captures: Error getting .NET Captures property or object is null. Initializing as empty."
                err.Clear
                Call m_capturesCache.Initialize(Nothing, m_host)
            End If
            err.Clear
            On Error GoTo 0
        End If
    End If
    Set Captures = m_capturesCache
End Property

Private Sub Class_Terminate()
    Set m_cobjGroup = Nothing
    Set m_host = Nothing
    Set m_capturesCache = Nothing
End Sub

