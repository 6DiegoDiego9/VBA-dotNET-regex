VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegexCaptureCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLRRegexCaptureCollection.cls
Option Explicit

Private m_cobjCaptureCollection As mscorlib.Object
Private m_host As CLRHost
Private m_captureCache As Collection ' Caches CLRRegexCapture VBA objects
Private m_isInitialized As Boolean
Private m_countCache As Long

Public Sub Initialize(ByVal cobjCCInstance As mscorlib.Object, ByVal hostInstance As CLRHost)
    Set m_host = hostInstance
    Set m_cobjCaptureCollection = cobjCCInstance
    Set m_captureCache = New Collection

    If m_cobjCaptureCollection Is Nothing Or m_host Is Nothing Then
        m_isInitialized = False
        m_countCache = 0
        Exit Sub
    End If

    On Error Resume Next
    m_countCache = m_host.CLRProperty(m_cobjCaptureCollection, "Count")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexCaptureCollection: Error getting Count: " & err.Description
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

Public Function Item(ByVal idx As Long) As CLRRegexCapture ' idx is 0-based
    Dim vbaCapture As CLRRegexCapture
    Set vbaCapture = New CLRRegexCapture ' Create upfront

    If Not m_isInitialized Or idx < 0 Or idx >= m_countCache Then
        Call vbaCapture.Initialize(Nothing, m_host)
        Set Item = vbaCapture
        Exit Function
    End If

    On Error Resume Next ' Check cache
    Set vbaCapture = m_captureCache(CStr(idx + 1))
    If err.Number <> 0 Or vbaCapture Is Nothing Then
        err.Clear
        On Error GoTo 0
        
        Dim cobjCapture As mscorlib.Object
        On Error Resume Next
        Set cobjCapture = m_host.CLRInvokeMethod(m_cobjCaptureCollection, "get_Item", idx)
        
        Set vbaCapture = New CLRRegexCapture ' Re-New
        If err.Number = 0 And Not cobjCapture Is Nothing Then
            Call vbaCapture.Initialize(cobjCapture, m_host)
            On Error Resume Next ' Add to cache
             m_captureCache.Add vbaCapture, CStr(idx + 1)
            err.Clear
        Else
            Debug.Print "CLRRegexCaptureCollection: Error getting item at " & idx & ": " & err.Description
            Call vbaCapture.Initialize(Nothing, m_host)
            err.Clear
        End If
        On Error GoTo 0
    End If
    Set Item = vbaCapture
End Function

'Attribute NewEnum.VB_UserMemId = -4
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    If m_isInitialized And m_captureCache.Count <> m_countCache And m_countCache > 0 Then
        Dim i As Long
        For i = 0 To m_countCache - 1
            Dim tempCap As CLRRegexCapture
            Set tempCap = Me.Item(i)
        Next i
    End If
    Set NewEnum = m_captureCache.[_NewEnum]
End Property

Private Sub Class_Terminate()
    Set m_cobjCaptureCollection = Nothing
    Set m_host = Nothing
    Set m_captureCache = Nothing
End Sub

