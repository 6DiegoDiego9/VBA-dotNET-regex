VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegexCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLRRegexCapture.cls
Option Explicit

Private m_cobjCapture As mscorlib.Object
Private m_host As CLRHost

' Cached properties
Private m_valueCache As String
Private m_indexCache As Long
Private m_lengthCache As Long
Private m_isInitialized As Boolean ' Indicates if Initialize was called and properties cached

Public Sub Initialize(ByVal cobjCaptureInstance As mscorlib.Object, ByVal hostInstance As CLRHost)
    Set m_cobjCapture = cobjCaptureInstance
    Set m_host = hostInstance

    If m_cobjCapture Is Nothing Or m_host Is Nothing Then
        m_isInitialized = False
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
        Exit Sub
    End If
    
    On Error Resume Next
    m_valueCache = m_host.CLRProperty(m_cobjCapture, "Value")
    m_indexCache = m_host.CLRProperty(m_cobjCapture, "Index")
    m_lengthCache = m_host.CLRProperty(m_cobjCapture, "Length")
    If err.Number <> 0 Then
        Debug.Print "CLRRegexCapture: Error initializing properties: " & err.Description
        m_isInitialized = False ' Failed to cache properties
        m_valueCache = ""
        m_indexCache = -1
        m_lengthCache = 0
        err.Clear
        Exit Sub
    End If
    err.Clear
    On Error GoTo 0
    m_isInitialized = True
End Sub

Public Property Get Value() As String
    Value = m_valueCache
End Property

Public Property Get Index() As Long
    Index = m_indexCache
End Property

Public Property Get Length() As Long
    Length = m_lengthCache
End Property

' Added to check if object is usable
Public Property Get IsValid() As Boolean
    IsValid = m_isInitialized ' And potentially other checks like m_cobjCapture Is Not Nothing
End Property

Private Sub Class_Terminate()
    Set m_cobjCapture = Nothing
    Set m_host = Nothing
End Sub

