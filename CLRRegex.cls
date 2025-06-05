VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_host As CLRHost
Private m_asmSystem As mscorlib.Assembly
Private m_cobjRegex As mscorlib.Object
Private m_pattern As String
Private m_options As Long ' Represents System.Text.RegularExpressions.RegexOptions

Public Enum RegexOptionsCLR
    None = 0
    ignoreCase = 1
    multiLine = 2
    ExplicitCapture = 4
    compiled = 8
    Singleline = 16
    IgnorePatternWhitespace = 32
    RightToLeft = 64
    ECMAScript = 256
    CultureInvariant = 512
End Enum

' ============================
' Lifecycle helpers
' ============================
Private Sub Class_Initialize()
    Set m_host = New CLRHost
    On Error Resume Next
    Call m_host.Initialize(False)
    If err.Number <> 0 Then
        Debug.Print "FATAL: CLRHost initialization failed: " & err.Description
        Set m_host = Nothing
        Exit Sub
    End If
    err.Clear

    Set m_asmSystem = m_host.CLRLoadAssembly("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    If err.Number <> 0 Then
        Debug.Print "Error loading System assembly: " & err.Description
        If Not m_host Is Nothing Then Set m_host = Nothing
        Exit Sub
    End If
    err.Clear
    On Error GoTo 0
End Sub

' ============================
' Public API
' ============================
Public Sub InitializeRegex(ByVal rgxPattern As String, Optional ByVal rgxOptions As RegexOptionsCLR = RegexOptionsCLR.None)
    ' Revised: adds resilient fallback constructor that handles RegexOptions
    ' via inline flags when the direct .NET constructor cannot be resolved.

    If m_host Is Nothing Then
        Debug.Print "CLRHost not initialized. Cannot create Regex object."
        Exit Sub
    End If
    If Not m_cobjRegex Is Nothing Then
        Debug.Print "Regex object already initialized. Create a new CLRRegex instance for a new pattern/options."
        Exit Sub
    End If

    m_pattern = rgxPattern
    m_options = rgxOptions

    On Error Resume Next
    If rgxOptions = RegexOptionsCLR.None Then
        ' The simple String-only constructor always works.
        Set m_cobjRegex = m_host.CLRCreateObjectWithParams("System.Text.RegularExpressions.Regex", rgxPattern)
    Else
        ' First, try the native (String, RegexOptions) overload.
        Set m_cobjRegex = m_host.CLRCreateObjectWithParams("System.Text.RegularExpressions.Regex", rgxPattern, CLng(rgxOptions))
        ' If that fails, fall back to embedding inline option flags into the pattern.
        If (err.Number <> 0) Or (m_cobjRegex Is Nothing) Then
            err.Clear
            Dim patchedPattern As String
            patchedPattern = BuildInlinePattern(rgxPattern, rgxOptions)
            Set m_cobjRegex = m_host.CLRCreateObjectWithParams("System.Text.RegularExpressions.Regex", patchedPattern)
        End If
    End If

    If err.Number <> 0 Or m_cobjRegex Is Nothing Then
        Debug.Print "Error creating .NET Regex object for pattern '" & rgxPattern & "' with options " & CLng(rgxOptions) & ": " & err.Description
        Set m_cobjRegex = Nothing
    End If
    err.Clear
    On Error GoTo 0
End Sub

Public Property Get Pattern() As String
    Pattern = m_pattern
End Property

Public Property Get Options() As Long
    Options = m_options
End Property

Public Function IsMatch(ByVal sourceText As String) As Boolean
    If m_cobjRegex Is Nothing Or m_host Is Nothing Then
        Debug.Print "Regex not initialized. Call InitializeRegex first (Pattern: " & m_pattern & ")."
        IsMatch = False
        Exit Function
    End If

    On Error Resume Next
    IsMatch = m_host.CLRInvokeMethod(m_cobjRegex, "IsMatch", sourceText)
    If err.Number <> 0 Then
        Debug.Print "Error in IsMatch for pattern '" & m_pattern & "' on text '" & left$(sourceText, 50) & "...': " & err.Description
        IsMatch = False
    End If
    err.Clear
    On Error GoTo 0
End Function

Public Function Match(ByVal sourceText As String) As CLRRegexMatch
    Dim vbaMatch As New CLRRegexMatch

    If m_cobjRegex Is Nothing Or m_host Is Nothing Then
        Debug.Print "Regex not initialized for Match method (Pattern: " & m_pattern & "). Returning uninitialized Match object."
        Call vbaMatch.Initialize(Nothing, m_host)
        Set Match = vbaMatch
        Exit Function
    End If

    Dim cobjMatch As mscorlib.Object
    On Error Resume Next
    Set cobjMatch = m_host.CLRInvokeMethod(m_cobjRegex, "Match", sourceText)
    If err.Number <> 0 Or cobjMatch Is Nothing Then
        Debug.Print "Error invoking .NET Match method for pattern '" & m_pattern & "': " & err.Description
        Call vbaMatch.Initialize(Nothing, m_host)
    Else
        Call vbaMatch.Initialize(cobjMatch, m_host)
    End If
    err.Clear
    On Error GoTo 0

    Set Match = vbaMatch
End Function

Public Function Matches(ByVal sourceText As String) As CLRRegexMatchCollection
    Dim vbaMatches As New CLRRegexMatchCollection

    If m_cobjRegex Is Nothing Or m_host Is Nothing Then
        Debug.Print "Regex not initialized for Matches method (Pattern: " & m_pattern & "). Returning empty Matches collection."
        Call vbaMatches.Initialize(Nothing, m_host)
        Set Matches = vbaMatches
        Exit Function
    End If

    Dim cobjMatchCollection As mscorlib.Object
    On Error Resume Next
    Set cobjMatchCollection = m_host.CLRInvokeMethod(m_cobjRegex, "Matches", sourceText)
    If err.Number <> 0 Or cobjMatchCollection Is Nothing Then
        Debug.Print "Error invoking .NET Matches method for pattern '" & m_pattern & "': " & err.Description
        Call vbaMatches.Initialize(Nothing, m_host)
    Else
        Call vbaMatches.Initialize(cobjMatchCollection, m_host)
    End If
    err.Clear
    On Error GoTo 0

    Set Matches = vbaMatches
End Function

Public Function ReplaceText(ByVal sourceText As String, ByVal replacementPattern As String) As String
    If m_cobjRegex Is Nothing Or m_host Is Nothing Then
        Debug.Print "Regex not initialized for ReplaceText (Pattern: " & m_pattern & ")."
        ReplaceText = sourceText
        Exit Function
    End If

    On Error Resume Next
    ReplaceText = m_host.CLRInvokeMethod(m_cobjRegex, "Replace", sourceText, replacementPattern)
    If err.Number <> 0 Then
        Debug.Print "Error in ReplaceText for pattern '" & m_pattern & "': " & err.Description
        ReplaceText = sourceText
    End If
    err.Clear
    On Error GoTo 0
End Function

Public Function SplitText(ByVal sourceText As String) As Variant
    ' Revised implementation – gracefully handles both object (string[]) and native VBA array returns
    If m_cobjRegex Is Nothing Or m_host Is Nothing Then
        Debug.Print "Regex not initialized for SplitText (Pattern: " & m_pattern & ")."
        SplitText = CVErr(xlErrValue)
        Exit Function
    End If

    On Error Resume Next
    Dim rawResult As Variant
    rawResult = m_host.CLRInvokeMethod(m_cobjRegex, "Split", sourceText)

    ' 1) If the result is a native VBA array, return it directly.
    If IsArray(rawResult) Then
        SplitText = rawResult
        err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' 2) If the result is an object (e.g., System.String[]), convert as before.
    Dim cobjNetArray As mscorlib.Object
    If IsObject(rawResult) Then Set cobjNetArray = rawResult

    If (cobjNetArray Is Nothing) Then
        Debug.Print "Error invoking .NET Split method for pattern '" & m_pattern & "' on text '" & left$(sourceText, 50) & "...': " & err.Description
        SplitText = CVErr(xlErrValue)
        err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' Existing conversion logic (unchanged)
    Dim arrayLength As Long
    arrayLength = m_host.CLRProperty(cobjNetArray, "Length")
    If err.Number <> 0 Then
        Debug.Print "Error getting array length from .NET Split result: " & err.Description
        SplitText = CVErr(xlErrValue)
        err.Clear
        On Error GoTo 0
        Exit Function
    End If
    err.Clear

    Dim vbaArray() As String
    If arrayLength = 0 Then
        SplitText = Array()
        Exit Function
    End If
    ReDim vbaArray(0 To arrayLength - 1)

    Dim tempColl As Object
    Dim element As Variant
    Set tempColl = m_host.ToEnumerable(cobjNetArray)
    If err.Number <> 0 Or tempColl Is Nothing Then
        Debug.Print "Error converting .NET array to Enumerable for Split result: " & err.Description
        SplitText = CVErr(xlErrValue)
        err.Clear
        On Error GoTo 0
        Exit Function
    End If
    err.Clear

    Dim i As Long
    For Each element In tempColl
        If i > UBound(vbaArray) Then Exit For
        vbaArray(i) = CStr(element)
        i = i + 1
    Next element

    If err.Number <> 0 Then
        Debug.Print "Error during .NET array to VBA array conversion for Split result: " & err.Description
        SplitText = CVErr(xlErrValue)
    Else
        SplitText = vbaArray
    End If
    err.Clear
    On Error GoTo 0
End Function

Private Sub Class_Terminate()
    Set m_cobjRegex = Nothing
    Set m_asmSystem = Nothing
    If Not m_host Is Nothing Then Set m_host = Nothing
End Sub

' ============================
' Helper – converts RegexOptions flags into inline option syntax
' ============================
Private Function BuildInlinePattern(ByVal basePattern As String, ByVal opts As RegexOptionsCLR) As String
    Dim flags As String
    flags = ""
    If (opts And RegexOptionsCLR.ignoreCase) <> 0 Then flags = flags & "i"
    If (opts And RegexOptionsCLR.multiLine) <> 0 Then flags = flags & "m"
    If (opts And RegexOptionsCLR.Singleline) <> 0 Then flags = flags & "s"
    ' NOTE: other options (compiled, etc.) are ignored by inline syntax because
    '       they do not affect pattern semantics in this context.

    If flags <> "" Then
        BuildInlinePattern = "(?" & flags & ")" & basePattern
    Else
        BuildInlinePattern = basePattern
    End If
End Function

