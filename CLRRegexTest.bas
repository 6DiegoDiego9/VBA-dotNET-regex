Attribute VB_Name = "CLRRegexTest"
Option Explicit

' High-resolution timer API declarations for performance testing
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
#End If


Public Sub RunCLRRegexTests()
    Dim clrRegex As clrRegex
    Dim testPattern As String
    Dim textToSearch As String
    Dim rgxOptions As RegexOptionsCLR

    Debug.Print "--- Starting CLRRegex Tests ---"

    ' =============================
    '  Test 1 – IsMatch
    ' =============================
    Debug.Print vbCrLf & "--- Test 1: IsMatch ---"
    testPattern = "\d+"
    textToSearch = "abc 123 def"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Debug.Print "IsMatch: " & clrRegex.IsMatch(textToSearch) ' --> True

    textToSearch = "abc def"
    Debug.Print "Text: '" & textToSearch & "'"
    Debug.Print "IsMatch: " & clrRegex.IsMatch(textToSearch) ' --> False
    Set clrRegex = Nothing

    ' =============================
    '  Test 2 – Match object and its properties
    ' =============================
    Debug.Print vbCrLf & "--- Test 2: Match object ---"
    testPattern = "(\w+)\s*=\s*(\d+)" ' e.g. key = 123
    textToSearch = "item1 = 100, item2 = 200"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"

    Dim firstMatch As CLRRegexMatch
    Set firstMatch = clrRegex.Match(textToSearch)
    If Not firstMatch Is Nothing And firstMatch.Success Then
        Debug.Print "First Match Success: " & firstMatch.Success
        Debug.Print "First Match Value: '" & firstMatch.Value & "'"
        Debug.Print "First Match Index: " & firstMatch.Index
        Debug.Print "First Match Length: " & firstMatch.Length

        Debug.Print "Groups Count: " & firstMatch.Groups.Count
        Dim grp As CLRRegexGroup
        For Each grp In firstMatch.Groups
            Debug.Print "  Group Value: '" & grp.Value & "', Index: " & grp.Index & ", Name: '" & grp.Name & "', Success: " & grp.Success
        Next grp
        If firstMatch.Groups.Count > 1 Then
             If firstMatch.Groups.Item(1).Success Then
                Debug.Print "  Group 1 Value (by index 1): '" & firstMatch.Groups.Item(1).Value & "'"
             Else
                Debug.Print "  Group 1 (by index 1) was not successful."
             End If
        End If
    Else
        Debug.Print "No match found or match failed for Test 2."
    End If
    Set clrRegex = Nothing
    Set firstMatch = Nothing

    ' =============================
    '  Test 3 – Matches collection & NextMatch
    ' =============================
    Debug.Print vbCrLf & "--- Test 3: Matches collection and NextMatch ---"
    testPattern = "(\w+)\s*=\s*(\d+)"
    textToSearch = "item1 = 100, item2 = 200, item3=300"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"

    Dim allMatches As CLRRegexMatchCollection
    Set allMatches = clrRegex.Matches(textToSearch)
    If Not allMatches Is Nothing Then
        Debug.Print "Matches Count: " & allMatches.Count
        Dim currentMatch As CLRRegexMatch
        For Each currentMatch In allMatches ' [1]
            If currentMatch.Success Then
                Debug.Print "  Match Value: '" & currentMatch.Value & "'"
                If currentMatch.Groups.Count > 2 Then ' groups 0,1,2
                    If currentMatch.Groups.Item(1).Success Then Debug.Print "    Group 1: '" & currentMatch.Groups.Item(1).Value & "'"
                    If currentMatch.Groups.Item(2).Success Then Debug.Print "    Group 2: '" & currentMatch.Groups.Item(2).Value & "'"
                End If
            Else
                Debug.Print "  Encountered an unsuccessful match in Matches collection."
            End If
        Next currentMatch

        Set firstMatch = clrRegex.Match(textToSearch)
        If Not firstMatch Is Nothing And firstMatch.Success Then
            Debug.Print "Testing NextMatch from first match ('" & firstMatch.Value & "'):";
            Dim nextM As CLRRegexMatch
            Set nextM = firstMatch.NextMatch()
            If Not nextM Is Nothing And nextM.Success Then
                Debug.Print "  Next Match Value: '" & nextM.Value & "'"
            Else
                Debug.Print "  No successful NextMatch found or NextMatch returned unsuccessful match."
            End If
        Else
             Debug.Print "Could not get first match for NextMatch test."
        End If
    Else
        Debug.Print "Matches collection is Nothing (should not happen, should be empty initialized)."
    End If
    Set clrRegex = Nothing: Set allMatches = Nothing: Set firstMatch = Nothing

    ' =============================
    '  Test 4 – ReplaceText (ignoreCase)
    ' =============================
    Debug.Print vbCrLf & "--- Test 4: ReplaceText ---"
    testPattern = "\bapple\b"
    textToSearch = "I have an apple. An Apple a day."
    rgxOptions = RegexOptionsCLR.ignoreCase
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern, rgxOptions)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "', Options: " & clrRegex.Options
    Debug.Print "Original Text: '" & textToSearch & "'"
    Dim replacedText As String
    replacedText = clrRegex.ReplaceText(textToSearch, "orange")
    Debug.Print "Replaced Text: '" & replacedText & "'"
    Set clrRegex = Nothing

    ' =============================
    '  Test 5 – SplitText
    ' =============================
    Debug.Print vbCrLf & "--- Test 5: SplitText ---"
    testPattern = "[,;\s]+"
    textToSearch = "one,two;three four five"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"

    Dim splitResult As Variant
    splitResult = clrRegex.SplitText(textToSearch)
    If IsArray(splitResult) Then
        Debug.Print "Split Result (Array with " & UBound(splitResult) - LBound(splitResult) + 1 & " elements):"
        Dim i As Long
        For i = LBound(splitResult) To UBound(splitResult)
            Debug.Print "  Element " & i & ": '" & splitResult(i) & "'"
        Next i
    ElseIf IsError(splitResult) Then
        Debug.Print "Split Result: Error " & CStr(splitResult)
    Else
        Debug.Print "Split Result: Not an array. TypeName: " & TypeName(splitResult)
    End If
    Set clrRegex = Nothing

    ' =============================
    '  Test 6 – RegexOptions (Compiled)
    ' =============================
    Debug.Print vbCrLf & "--- Test 6: RegexOptions (Compiled) ---"
    testPattern = "\d{4}-\d{2}-\d{2}"
    textToSearch = "Date: 2025-05-30"
    rgxOptions = RegexOptionsCLR.compiled Or RegexOptionsCLR.ignoreCase
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern, rgxOptions)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "', Options: " & clrRegex.Options
    Debug.Print "Text: '" & textToSearch & "'"
    Debug.Print "IsMatch with Compiled: " & clrRegex.IsMatch(textToSearch)
    Set clrRegex = Nothing

    ' =============================
    '  Test 7 – Named Groups
    ' =============================
    Debug.Print vbCrLf & "--- Test 7: Named Groups ---"
    testPattern = "(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})"
    textToSearch = "2025-05-30"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Set firstMatch = clrRegex.Match(textToSearch)
    If Not firstMatch Is Nothing And firstMatch.Success Then
        Debug.Print "Match: '" & firstMatch.Value & "'"
        If firstMatch.Groups.Item("year").Success Then Debug.Print "  Group 'year' (by name): '" & firstMatch.Groups.Item("year").Value & "'"
        If firstMatch.Groups.Item("month").Success Then Debug.Print "  Group 'month' (by name): '" & firstMatch.Groups.Item("month").Value & "'"
        If firstMatch.Groups.Item("day").Success Then Debug.Print "  Group 'day' (by name): '" & firstMatch.Groups.Item("day").Value & "'"
        If firstMatch.Groups.Item(3).Success Then Debug.Print "  Group 'day' (by index 3): '" & firstMatch.Groups.Item(3).Value & "'"
    Else
        Debug.Print "No match for named groups test."
    End If
    Set clrRegex = Nothing
    Set firstMatch = Nothing

    ' =============================
    '  Test 8 – Captures (repeated capturing groups)
    ' =============================
    Debug.Print vbCrLf & "--- Test 8: Captures ---"
    testPattern = "((?<word>\w+)\s*)+"
    textToSearch = "one two three"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Set firstMatch = clrRegex.Match(textToSearch)
    If Not firstMatch Is Nothing And firstMatch.Success Then
        Debug.Print "Match: '" & firstMatch.Value & "'"
        Dim wordGroup As CLRRegexGroup
        Set wordGroup = firstMatch.Groups.Item("word")
        If Not wordGroup Is Nothing And wordGroup.Success Then
            Debug.Print "  Group 'word' Value (last capture): '" & wordGroup.Value & "'"
            Debug.Print "  Group 'word' Captures Count: " & wordGroup.Captures.Count
            Dim cap As CLRRegexCapture
            For Each cap In wordGroup.Captures
                If cap.IsValid Then Debug.Print "    Capture Value: '" & cap.Value & "', Index: " & cap.Index
            Next cap
        Else
            Debug.Print "  Group 'word' not found or not successful."
        End If
    Else
        Debug.Print "No match for captures test."
    End If
    Set clrRegex = Nothing: Set firstMatch = Nothing

    ' =============================
    '  Test 9 – Empty input / no-match scenarios
    ' =============================
    Debug.Print vbCrLf & "--- Test 9: Empty/No Match Scenarios ---"
    testPattern = "xyz"
    textToSearch = ""
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Debug.Print "IsMatch (empty text): " & clrRegex.IsMatch(textToSearch)
    Set firstMatch = clrRegex.Match(textToSearch)
    Debug.Print "Match Success (empty text): " & firstMatch.Success

    textToSearch = "abc"
    Debug.Print "Text: '" & textToSearch & "'"
    Debug.Print "IsMatch (no match): " & clrRegex.IsMatch(textToSearch)
    Set firstMatch = clrRegex.Match(textToSearch)
    Debug.Print "Match Success (no match): " & firstMatch.Success

    Dim emptyMatches As CLRRegexMatchCollection
    Set emptyMatches = clrRegex.Matches(textToSearch)
    Debug.Print "Matches Count (no match): " & emptyMatches.Count
    Debug.Print "Replace (no match): '" & clrRegex.ReplaceText(textToSearch, "---") & "'"

    Dim splitNoMatch As Variant
    splitNoMatch = clrRegex.SplitText(textToSearch)
    If IsArray(splitNoMatch) Then
        If LBound(splitNoMatch) <= UBound(splitNoMatch) Then
            Debug.Print "Split (no match) Element 0: '" & splitNoMatch(LBound(splitNoMatch)) & "'"
        Else
            Debug.Print "Split (no match) returned empty array."
        End If
    End If
    Set clrRegex = Nothing: Set firstMatch = Nothing: Set emptyMatches = Nothing

    ' =============================
    '  Test 10 – Invalid pattern error handling
    ' =============================
    Debug.Print vbCrLf & "--- Test 10: Invalid Pattern ---"
    testPattern = "(?invalid"
    Set clrRegex = New clrRegex
    Debug.Print "Attempting to initialize with invalid pattern: '" & testPattern & "'"
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "IsMatch after invalid init: " & clrRegex.IsMatch("test")
    Set firstMatch = clrRegex.Match("test")
    If Not firstMatch Is Nothing And firstMatch.Success Then
         Debug.Print "Error: Regex object seems to be working with invalid pattern."
    Else
         Debug.Print "Regex object Match.Success is False after invalid init, as expected."
    End If
    Set clrRegex = Nothing: Set firstMatch = Nothing

    ' ==========================================================================================
    '  NEW TESTS – Advanced .NET-only regex features (not supported by classic VBScript RegExp)
    ' ==========================================================================================

    ' -----------------------------------------------------------------------------
    '  Test 11 – Positive Lookahead (?=...)
    '           Matches a word only when it is immediately followed by a 4-digit number.
    ' -----------------------------------------------------------------------------
    Debug.Print vbCrLf & "--- Test 11: Positive Lookahead ---"
    testPattern = "\w+(?=\s\d{4})"
    textToSearch = "Order 1234 Item 5678 Item 42"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Set allMatches = clrRegex.Matches(textToSearch)
    Debug.Print "Matches Count: " & allMatches.Count ' --> Should be 2 (Order & Item)
    For Each currentMatch In allMatches
        If currentMatch.Success Then Debug.Print "  Lookahead Match: '" & currentMatch.Value & "'"
    Next currentMatch
    Set clrRegex = Nothing: Set allMatches = Nothing

    ' -----------------------------------------------------------------------------
    '  Test 12 – Negative Lookahead (?!...)
    '           Match "foo" only when NOT followed by "bar". Should skip "foobar".
    ' -----------------------------------------------------------------------------
    Debug.Print vbCrLf & "--- Test 12: Negative Lookahead ---"
    testPattern = "foo(?!bar)"
    textToSearch = "foobar foo fooqux foobarfoo"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Set allMatches = clrRegex.Matches(textToSearch)
    Debug.Print "Matches Count: " & allMatches.Count ' --> 2 ( second "foo" & embedded "foo" in "foobarfoo")
    For Each currentMatch In allMatches
        If currentMatch.Success Then Debug.Print "  Neg-LA Match: '" & currentMatch.Value & "' at Index " & currentMatch.Index
    Next currentMatch
    Set clrRegex = Nothing: Set allMatches = Nothing

    ' -----------------------------------------------------------------------------
    '  Test 13 – Positive Lookbehind (?<=...)
    '           Capture numbers that are preceded by a $ symbol.
    ' -----------------------------------------------------------------------------
    Debug.Print vbCrLf & "--- Test 13: Positive Lookbehind ---"
    testPattern = "(?<=\$)\d+"
    textToSearch = "Prices: $100 and $200; not_a_price 300"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Set allMatches = clrRegex.Matches(textToSearch)
    Debug.Print "Matches Count: " & allMatches.Count ' --> 2 (100, 200)
    For Each currentMatch In allMatches
        If currentMatch.Success Then Debug.Print "  Lookbehind Match: '" & currentMatch.Value & "'"
    Next currentMatch
    Set clrRegex = Nothing: Set allMatches = Nothing

    ' -----------------------------------------------------------------------------
    '  Test 14 – Negative Lookbehind (?<!...)
    '           Capture numbers NOT preceded by a $ symbol.
    ' -----------------------------------------------------------------------------
    Debug.Print vbCrLf & "--- Test 14: Negative Lookbehind ---"
    testPattern = "(?<!\$)\b\d+\b"
    textToSearch = "$100 200 300$ $400 500"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Set allMatches = clrRegex.Matches(textToSearch)
    Debug.Print "Matches Count: " & allMatches.Count ' --> 2 (200 & 500)
    For Each currentMatch In allMatches
        If currentMatch.Success Then Debug.Print "  Neg-LB Match: '" & currentMatch.Value & "'"
    Next currentMatch
    Set clrRegex = Nothing: Set allMatches = Nothing

    ' -----------------------------------------------------------------------------
    '  Test 15 – Atomic Grouping (?>...)
    '           Demonstrate backtracking suppression.
    '           Pattern "(?>\d+)\d" on "1234" will fail (no backtracking);
    '           Equivalent lazy / normal grouping would succeed.
    ' -----------------------------------------------------------------------------
    Debug.Print vbCrLf & "--- Test 15: Atomic Grouping ---"
    ' (a) Without atomic group – should match
    testPattern = "(\d+)\d"
    textToSearch = "1234"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern (non-atomic): '" & testPattern & "'"
    Debug.Print "Text: '" & textToSearch & "'"
    Debug.Print "IsMatch (non-atomic, expects True): " & clrRegex.IsMatch(textToSearch)
    Set clrRegex = Nothing

    ' (b) With atomic group – should NOT match because backtracking is blocked
    testPattern = "(?>\d+)\d"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern (atomic): '" & testPattern & "'"
    Debug.Print "IsMatch (atomic, expects False): " & clrRegex.IsMatch(textToSearch)
    Set clrRegex = Nothing

    ' -----------------------------------------------------------------------------
    '  Test 16 – Balancing Groups
    '           Match correctly nested parentheses.
    '           Uses (?<Open>\() to push onto stack, (?<Close-Open>\)) to pop.
    ' -----------------------------------------------------------------------------
    Debug.Print vbCrLf & "--- Test 16: Balancing Groups ---"
    testPattern = "^[^()]*(((?'Open'\()[^()]*)+((?'Close-Open'\))[^()]*)+)*(?(Open)(?!))$"
    
    ' (a) Balanced string - should match
    textToSearch = "a(b(c)d)e"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Pattern: '" & clrRegex.Pattern & "'"
    Debug.Print "Text (balanced): '" & textToSearch & "'"
    Debug.Print "IsMatch (balanced, expects True): " & clrRegex.IsMatch(textToSearch)
    Set clrRegex = Nothing

    ' (b) Unbalanced string (too many opens) - should fail
    textToSearch = "a(b(c)de"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Text (unbalanced): '" & textToSearch & "'"
    Debug.Print "IsMatch (unbalanced, expects False): " & clrRegex.IsMatch(textToSearch)
    Set clrRegex = Nothing
    
    ' (c) Unbalanced string (too many closes) - should fail
    textToSearch = "a(b(c))d)e"
    Set clrRegex = New clrRegex
    Call clrRegex.InitializeRegex(testPattern)
    Debug.Print "Text (unbalanced): '" & textToSearch & "'"
    Debug.Print "IsMatch (unbalanced, expects False): " & clrRegex.IsMatch(textToSearch)
    Set clrRegex = Nothing

    ' ==========================================================================================
    '  Test 17 – Performance Comparison vs. VBScript.RegExp, StaticRegexSingle, and vb2net
    ' ==========================================================================================
    Debug.Print vbCrLf & "--- Test 17: Performance Comparison ---"
    
    Dim iterations As Long
    iterations = 100
    
    testPattern = "\b([a-z0-9]+)\s*=\s*([0-9.]+)\b" ' e.g., "value = 123.45"
    Dim originalTextToSearch As String
    originalTextToSearch = "val1=100, val2 = 200.5, val3= 300, val4=400. This is a moderately long string to search through for key-value pairs. val5=500, val6 = 600.9, val7=700, val8=800"
    Dim sizeMultiplier As Long
    sizeMultiplier = 100
    textToSearch = "" ' Initialize
    For i = 1 To sizeMultiplier
        textToSearch = textToSearch & originalTextToSearch & " " ' Add a space to separate repetitions
    Next i
    ' Remove the last trailing space if any
    If Len(textToSearch) > 0 Then textToSearch = left$(textToSearch, Len(textToSearch) - 1)
    
    Debug.Print "Pattern: '" & testPattern & "'"
    Debug.Print "Iterations: " & iterations
    Debug.Print "Text length: " & Len(textToSearch)
    
    Dim startTime As Currency, endTime As Currency, freq As Currency
    QueryPerformanceFrequency freq
    
    Dim clrTime As Double
    Dim vbTime As Double
    Dim staticTime As Double
    Dim vb2netTime As Double
    Dim k As Long
    Dim matchCount As Long ' To verify consistency
    
    ' --- VBScript.RegExp Test ---
    Dim vbRegEx As Object
    Dim vbMatches As Object
    Set vbRegEx = CreateObject("VBScript.RegExp")
    vbRegEx.Pattern = testPattern
    vbRegEx.Global = True
    vbRegEx.ignoreCase = True
    
    QueryPerformanceCounter startTime
    For k = 1 To iterations
        Set vbMatches = vbRegEx.execute(textToSearch)
    Next k
    QueryPerformanceCounter endTime
    
    vbTime = (endTime - startTime) / freq
    If Not vbMatches Is Nothing Then matchCount = vbMatches.Count Else matchCount = -1
    Debug.Print "VBScript.RegExp (" & matchCount & " matches) x " & iterations & " iterations: " & format$(vbTime, "0.0000") & " seconds."
    Set vbRegEx = Nothing: Set vbMatches = Nothing
    
    ' --- CLRRegex Test ---
    Dim clrMatches As CLRRegexMatchCollection
    Set clrRegex = New clrRegex
    ' Using the 'Compiled' option for a fair, high-performance comparison
    Call clrRegex.InitializeRegex(testPattern, RegexOptionsCLR.ignoreCase Or RegexOptionsCLR.compiled)
    
    QueryPerformanceCounter startTime
    For k = 1 To iterations
        Set clrMatches = clrRegex.Matches(textToSearch)
    Next k
    QueryPerformanceCounter endTime
    
    clrTime = (endTime - startTime) / freq
    If Not clrMatches Is Nothing Then matchCount = clrMatches.Count Else matchCount = -1
    Debug.Print "CLRRegex (Compiled) (" & matchCount & " matches) x " & iterations & " iterations: " & format$(clrTime, "0.0000") & " seconds."
    Set clrRegex = Nothing: Set clrMatches = Nothing
    
    ' --- vb2net Test ---
    ' Initialize vb2net (do this before the timing loop)
    Dim vb2netFile As String
    vb2netFile = "C:\temp\vb2net\vb2net.dll"
    Call InitializeVb2net(vb2netFile)

    ' Get types needed for regex
    Dim asmMscorlib As Object, typeString As Object, typeInt32 As Object
    Set asmMscorlib = LoadAssembly("mscorlib")
    Set typeString = asmMscorlib.GetType_2("System.String")
    Set typeInt32 = asmMscorlib.GetType_2("System.Int32")
    
    ' Get regex types and create regex instance with pattern and options
    Dim asmRegex As Object, typeRegex As Object, typeRegexOptions As Object, ctorRegex As Object
    Set asmRegex = LoadAssembly("System.Text.RegularExpressions")
    Set typeRegex = asmRegex.GetType_2("System.Text.RegularExpressions.Regex")
    Set typeRegexOptions = asmRegex.GetType_2("System.Text.RegularExpressions.RegexOptions")
    
    ' Get constructor that takes pattern string and RegexOptions
    Set ctorRegex = typeRegex.GetConstructor(Array(typeString, typeRegexOptions))
    
    ' Create regex with pattern and options (IgnoreCase = 1, Compiled = 8)
    Dim regex As Object
    Dim regexOptionsValue As Long
    regexOptionsValue = 1 Or 8 ' IgnoreCase | Compiled
    Set regex = ctorRegex.Invoke_3(Array(testPattern, regexOptionsValue))
    
    ' Perform the timing test
    QueryPerformanceCounter startTime
    For k = 1 To iterations
        Dim vb2netMatches As Object
        ' Try different suffixes for Matches method
        On Error Resume Next
        Set vb2netMatches = regex.Matches_2(textToSearch)
        If err.Number <> 0 Then
            err.Clear
            Set vb2netMatches = regex.Matches_3(textToSearch)
        End If
        If err.Number <> 0 Then
            err.Clear
            Set vb2netMatches = regex.Matches_4(textToSearch)
        End If
        If err.Number <> 0 Then
            err.Clear
            Set vb2netMatches = regex.Matches_5(textToSearch)
        End If
        On Error GoTo 0
    Next k
    QueryPerformanceCounter endTime
    
    vb2netTime = (endTime - startTime) / freq
    If Not vb2netMatches Is Nothing Then
        matchCount = vb2netMatches.Count
    Else
        matchCount = -1
    End If
    Debug.Print "vb2net (Compiled) (" & matchCount & " matches) x " & iterations & " iterations: " & format$(vb2netTime, "0.0000") & " seconds."
    Set regex = Nothing
    
    ' --- StaticRegexSingle (VBA-Native) Test ---
    Dim staticRegexObj As StaticRegexSingle.RegexTy
    Dim staticMatcher As StaticRegexSingle.MatcherStateTy
    Dim staticMatchCount As Long
    
    ' Initialize Regex (outside timing loop)
    Call StaticRegexSingle.InitializeRegex(staticRegexObj, testPattern, True) ' True for ignoreCase
    
    QueryPerformanceCounter startTime
    For k = 1 To iterations
        staticMatchCount = 0
        ' Initialize the matcher state for a global search
        Call StaticRegexSingle.InitializeMatcherState(staticMatcher, localMatch:=False, multiLine:=False)
        
        ' Loop through all matches using the iterator pattern
        Do While StaticRegexSingle.MatchNext(staticMatcher, staticRegexObj, textToSearch)
            staticMatchCount = staticMatchCount + 1
        Loop
    Next k
    QueryPerformanceCounter endTime
    
    staticTime = (endTime - startTime) / freq
    Debug.Print "StaticRegexSingle (" & staticMatchCount & " matches) x " & iterations & " iterations: " & format$(staticTime, "0.0000") & " seconds."
    
    ' --- Performance Summary ---
    Debug.Print ""
    Debug.Print "--- Performance Summary for " & iterations & " iterations ---"
    Debug.Print "VBScript.RegExp time:     " & format$(vbTime, "0.0000") & " s"
    Debug.Print "CLRRegex (Compiled) time: " & format$(clrTime, "0.0000") & " s"
    If vb2netTime >= 0 Then
        Debug.Print "vb2net (Compiled) time:   " & format$(vb2netTime, "0.0000") & " s"
    End If
    Debug.Print "StaticRegexSingle time:   " & format$(staticTime, "0.0000") & " s"
    
    Dim minTime As Double
    minTime = vbTime
    Dim fastestEngine As String
    fastestEngine = "VBScript.RegExp"
    
    If clrTime < minTime Then
        minTime = clrTime
        fastestEngine = "CLRRegex (Compiled)"
    End If
    If vb2netTime >= 0 And vb2netTime < minTime Then
        minTime = vb2netTime
        fastestEngine = "vb2net (Compiled)"
    End If
    If staticTime < minTime Then
        minTime = staticTime
        fastestEngine = "StaticRegexSingle"
    End If
    
    Debug.Print "Fastest engine: " & fastestEngine & " at " & format$(minTime, "0.0000") & " s"
    If minTime > 0 Then
        Debug.Print "Relative Speeds (lower is better, 1.00x is fastest):"
        Debug.Print "  VBScript.RegExp: " & format$(vbTime / minTime, "0.00") & "x"
        Debug.Print "  CLRRegex: " & format$(clrTime / minTime, "0.00") & "x"
        If vb2netTime >= 0 Then
            Debug.Print "  vb2net: " & format$(vb2netTime / minTime, "0.00") & "x"
        End If
        Debug.Print "  StaticRegexSingle: " & format$(staticTime / minTime, "0.00") & "x"
    End If
End Sub


Public Sub SimpleMatchExample()
    Dim rgx As New clrRegex
    Dim textToSearch As String
    Dim allMatches As CLRRegexMatchCollection
    Dim aMatch As CLRRegexMatch

    ' The .NET regex supports named groups 'key' and 'value'
    Call rgx.InitializeRegex("(?<key>\w+)\s*=\s*(?<value>\d+)")

    textToSearch = "item1 = 100, item2 = 200, invalid_item, item3 = 300"

    Set allMatches = rgx.Matches(textToSearch)

    Debug.Print "Found " & allMatches.Count & " matches."

    For Each aMatch In allMatches
        Debug.Print "--- Match Found ---"
        Debug.Print "Full Match: " & aMatch.Value
        Debug.Print "Key: " & aMatch.Groups.Item("key").Value
        Debug.Print "Value: " & aMatch.Groups.Item("value").Value
    Next aMatch
End Sub
