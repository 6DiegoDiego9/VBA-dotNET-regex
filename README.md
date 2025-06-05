# VBA-dotNET-regex

A powerful, high-performance, zero dependencies VBA library that leverages .NET's System.Text.RegularExpressions through CLR integration.

This library aims to provide a complete wrapper around the System.Text.RegularExpressions namespace, bringing modern regex capabilities directly into Excel, Access, and other VBA environments, providing VBA with advanced regex features not available in VBScript.RegExp.

## Key Features

*   **Full .NET Regex Engine:** Access the complete, powerful regex engine from the .NET Framework.
*   **Modern Syntax Support:**
    *   Positive & Negative Lookaheads (`(?=...)`, `(?!...)`)
    *   Positive & Negative Lookbehinds (`(?<=...)`, `(?<!...)`)
    *   Named Capture Groups (`(?<name>...)`)
    *   Atomic Grouping (`(?>...)`)
    *   Balancing Groups (`(?<Open-Close>...)`)
    *   And much more!
*   **High Performance:** Includes support for the `RegexOptions.Compiled` flag for maximum speed in repetitive tasks.
*   **Complete Object Model:** Provides VBA-native wrapper classes for `Match`, `Group`, `Capture` and their collections, making the API intuitive to use.
*   **Graceful Cleanup:** The hosted CLR is automatically unloaded when the application exits, preventing memory leaks.

## Requirements

*   Windows OS
*   Microsoft Office 32-bit or 64-bit
*   .NET Framework 4.0 or later (most Windows systems have this pre-installed)
*   vb2clr (https://github.com/jet2jet/vb2clr)

## Installation

**Manual Method**
   - In your VBA project, import all the `.cls` and `.bas` files from the root directory of this repository.
   - import the files CLRHost.cls and ExitHandler.bas from the vb2clr library (https://github.com/jet2jet/vb2clr)

## Quick Start

Here's a simple example of how to find all key-value pairs in a string.

```vba
Public Sub SimpleMatchExample()
    Dim rgx As New CLRRegex
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
```

## API Overview

- **`CLRRegex`**: The main class.
  - `.InitializeRegex(Pattern, [Options])`: Creates the regex object.
  - `.IsMatch(Text)`: Returns `True` or `False`.
  - `.Match(Text)`: Returns a single `CLRRegexMatch` object.
  - `.Matches(Text)`: Returns a `CLRRegexMatchCollection`.
  - `.ReplaceText(Text, Replacement)`: Replaces all matches.
  - `.SplitText(Text)`: Splits the text by the pattern.

- **`CLRRegexMatch`**: Represents a single match.
  - `.Success`, `.Value`, `.Index`, `.Length`
  - `.Groups`: A collection of `CLRRegexGroup` objects.
  - `.NextMatch()`: Finds the next match in the string.

- **`CLRRegexGroup`**: Represents a captured group.
  - `.Value`, `.Index`, `.Length`, `.Name`
  - `.Captures`: A collection of all captures made by this group.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
