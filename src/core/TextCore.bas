Attribute VB_Name = "TextCore"
'==============================================================================
' MODULE: TextCore
'------------------------------------------------------------------------------
' PURPOSE
'   Implement the library profile's text and path utilities behind the
'   supported TextFacade boundary.
'
' PUBLIC SURFACE
'   None outside this VBA project. The Public procedures and constants below
'   serve TextFacade and focused in-project tests; Option Private Module keeps
'   them off the supported external surface.
'
' DEPENDENCIES
'   VBA runtime only. This core never depends on the facade, tests, examples,
'   workbook objects, the file system, or optional references.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on explicit arguments and the compile-time
'   platform; nothing reads the Excel object model or the file system.
'
' ERROR POLICY
'   Own the single invalid-argument error number. Every rejected argument
'   raises it with a description that names the argument and the rule.
'   TextFacade exposes the same value and normalizes the public error source.
'
' WORKSHEET SAFETY
'   Performs no Excel object-model access and changes no caller-owned state.
'   Path procedures manipulate strings only; they never touch the disk.
'
' TEST SEAM
'   TextTests verifies behavior through TextFacade. Direct core access remains
'   available inside the project without widening the supported API.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS. Conditional compilation selects the path
'   separator, the separator set, rooted-path forms and file-name rules that
'   differ between the two platforms.
'
' UPDATED
'   2026-09-30
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; preserve the configured component visibility.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Own the internal error code and the padding modes used by the facade.
        Public Const ERR_INVALID_ARGUMENT   As Long = vbObjectError + 2100    'Core-owned invalid-argument code
        Public Const PAD_RIGHT              As Long = 0                       'Append fill characters after the text
        Public Const PAD_LEFT               As Long = 1                       'Insert fill characters before the text
        Public Const PAD_BOTH               As Long = 2                       'Center; the extra character goes right

    'Select the separator this platform writes when composing paths.
#If Mac Then
        Public Const PATH_SEPARATOR         As String = "/"                   'POSIX separator on macOS
#Else
        Public Const PATH_SEPARATOR         As String = "\"                   'Canonical Windows separator
#End If


'
'------------------------------------------------------------------------------
'
'                                TEXT UTILITIES
'
'------------------------------------------------------------------------------
'

Public Function PadText( _
    ByVal value As String, _
    ByVal width As Long, _
    ByVal side As Long, _
    ByVal fill As String) _
    As String
'
'==============================================================================
'                                   PadText
'------------------------------------------------------------------------------
' PURPOSE
'   Widen text to a minimum length with one repeated fill character.
'
' INPUTS
'   value: text to widen; never truncated.
'   width: minimum result length in characters; zero or greater.
'   side:  PAD_RIGHT, PAD_LEFT or PAD_BOTH.
'   fill:  exactly one character.
'
' RETURNS
'   value unchanged when it already has width characters or more; otherwise
'   value padded to exactly width characters. PAD_BOTH puts the odd extra
'   character on the right.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for a negative width, an unknown side, or a
'   fill that is not exactly one character.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Split the missing length between the two sides before building text.
    Dim leftCount    As Long    'Fill characters placed before the text
    Dim missing      As Long    'Characters needed to reach the width

'------------------------------------------------------------------------------
' VALIDATE ARGUMENTS
'------------------------------------------------------------------------------
    'Reject every invalid argument before any text is built, so a caller
    'never receives a partially padded result.
        If width < 0 Then
            RaiseInvalid "TextCore.PadText", "width must be zero or greater."
        End If
        If side < PAD_RIGHT Or side > PAD_BOTH Then
            RaiseInvalid "TextCore.PadText", "side must be TextPadRight, TextPadLeft or TextPadBoth."
        End If
        If Len(fill) <> 1 Then
            RaiseInvalid "TextCore.PadText", "fill must be exactly one character."
        End If

'------------------------------------------------------------------------------
' BUILD RESULT
'------------------------------------------------------------------------------
    'Leave text that already meets the width untouched; padding never
    'shortens a value.
        missing = width - Len(value)
        If missing <= 0 Then
            PadText = value
            Exit Function
        End If

    'Place the fill according to the side; centering rounds the left share
    'down so the odd character lands on the right.
        Select Case side
            Case PAD_LEFT
                leftCount = missing
            Case PAD_BOTH
                leftCount = missing \ 2
            Case Else
                leftCount = 0
        End Select
        PadText = String$(leftCount, fill) & value & String$(missing - leftCount, fill)

End Function


Public Function CollapseWhitespace( _
    ByVal value As String) _
    As String
'
'==============================================================================
'                              CollapseWhitespace
'------------------------------------------------------------------------------
' PURPOSE
'   Trim whitespace at both ends and reduce each inner run to one space.
'
' INPUTS
'   value: text to normalize. Space, tab, line feed, vertical tab, form feed,
'   carriage return and the no-break space (U+00A0) count as whitespace.
'
' RETURNS
'   Normalized text; an all-whitespace value becomes an empty string.
'
' ERROR POLICY
'   Raises no errors.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Scan once, remembering whether a separator is owed before the next word.
    Dim character      As String     'Current character under inspection
    Dim index          As Long       'One-based scan position
    Dim pendingSpace   As Boolean    'True after whitespace that follows a word
    Dim result         As String     'Normalized text built so far

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Emit one space only between words: leading whitespace never sets the
    'pending flag, and trailing whitespace leaves it unused.
        For index = 1 To Len(value)
            character = Mid$(value, index, 1)
            If IsWhitespaceCharacter(character) Then
                pendingSpace = (Len(result) > 0)
            Else
                If pendingSpace Then
                    result = result & " "
                    pendingSpace = False
                End If
                result = result & character
            End If
        Next index
        CollapseWhitespace = result

End Function


Public Function StartsWithText( _
    ByVal value As String, _
    ByVal prefix As String, _
    ByVal compareMode As Long) _
    As Boolean
'
'==============================================================================
'                                StartsWithText
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether value begins with prefix under a stated comparison.
'
' INPUTS
'   value, prefix: text to compare; an empty prefix always matches.
'   compareMode: vbBinaryCompare (exact) or vbTextCompare (case-insensitive).
'
' RETURNS
'   True when the first Len(prefix) characters of value equal prefix.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for any other comparison mode.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Validate the mode first; a prefix longer than the value cannot match.
        ValidateCompare "TextCore.StartsWithText", compareMode
        If Len(prefix) > Len(value) Then
            StartsWithText = False
        Else
            StartsWithText = (StrComp(Left$(value, Len(prefix)), prefix, compareMode) = 0)
        End If

End Function


Public Function EndsWithText( _
    ByVal value As String, _
    ByVal suffix As String, _
    ByVal compareMode As Long) _
    As Boolean
'
'==============================================================================
'                                 EndsWithText
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether value ends with suffix under a stated comparison.
'
' INPUTS
'   value, suffix: text to compare; an empty suffix always matches.
'   compareMode: vbBinaryCompare (exact) or vbTextCompare (case-insensitive).
'
' RETURNS
'   True when the last Len(suffix) characters of value equal suffix.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for any other comparison mode.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Validate the mode first; a suffix longer than the value cannot match.
        ValidateCompare "TextCore.EndsWithText", compareMode
        If Len(suffix) > Len(value) Then
            EndsWithText = False
        Else
            EndsWithText = (StrComp(Right$(value, Len(suffix)), suffix, compareMode) = 0)
        End If

End Function


Public Function CountText( _
    ByVal value As String, _
    ByVal find As String, _
    ByVal compareMode As Long) _
    As Long
'
'==============================================================================
'                                  CountText
'------------------------------------------------------------------------------
' PURPOSE
'   Count non-overlapping occurrences of find in value.
'
' INPUTS
'   value: text to search.
'   find: nonempty text to count.
'   compareMode: vbBinaryCompare (exact) or vbTextCompare (case-insensitive).
'
' RETURNS
'   Occurrences found scanning left to right; each match resumes the search
'   after its last character, so "aaaa" contains "aa" twice.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for an empty find or an unknown mode.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Track the next search start and the matches found so far.
    Dim found      As Long    'Position of the latest match; 0 ends the scan
    Dim matches    As Long    'Non-overlapping occurrences counted
    Dim start      As Long    'One-based position where the next search begins

'------------------------------------------------------------------------------
' VALIDATE ARGUMENTS
'------------------------------------------------------------------------------
    'An empty needle would match everywhere and never advance the scan.
        ValidateCompare "TextCore.CountText", compareMode
        If Len(find) = 0 Then
            RaiseInvalid "TextCore.CountText", "find must not be empty."
        End If

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Resume after each whole match so occurrences never overlap.
        start = 1
        Do
            found = InStr(start, value, find, compareMode)
            If found = 0 Then Exit Do
            matches = matches + 1
            start = found + Len(find)
        Loop
        CountText = matches

End Function


Public Function SplitText( _
    ByVal value As String, _
    ByVal delimiter As String, _
    ByVal trimItems As Boolean, _
    ByVal removeEmpty As Boolean) _
    As String()
'
'==============================================================================
'                                  SplitText
'------------------------------------------------------------------------------
' PURPOSE
'   Split text on a delimiter with optional trimming and empty-item removal.
'
' INPUTS
'   value: text to split; an empty value yields no items.
'   delimiter: nonempty literal separator, matched exactly.
'   trimItems: trim whitespace at both ends of each item.
'   removeEmpty: drop items that are empty after optional trimming.
'
' RETURNS
'   A zero-based String array. No items gives LBound 0 and UBound -1.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for an empty delimiter.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Filter the raw pieces into a second array of the same maximum size.
    Dim index      As Long       'Position in the raw pieces
    Dim item       As String     'Current piece after optional trimming
    Dim kept       As Long       'Items retained so far
    Dim pieces()   As String     'Raw pieces from the VBA Split function
    Dim result()   As String     'Retained items, resized at the end

'------------------------------------------------------------------------------
' VALIDATE AND SPLIT
'------------------------------------------------------------------------------
    'Split never returns an item for an empty value, so an empty input and a
    'fully filtered input share the same empty-array shape.
        If Len(delimiter) = 0 Then
            RaiseInvalid "TextCore.SplitText", "delimiter must not be empty."
        End If
        pieces = Split(value, delimiter)
        If UBound(pieces) < 0 Then
            SplitText = pieces
            Exit Function
        End If

'------------------------------------------------------------------------------
' FILTER ITEMS
'------------------------------------------------------------------------------
    'Keep the original order; trimming happens before the emptiness test so
    'whitespace-only items can be removed.
        ReDim result(0 To UBound(pieces))
        For index = 0 To UBound(pieces)
            item = pieces(index)
            If trimItems Then item = TrimWhitespace(item)
            If Not (removeEmpty And Len(item) = 0) Then
                result(kept) = item
                kept = kept + 1
            End If
        Next index

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the canonical empty array when every item was removed; otherwise
    'shrink the buffer to the retained items.
        If kept = 0 Then
            SplitText = Split(vbNullString, delimiter)
        Else
            ReDim Preserve result(0 To kept - 1)
            SplitText = result
        End If

End Function


Public Function JoinText( _
    ByVal items As Variant, _
    ByVal delimiter As String, _
    ByVal skipEmpty As Boolean) _
    As String
'
'==============================================================================
'                                   JoinText
'------------------------------------------------------------------------------
' PURPOSE
'   Join the scalar items of a one-dimensional or single-row/column array.
'
' INPUTS
'   items: a one-dimensional array with any bounds, or a two-dimensional
'          array with exactly one row or one column, such as Range.Value.
'          An unallocated dynamic array has no items.
'   delimiter: text placed between items; may be empty.
'   skipEmpty: omit items whose text is empty.
'
' RETURNS
'   Items in array order. Empty and Null become empty text; other scalars use
'   CStr, so numbers and dates follow the host's regional settings.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT when items is not an array, has more than one
'   row and column, has more than two dimensions, or contains an object,
'   nested array or error value.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Resolve the shape once, then visit every element in reading order.
    Dim column     As Long       'Column index in a two-dimensional array
    Dim hasItem    As Boolean    'True after the first retained item
    Dim rank       As Long       'Number of array dimensions
    Dim result     As String     'Joined text built so far
    Dim row        As Long       'Row index in a two-dimensional array
    Dim itemValue  As String     'Current item converted to text

'------------------------------------------------------------------------------
' VALIDATE SHAPE
'------------------------------------------------------------------------------
    'Accept only shapes with an unambiguous reading order.
        If Not IsArray(items) Then
            RaiseInvalid "TextCore.JoinText", "items must be an array."
        End If
        rank = ArrayRank(items)
        If rank > 2 Then
            RaiseInvalid "TextCore.JoinText", "items must have one or two dimensions."
        End If
        If rank = 2 Then
            If UBound(items, 1) > LBound(items, 1) And UBound(items, 2) > LBound(items, 2) Then
                RaiseInvalid "TextCore.JoinText", "a two-dimensional items array must have one row or one column."
            End If
        End If

'------------------------------------------------------------------------------
' JOIN ITEMS
'------------------------------------------------------------------------------
    'Visit rows outermost so a single column reads top to bottom and a single
    'row reads left to right; the delimiter precedes every later item.
        Select Case rank
            Case 1
                For row = LBound(items) To UBound(items)
                    itemValue = ItemText(items(row))
                    If Not (skipEmpty And Len(itemValue) = 0) Then
                        If hasItem Then result = result & delimiter
                        result = result & itemValue
                        hasItem = True
                    End If
                Next row
            Case 2
                For row = LBound(items, 1) To UBound(items, 1)
                    For column = LBound(items, 2) To UBound(items, 2)
                        itemValue = ItemText(items(row, column))
                        If Not (skipEmpty And Len(itemValue) = 0) Then
                            If hasItem Then result = result & delimiter
                            result = result & itemValue
                            hasItem = True
                        End If
                    Next column
                Next row
        End Select
        JoinText = result

End Function


Public Function IsBlankValue( _
    ByVal value As Variant) _
    As Boolean
'
'==============================================================================
'                                 IsBlankValue
'------------------------------------------------------------------------------
' PURPOSE
'   Classify a cell-style value as blank.
'
' INPUTS
'   value: any Variant.
'
' RETURNS
'   True for Empty, Null, and text that is empty or whitespace only. False
'   for every other value, including zero, False, error values, objects and
'   arrays.
'
' ERROR POLICY
'   Raises no errors.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Test objects and arrays before VarType so neither is converted to text.
        If IsObject(value) Or IsArray(value) Then
            IsBlankValue = False
        ElseIf IsEmpty(value) Or IsNull(value) Then
            IsBlankValue = True
        ElseIf VarType(value) = vbString Then
            IsBlankValue = (Len(TrimWhitespace(value)) = 0)
        Else
            IsBlankValue = False
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                PATH UTILITIES
'
'------------------------------------------------------------------------------
'

Public Function JoinPathText( _
    ByVal basePath As String, _
    ByVal childPath As String) _
    As String
'
'==============================================================================
'                                 JoinPathText
'------------------------------------------------------------------------------
' PURPOSE
'   Append a relative path to a base path with exactly one separator.
'
' INPUTS
'   basePath: any path text; trailing separators are replaced by one.
'   childPath: a relative path; empty returns basePath unchanged.
'
' RETURNS
'   basePath, one PATH_SEPARATOR and childPath. An empty basePath returns
'   childPath; a root such as "C:\" or "/" keeps its separator.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT when childPath is rooted: it starts with a
'   separator, or on Windows it names a drive such as "D:".
'
' SIDE EFFECTS
'   None. The path is text; the file system is never consulted.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Work on a copy of the base so trailing separators can be removed.
    Dim trimmedBase   As String    'Base path without trailing separators

'------------------------------------------------------------------------------
' VALIDATE AND SHORTCUT
'------------------------------------------------------------------------------
    'A rooted child would silently discard the base on most platforms, so it
    'is rejected; empty inputs return the other side unchanged.
        If IsRootedPath(childPath) Then
            RaiseInvalid "TextCore.JoinPathText", "childPath must be relative."
        End If
        If Len(childPath) = 0 Then
            JoinPathText = basePath
            Exit Function
        End If
        If Len(basePath) = 0 Then
            JoinPathText = childPath
            Exit Function
        End If

'------------------------------------------------------------------------------
' JOIN
'------------------------------------------------------------------------------
    'Remove trailing separators, then add exactly one. A root such as "/"
    'becomes empty here and regains its single separator.
        trimmedBase = basePath
        Do While Len(trimmedBase) > 0
            If Not IsPathSeparator(Right$(trimmedBase, 1)) Then Exit Do
            trimmedBase = Left$(trimmedBase, Len(trimmedBase) - 1)
        Loop
        JoinPathText = trimmedBase & PATH_SEPARATOR & childPath

End Function


Public Sub SplitPathText( _
    ByVal path As String, _
    ByRef directory As String, _
    ByRef baseName As String, _
    ByRef extension As String)
'
'==============================================================================
'                                SplitPathText
'------------------------------------------------------------------------------
' PURPOSE
'   Separate a path into directory, base name and extension.
'
' INPUTS
'   path: any path text.
'
' RETURNS
'   directory: everything up to and including the last separator, or empty.
'   baseName:  the file name without its extension.
'   extension: text after the last dot of the file name, without the dot;
'              empty for no dot, a trailing dot, or a leading-dot name such
'              as ".gitignore".
'   The parts always satisfy directory & baseName & "." & extension = path
'   when extension is nonempty, and directory & baseName = path otherwise.
'
' ERROR POLICY
'   Raises no errors.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Locate the last separator, then the last dot inside the file name.
    Dim dot        As Long      'Position of the last dot in the file name
    Dim fileName   As String    'Text after the directory part
    Dim index      As Long      'Backward scan position

'------------------------------------------------------------------------------
' SPLIT DIRECTORY
'------------------------------------------------------------------------------
    'Scan from the end so both Windows separators are recognized; the
    'directory keeps its separator so the parts rebuild the original text.
        For index = Len(path) To 1 Step -1
            If IsPathSeparator(Mid$(path, index, 1)) Then Exit For
        Next index
        directory = Left$(path, index)
        fileName = Mid$(path, index + 1)

'------------------------------------------------------------------------------
' SPLIT EXTENSION
'------------------------------------------------------------------------------
    'A dot at the first or last position does not start an extension, so
    'dot-files and names ending in a dot keep their whole text as the base.
        dot = InStrRev(fileName, ".")
        If dot > 1 And dot < Len(fileName) Then
            baseName = Left$(fileName, dot - 1)
            extension = Mid$(fileName, dot + 1)
        Else
            baseName = fileName
            extension = vbNullString
        End If

End Sub


Public Function SafeFileNameText( _
    ByVal value As String, _
    ByVal replacement As String) _
    As String
'
'==============================================================================
'                               SafeFileNameText
'------------------------------------------------------------------------------
' PURPOSE
'   Turn arbitrary text into a file name this platform accepts.
'
' INPUTS
'   value: proposed file name, without a directory.
'   replacement: empty to delete invalid characters, or one valid character.
'
' RETURNS
'   value with each invalid character replaced. On Windows the characters
'   < > : " / \ | ? * and control characters are invalid, trailing spaces and
'   dots are removed, and a reserved device name such as CON or LPT1 (alone
'   or before an extension) gains a leading underscore. On macOS ":" and "/"
'   and the null character are invalid.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT when replacement is longer than one character
'   or is itself invalid, or when no usable character remains.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Build the cleaned name character by character.
    Dim character   As String    'Current character under inspection
    Dim index       As Long      'One-based scan position
    Dim result      As String    'Cleaned name built so far

'------------------------------------------------------------------------------
' VALIDATE REPLACEMENT
'------------------------------------------------------------------------------
    'An invalid replacement would reintroduce the characters being removed.
        If Len(replacement) > 1 Then
            RaiseInvalid "TextCore.SafeFileNameText", "replacement must be empty or one character."
        End If
        If Len(replacement) = 1 Then
            If IsInvalidFileNameCharacter(replacement) Then
                RaiseInvalid "TextCore.SafeFileNameText", "replacement must be a valid file-name character."
            End If
        End If

'------------------------------------------------------------------------------
' REPLACE INVALID CHARACTERS
'------------------------------------------------------------------------------
    'Apply the platform's character rule to every position.
        For index = 1 To Len(value)
            character = Mid$(value, index, 1)
            If IsInvalidFileNameCharacter(character) Then
                result = result & replacement
            Else
                result = result & character
            End If
        Next index

'------------------------------------------------------------------------------
' APPLY WINDOWS NAME RULES
'------------------------------------------------------------------------------
#If Not Mac Then
    'Windows silently drops trailing spaces and dots and reserves device
    'names, so both are repaired here rather than failing at save time.
        Do While Len(result) > 0
            If Right$(result, 1) <> " " And Right$(result, 1) <> "." Then Exit Do
            result = Left$(result, Len(result) - 1)
        Loop
        If IsReservedDeviceName(result) Then result = "_" & result
#End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Refuse a name with no usable characters instead of returning an empty
    'string that a caller could append to a directory by mistake.
        If Len(result) = 0 Then
            RaiseInvalid "TextCore.SafeFileNameText", "value has no usable file-name characters."
        End If
        SafeFileNameText = result

End Function


'
'------------------------------------------------------------------------------
'
'                               PRIVATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Function IsWhitespaceCharacter( _
    ByVal character As String) _
    As Boolean
'
'==============================================================================
'                            IsWhitespaceCharacter
'------------------------------------------------------------------------------
' PURPOSE
'   Recognize the whitespace set shared by trimming and collapsing.
'
' RETURNS
'   True for space, tab, line feed, vertical tab, form feed, carriage return
'   and the no-break space U+00A0.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Use code points so the no-break space is matched independently of the
    'module's code page; the mask turns AscW's signed result into 0-65535.
        Select Case AscW(character) And &HFFFF&
            Case 9 To 13, 32, 160
                IsWhitespaceCharacter = True
            Case Else
                IsWhitespaceCharacter = False
        End Select

End Function


Private Function TrimWhitespace( _
    ByVal value As String) _
    As String
'
'==============================================================================
'                                TrimWhitespace
'------------------------------------------------------------------------------
' PURPOSE
'   Remove leading and trailing whitespace, including tabs and U+00A0.
'
' RETURNS
'   value without whitespace at either end; inner text is unchanged.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Move the two bounds inward until each rests on a non-whitespace character.
    Dim first   As Long    'First retained position
    Dim last    As Long    'Last retained position

'------------------------------------------------------------------------------
' TRIM
'------------------------------------------------------------------------------
    'The bounds cross for an all-whitespace value, which yields empty text.
        first = 1
        last = Len(value)
        Do While first <= last
            If Not IsWhitespaceCharacter(Mid$(value, first, 1)) Then Exit Do
            first = first + 1
        Loop
        Do While last >= first
            If Not IsWhitespaceCharacter(Mid$(value, last, 1)) Then Exit Do
            last = last - 1
        Loop
        TrimWhitespace = Mid$(value, first, last - first + 1)

End Function


Private Function IsPathSeparator( _
    ByVal character As String) _
    As Boolean
'
'==============================================================================
'                               IsPathSeparator
'------------------------------------------------------------------------------
' PURPOSE
'   Recognize the separators this platform accepts in a path.
'
' RETURNS
'   On Windows, True for both "\" and "/"; on macOS, True for "/" only.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Windows APIs accept either slash, while a backslash is an ordinary
    'file-name character on macOS.
#If Mac Then
        IsPathSeparator = (character = "/")
#Else
        IsPathSeparator = (character = "\" Or character = "/")
#End If

End Function


Private Function IsRootedPath( _
    ByVal path As String) _
    As Boolean
'
'==============================================================================
'                                 IsRootedPath
'------------------------------------------------------------------------------
' PURPOSE
'   Recognize a path that does not depend on a base directory.
'
' RETURNS
'   True when path starts with a separator (including a Windows UNC prefix)
'   or, on Windows, starts with a drive letter and a colon.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'A leading separator roots a path on both platforms; only Windows has
    'drive letters.
        If Len(path) = 0 Then
            IsRootedPath = False
        ElseIf IsPathSeparator(Left$(path, 1)) Then
            IsRootedPath = True
        Else
#If Mac Then
            IsRootedPath = False
#Else
            IsRootedPath = (Mid$(path, 2, 1) = ":" And UCase$(Left$(path, 1)) Like "[A-Z]")
#End If
        End If

End Function


Private Function IsInvalidFileNameCharacter( _
    ByVal character As String) _
    As Boolean
'
'==============================================================================
'                          IsInvalidFileNameCharacter
'------------------------------------------------------------------------------
' PURPOSE
'   Apply this platform's rule for characters a file name may not contain.
'
' RETURNS
'   On Windows, True for < > : " / \ | ? * and code points below 32; on macOS,
'   True for ":" , "/" and the null character.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Keep each platform's documented set in one place so the replacement
    'check and the character scan agree. AscW is signed, so the mask keeps
    'code points from U+8000 upward out of the control range.
#If Mac Then
        IsInvalidFileNameCharacter = (character = ":" Or character = "/" Or AscW(character) = 0)
#Else
        IsInvalidFileNameCharacter = ((AscW(character) And &HFFFF&) < 32 Or InStr(1, "<>:""/\|?*", character, vbBinaryCompare) > 0)
#End If

End Function


Private Function IsReservedDeviceName( _
    ByVal fileName As String) _
    As Boolean
'
'==============================================================================
'                             IsReservedDeviceName
'------------------------------------------------------------------------------
' PURPOSE
'   Recognize a Windows device name, alone or before an extension.
'
' RETURNS
'   True when the text before the first dot is CON, PRN, AUX, NUL, COM1 to
'   COM9 or LPT1 to LPT9, compared without regard to case.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Compare only the stem; "con.txt" is as reserved as "CON".
    Dim dot    As Long      'Position of the first dot, or 0
    Dim stem   As String    'Upper-case text before the first dot

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Numbered device names accept only the digits 1 to 9.
        dot = InStr(1, fileName, ".", vbBinaryCompare)
        If dot > 0 Then
            stem = UCase$(Left$(fileName, dot - 1))
        Else
            stem = UCase$(fileName)
        End If
        Select Case stem
            Case "CON", "PRN", "AUX", "NUL"
                IsReservedDeviceName = True
            Case Else
                IsReservedDeviceName = (stem Like "COM[1-9]" Or stem Like "LPT[1-9]")
        End Select

End Function


Private Function ArrayRank( _
    ByRef values As Variant) _
    As Long
'
'==============================================================================
'                                  ArrayRank
'------------------------------------------------------------------------------
' PURPOSE
'   Count the dimensions of an array; VBA has no built-in for this.
'
' RETURNS
'   The number of dimensions, or 0 for an unallocated dynamic array.
'
' ERROR POLICY
'   Probing one dimension past the last raises error 9, which ends the count;
'   the handler is local and the error does not escape.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Probe successive dimensions until UBound fails.
    Dim bound       As Long    'Discarded upper bound of the probed dimension
    Dim dimension   As Long    'Dimension currently being probed

'------------------------------------------------------------------------------
' PROBE DIMENSIONS
'------------------------------------------------------------------------------
    'VBA arrays have at most 60 dimensions, so the loop always terminates.
        On Error GoTo RankFound
        For dimension = 1 To 60
            bound = UBound(values, dimension)
        Next dimension
        ArrayRank = 60
        Exit Function

'------------------------------------------------------------------------------
' HANDLE LAST DIMENSION
'------------------------------------------------------------------------------
RankFound:
    'The failing probe is one past the last dimension; leaving the procedure
    'from the handler clears the error, so it never reaches the caller.
        ArrayRank = dimension - 1

End Function


Private Function ItemText( _
    ByRef item As Variant) _
    As String
'
'==============================================================================
'                                   ItemText
'------------------------------------------------------------------------------
' PURPOSE
'   Convert one array element to text for JoinText.
'
' RETURNS
'   Empty text for Empty and Null; the element's CStr text otherwise.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for an object, nested array or error value.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Reject values with no faithful scalar text before converting.
        If IsObject(item) Or IsArray(item) Or IsError(item) Then
            RaiseInvalid "TextCore.JoinText", "items must contain only scalar values."
        ElseIf IsEmpty(item) Or IsNull(item) Then
            ItemText = vbNullString
        Else
            ItemText = CStr(item)
        End If

End Function


Private Sub ValidateCompare( _
    ByVal source As String, _
    ByVal compareMode As Long)
'
'==============================================================================
'                               ValidateCompare
'------------------------------------------------------------------------------
' PURPOSE
'   Accept only the two comparison modes available in every Office host.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for anything other than vbBinaryCompare or
'   vbTextCompare; vbDatabaseCompare exists only in Access.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Name the calling procedure so the diagnostic identifies the operation.
        If compareMode <> vbBinaryCompare And compareMode <> vbTextCompare Then
            RaiseInvalid source, "compareMode must be vbBinaryCompare or vbTextCompare."
        End If

End Sub


Private Sub RaiseInvalid( _
    ByVal source As String, _
    ByVal description As String)
'
'==============================================================================
'                                 RaiseInvalid
'------------------------------------------------------------------------------
' PURPOSE
'   Raise the core invalid-argument error with a stated rule.
'
' ERROR POLICY
'   Always raises ERR_INVALID_ARGUMENT; TextFacade replaces the source.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RAISE
'------------------------------------------------------------------------------
    'Keep one numeric contract for every rejected argument.
        Err.Raise ERR_INVALID_ARGUMENT, source, description

End Sub
