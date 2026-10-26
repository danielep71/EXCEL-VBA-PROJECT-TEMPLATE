Attribute VB_Name = "TextFacade"
'==============================================================================
' MODULE: TextFacade
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the library profile's supported text and path utilities and
'   translate core failures into a stable caller-facing error contract.
'
' PUBLIC SURFACE
'   TEXT_ERROR_INVALID_ARGUMENT, TEXT_PATH_SEPARATOR
'   TextPadSide (Enum), TextPathParts (Type)
'   TextPad, TextCollapseWhitespace, TextStartsWith, TextEndsWith, TextCount,
'   TextSplit, TextJoin, TextIsBlank
'   TextJoinPath, TextSplitPath, TextSafeFileName
'
' DEPENDENCIES
'   TextCore only. The dependency direction is facade -> core.
'
' STATE OWNERSHIP
'   Stateless. This module owns no Application, workbook, worksheet, Range, UI,
'   callback, file-system, or module-level mutable state.
'
' ERROR POLICY
'   Every rejected argument raises TEXT_ERROR_INVALID_ARGUMENT with the called
'   facade procedure as the source. The public constant aliases the core-owned
'   code, so the numeric contract has one source of truth; the description,
'   help file and help context supplied by the core are preserved.
'
' WORKSHEET SAFETY
'   Uses only explicit arguments. It never reads Application.Caller,
'   ActiveWorkbook, ActiveSheet, Selection, or other ambient Excel state, and
'   path procedures never touch the file system. Callers pass Range.Value,
'   never a Range object.
'
' TEST SEAM
'   TextTests exercises this public surface. TextCore remains separately
'   addressable inside the VBA project without becoming supported public API.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS; no optional references. Path separator,
'   rooted-path and file-name rules follow the compile-time platform.
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

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Expose core-owned values without defining a second source of truth.
        Public Const TEXT_ERROR_INVALID_ARGUMENT   As Long = TextCore.ERR_INVALID_ARGUMENT    'Public alias of the core error
        Public Const TEXT_PATH_SEPARATOR           As String = TextCore.PATH_SEPARATOR        'Separator this platform writes

'------------------------------------------------------------------------------
' MODULE TYPES
'------------------------------------------------------------------------------
    'Name the padding sides; each value aliases the core constant it selects.
        Public Enum TextPadSide
            TextPadRight = TextCore.PAD_RIGHT    'Fill after the text (default)
            TextPadLeft = TextCore.PAD_LEFT      'Fill before the text
            TextPadBoth = TextCore.PAD_BOTH      'Center; odd character goes right
        End Enum

    'Hold the three parts of a path returned by TextSplitPath.
        Public Type TextPathParts
            Directory   As String    'Text through the last separator, or empty
            BaseName    As String    'File name without its extension
            Extension   As String    'Text after the last dot, without the dot
        End Type


'
'------------------------------------------------------------------------------
'
'                             SUPPORTED PUBLIC API
'
'------------------------------------------------------------------------------
'

Public Function TextPad( _
    ByVal value As String, _
    ByVal width As Long, _
    Optional ByVal side As TextPadSide = TextPadRight, _
    Optional ByVal fill As String = " ") _
    As String
'
'==============================================================================
'                                   TextPad
'------------------------------------------------------------------------------
' PURPOSE
'   Widen text to a minimum length with one repeated fill character.
'
' INPUTS
'   value: text to widen; never truncated.
'   width: minimum result length in characters; zero or greater.
'   side: TextPadRight (default), TextPadLeft or TextPadBoth.
'   fill: exactly one character; a space by default.
'
' RETURNS
'   value unchanged when already wide enough; otherwise exactly width
'   characters. TextPadBoth puts the odd extra character on the right.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   a negative width, an unknown side, or a fill that is not one character.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextPad = TextCore.PadText(value, width, side, fill)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextPad"

End Function


Public Function TextCollapseWhitespace( _
    ByVal value As String) _
    As String
'
'==============================================================================
'                            TextCollapseWhitespace
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
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this operation has no failure to normalize.
        TextCollapseWhitespace = TextCore.CollapseWhitespace(value)

End Function


Public Function TextStartsWith( _
    ByVal value As String, _
    ByVal prefix As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) _
    As Boolean
'
'==============================================================================
'                                TextStartsWith
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether text begins with a prefix.
'
' INPUTS
'   value, prefix: text to compare; an empty prefix always matches.
'   compareMode: vbBinaryCompare (default, exact) or vbTextCompare.
'
' RETURNS
'   True when value begins with prefix under compareMode.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   a compareMode other than vbBinaryCompare or vbTextCompare.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextStartsWith = TextCore.StartsWithText(value, prefix, compareMode)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextStartsWith"

End Function


Public Function TextEndsWith( _
    ByVal value As String, _
    ByVal suffix As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) _
    As Boolean
'
'==============================================================================
'                                 TextEndsWith
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether text ends with a suffix.
'
' INPUTS
'   value, suffix: text to compare; an empty suffix always matches.
'   compareMode: vbBinaryCompare (default, exact) or vbTextCompare.
'
' RETURNS
'   True when value ends with suffix under compareMode.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   a compareMode other than vbBinaryCompare or vbTextCompare.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextEndsWith = TextCore.EndsWithText(value, suffix, compareMode)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextEndsWith"

End Function


Public Function TextCount( _
    ByVal value As String, _
    ByVal find As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) _
    As Long
'
'==============================================================================
'                                  TextCount
'------------------------------------------------------------------------------
' PURPOSE
'   Count non-overlapping occurrences of one text in another.
'
' INPUTS
'   value: text to search.
'   find: nonempty text to count.
'   compareMode: vbBinaryCompare (default, exact) or vbTextCompare.
'
' RETURNS
'   Occurrences found left to right; "aaaa" contains "aa" twice.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   an empty find or an unknown compareMode.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextCount = TextCore.CountText(value, find, compareMode)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextCount"

End Function


Public Function TextSplit( _
    ByVal value As String, _
    Optional ByVal delimiter As String = ",", _
    Optional ByVal trimItems As Boolean = False, _
    Optional ByVal removeEmpty As Boolean = False) _
    As String()
'
'==============================================================================
'                                  TextSplit
'------------------------------------------------------------------------------
' PURPOSE
'   Split text on a delimiter with optional trimming and empty-item removal.
'
' INPUTS
'   value: text to split; an empty value yields no items.
'   delimiter: nonempty literal separator; a comma by default.
'   trimItems: trim whitespace at both ends of each item.
'   removeEmpty: drop items that are empty after optional trimming.
'
' RETURNS
'   A zero-based String array; no items gives LBound 0 and UBound -1.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   an empty delimiter.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextSplit = TextCore.SplitText(value, delimiter, trimItems, removeEmpty)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextSplit"

End Function


Public Function TextJoin( _
    ByVal items As Variant, _
    Optional ByVal delimiter As String = ", ", _
    Optional ByVal skipEmpty As Boolean = False) _
    As String
'
'==============================================================================
'                                   TextJoin
'------------------------------------------------------------------------------
' PURPOSE
'   Join the scalar items of an array, such as a one-row or one-column
'   Range.Value.
'
' INPUTS
'   items: a one-dimensional array, or a two-dimensional array with one row
'          or one column. Pass Range.Value, never the Range object.
'   delimiter: text between items; comma and space by default.
'   skipEmpty: omit items whose text is empty.
'
' RETURNS
'   Items in array order. Empty and Null become empty text; other scalars use
'   CStr, so numbers and dates follow the host's regional settings.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   a non-array, a shape with several rows and columns, or an object,
'   nested array or error value among the items.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextJoin = TextCore.JoinText(items, delimiter, skipEmpty)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextJoin"

End Function


Public Function TextIsBlank( _
    ByVal value As Variant) _
    As Boolean
'
'==============================================================================
'                                 TextIsBlank
'------------------------------------------------------------------------------
' PURPOSE
'   Classify a cell-style value as blank.
'
' INPUTS
'   value: any Variant, such as one element of Range.Value.
'
' RETURNS
'   True for Empty, Null, and text that is empty or whitespace only; False
'   for every other value, including zero, False and error values.
'
' ERROR POLICY
'   Raises no errors.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this operation has no failure to normalize.
        TextIsBlank = TextCore.IsBlankValue(value)

End Function


Public Function TextJoinPath( _
    ByVal basePath As String, _
    ByVal childPath As String) _
    As String
'
'==============================================================================
'                                 TextJoinPath
'------------------------------------------------------------------------------
' PURPOSE
'   Append a relative path to a base path with exactly one separator.
'
' INPUTS
'   basePath: any path text; trailing separators are replaced by one.
'   childPath: a relative path; empty returns basePath unchanged.
'
' RETURNS
'   basePath, one TEXT_PATH_SEPARATOR and childPath. A root such as "C:\"
'   or "/" keeps its separator. The file system is never consulted.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   a rooted childPath: a leading separator or, on Windows, a drive.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextJoinPath = TextCore.JoinPathText(basePath, childPath)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextJoinPath"

End Function


Public Function TextSplitPath( _
    ByVal path As String) _
    As TextPathParts
'
'==============================================================================
'                                TextSplitPath
'------------------------------------------------------------------------------
' PURPOSE
'   Separate a path into directory, base name and extension.
'
' INPUTS
'   path: any path text; the file system is never consulted.
'
' RETURNS
'   Directory: everything through the last separator, or empty.
'   BaseName:  the file name without its extension.
'   Extension: text after the last dot, without the dot; empty for no dot, a
'              trailing dot, or a leading-dot name such as ".gitignore".
'   Directory & BaseName, plus "." & Extension when Extension is nonempty,
'   rebuilds path exactly.
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
    'Assemble the public type from the core's three output parameters.
    Dim parts   As TextPathParts    'Result returned to the caller

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this operation has no failure to normalize.
        TextCore.SplitPathText path, parts.Directory, parts.BaseName, parts.Extension
        TextSplitPath = parts

End Function


Public Function TextSafeFileName( _
    ByVal value As String, _
    Optional ByVal replacement As String = "_") _
    As String
'
'==============================================================================
'                               TextSafeFileName
'------------------------------------------------------------------------------
' PURPOSE
'   Turn arbitrary text into a file name this platform accepts.
'
' INPUTS
'   value: proposed file name, without a directory.
'   replacement: one valid character (underscore by default), or empty to
'                delete invalid characters.
'
' RETURNS
'   value with invalid characters replaced; on Windows also without trailing
'   spaces or dots and with reserved device names such as CON prefixed by an
'   underscore.
'
' ERROR POLICY
'   Raise TEXT_ERROR_INVALID_ARGUMENT from this procedure for
'   an unusable replacement, or a value with no usable characters.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        TextSafeFileName = TextCore.SafeFileNameText(value, replacement)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "TextFacade.TextSafeFileName"

End Function


'
'------------------------------------------------------------------------------
'
'                               PRIVATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub RaiseFromCore( _
    ByVal source As String)
'
'==============================================================================
'                                RaiseFromCore
'------------------------------------------------------------------------------
' PURPOSE
'   Re-raise the active core error with a facade procedure as its source.
'
' INPUTS
'   source: the public procedure name, such as "TextFacade.TextPad".
'
' ERROR POLICY
'   Called only from a facade error handler, so the core error is still the
'   active one. Raising here passes the error to the facade's caller with the
'   same number, description, help file and help context.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep every error field needed to preserve the core failure contract.
    Dim savedDescription   As String    'Original core diagnostic for re-raise
    Dim savedHelpContext   As Long      'Original help topic passed through the facade
    Dim savedHelpFile      As String    'Original help file passed through the facade
    Dim savedNumber        As Long      'Original error number for re-raise

'------------------------------------------------------------------------------
' RE-RAISE
'------------------------------------------------------------------------------
    'Capture all fields before Err.Raise replaces the active error record.
        savedNumber = Err.Number
        savedDescription = Err.Description
        savedHelpFile = Err.HelpFile
        savedHelpContext = Err.HelpContext

    'Re-raise with the public source while retaining all other core fields.
        Err.Raise _
            savedNumber, _
            source, _
            savedDescription, _
            savedHelpFile, _
            savedHelpContext

End Sub
