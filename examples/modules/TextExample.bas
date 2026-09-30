Attribute VB_Name = "TextExample"
'==============================================================================
' MODULE: TextExample
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate several supported TextFacade operations on explicit inputs.
'
' PUBLIC SURFACE
'   RunTextExample is an in-project example macro, not production API.
'
' DEPENDENCIES
'   TextFacade only.
'
' STATE OWNERSHIP
'   No mutable state. Output is written only to the VBE Immediate window.
'
' ERROR POLICY
'   Does not suppress facade errors; callers see the supported error contract.
'
' WORKSHEET SAFETY
'   Does not read or modify Application, workbook, worksheet, Range, selection,
'   calculation, events, display settings, the file system, or other host
'   state. The single-column array stands in for a Range.Value.
'
' TEST SEAM
'   TextTests covers the same facade behavior with deterministic assertions.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS; text-only inputs keep the output identical
'   under every regional setting. Path lines use the platform separator.
'
' USAGE
'   Import the required production modules first, then run
'   TextExample.RunTextExample from the VBE Immediate window.
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


'
'------------------------------------------------------------------------------
'
'                             EXAMPLE ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunTextExample()
'
'==============================================================================
'                                RunTextExample
'------------------------------------------------------------------------------
' PURPOSE
'   Show a small cleaning pipeline: normalize a label, split and rejoin a
'   list, pad a code, and build a safe output path.
'
' USAGE
'   Run TextExample.RunTextExample from the VBE Immediate window.
'
' SIDE EFFECTS
'   Print six lines to the Immediate window; no host state is changed.
'
' ERROR POLICY
'   Facade errors propagate to the caller.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep every input explicit so the printed output is reproducible.
    Dim column(1 To 3, 1 To 1)   As Variant          'Stand-in for a one-column Range.Value
    Dim items()                  As String           'Trimmed, non-empty list items
    Dim parts                    As TextPathParts    'Pieces of the composed output path
    Dim outputPath               As String           'Safe report path under the base folder

'------------------------------------------------------------------------------
' NORMALIZE AND LIST
'------------------------------------------------------------------------------
    'Clean a pasted label, then turn a loosely typed list into clean items.
        Debug.Print "Label: [" & TextFacade.TextCollapseWhitespace("  Quarterly" & vbTab & " report  ") & "]"
        items = TextFacade.TextSplit(" north, ,south ,  east ", ",", True, True)
        Debug.Print "Regions: " & TextFacade.TextJoin(items, " | ")

'------------------------------------------------------------------------------
' JOIN A COLUMN AND PAD A CODE
'------------------------------------------------------------------------------
    'Join a worksheet-shaped column, skipping its blank cell, and pad a
    'numeric code held as text.
        column(1, 1) = "alpha"
        column(2, 1) = Empty
        column(3, 1) = "gamma"
        Debug.Print "Column: " & TextFacade.TextJoin(column, ", ", True)
        Debug.Print "Code: " & TextFacade.TextPad("42", 6, TextPadLeft, "0")

'------------------------------------------------------------------------------
' BUILD A SAFE PATH
'------------------------------------------------------------------------------
    'Make a user-typed title safe for a file name, append it to a base folder
    'and read the extension back; the disk is never touched.
        outputPath = TextFacade.TextJoinPath( _
            "reports", _
            TextFacade.TextSafeFileName("Q3: North/South?.csv"))
        parts = TextFacade.TextSplitPath(outputPath)
        Debug.Print "Path: " & outputPath
        Debug.Print "Extension: " & parts.Extension

End Sub
