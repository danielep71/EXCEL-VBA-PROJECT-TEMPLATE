Attribute VB_Name = "TextTests"
'==============================================================================
' MODULE: TextTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run a deterministic, dependency-free regression suite for the library
'   profile's TextFacade and report complete evidence to the Immediate window.
'
' PUBLIC SURFACE
'   RunTextTests is the library suite's entry point. ResetTextTests is a
'   project-private recovery command for an interrupted run; Option Private
'   Module keeps both procedures out of the external workbook automation API.
'
' DEPENDENCIES
'   TextFacade and the built-in VBA/Excel object models only. No external
'   references, workbook fixture, worksheet, donor project, or test framework.
'
' STATE OWNERSHIP
'   Owns private counters, failure text, suite-completeness state, and a re-entry
'   flag. Cleanup clears the re-entry flag; report state is retained until the
'   next run/reset so the final summary remains inspectable.
'
' ERROR POLICY
'   Expected facade errors are captured and asserted. Unexpected runner errors
'   are recorded, cleanup runs, and the original number/source/description is
'   re-raised. Assertion failures raise one test-suite error after reporting.
'
' WORKSHEET SAFETY
'   Reads selected Application properties to report the environment and prove
'   calculation, display-alert, event, and screen-updating state did not change.
'   It never creates, selects, activates, edits, calculates, saves, or closes
'   workbook or worksheet state, and never touches the file system.
'
' TEST SEAM
'   Tests the supported TextFacade surface. Expected values are literals or
'   come from built-in VBA functions, never from the code under test; the
'   platform-specific expectations are selected by conditional compilation.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS; host evidence identifies the actual Office
'   bitness and VBA generation. No external framework or fixture is required.
'
' USAGE
'   Import the library production modules first, then run
'   TextTests.RunTextTests from the VBE Immediate window. ProjectTests keeps
'   covering the shared starter facade.
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
    'Keep harness error codes separate from the production failure contract;
    'the expected counts define when the suite is complete.
        Private Const TEST_ERROR_DIRTY_START   As Long = vbObjectError + 2160    'Refused re-entry or interrupted run
        Private Const TEST_ERROR_FAILURES      As Long = vbObjectError + 2161    'Failed or incomplete suite outcome
        Private Const INVALID_OPERATIONS       As Long = 15                      'Invalid-argument operations, one case each
        Private Const EXPECTED_ASSERTIONS      As Long = 123                     'Required assertions for a complete run
        Private Const EXPECTED_CASES           As Long = 26                      'Required cases for a complete run

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Own counters and diagnostics for one run; cleanup releases the active
    'flag while report values remain available until the next reset.
        Private mCaseCount        As Long       'Cases started in the current run
        Private mAssertionCount   As Long       'Assertions evaluated in the current run
        Private mFailureCount     As Long       'Failures recorded, including runner errors
        Private mFailureDetails   As String     'Diagnostics retained until the next reset
        Private mRunActive        As Boolean    'Re-entry guard; cleared by cleanup or reset
        Private mSuiteCompleted   As Boolean    'True only when both expected counts match


'
'------------------------------------------------------------------------------
'
'                              TEST ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunTextTests()
'
'==============================================================================
'                                 RunTextTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every library case and report its assertions with cleanup evidence.
'
' USAGE
'   Run TextTests.RunTextTests from the VBE Immediate window.
'
' STATE OWNERSHIP
'   Reject an active run before resetting counters. Snapshot four Excel
'   properties for comparison; the harness never changes those properties.
'
' ERROR POLICY
'   Record unexpected runner failures, attempt cleanup, print the summary,
'   then re-raise the saved error. Other failed outcomes raise the suite error.
'   A dirty start raises immediately and does not run cleanup.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host snapshot separate from the saved runner error so cleanup
    'can be reported without losing the original failure.
    Dim cleanupDetail           As String           'Cleanup explanation included in the report
    Dim cleanupPassed           As Boolean          'True only when cleanup verification succeeds
    Dim initialCalculation      As XlCalculation    'Calculation mode captured before the suite
    Dim initialDisplayAlerts    As Boolean          'Display-alert setting captured before the suite
    Dim initialEnableEvents     As Boolean          'Event setting captured before the suite
    Dim initialScreenUpdating   As Boolean          'Screen-update setting captured before the suite
    Dim savedDescription        As String           'Original diagnostic preserved through cleanup
    Dim savedNumber             As Long             'Original error number for later re-raise
    Dim savedSource             As String           'Original runner error source for re-raise
    Dim stateSnapshotValid      As Boolean          'True only after all four properties were read

'------------------------------------------------------------------------------
' GUARD ENTRY
'------------------------------------------------------------------------------
    'Refuse an active run before clearing its counters or failure details.
    'An interrupted execution must be reset explicitly by the caller.
        If mRunActive Then
            Debug.Print "RESULT=FAIL_DIRTY_START; cleanup=NOT_RUN"
            Err.Raise _
                TEST_ERROR_DIRTY_START, _
                "TextTests.RunTextTests", _
                "A TextTests run is already active. Run ResetTextTests after an interrupted execution."
        End If

'------------------------------------------------------------------------------
' INITIALIZE RUN
'------------------------------------------------------------------------------
    'Start with empty report state, then claim the run before executing
    'anything that can fail through the shared runner handler.
        ResetRun
        mRunActive = True
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Read all four properties before marking the snapshot valid. If a read
    'fails, cleanup must report an unavailable snapshot rather than compare
    'uninitialized values or claim that host state was unchanged.
        initialCalculation = Application.Calculation
        initialDisplayAlerts = Application.DisplayAlerts
        initialEnableEvents = Application.EnableEvents
        initialScreenUpdating = Application.ScreenUpdating
        stateSnapshotValid = True

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run cases in the evidence contract order; each case records unexpected
    'errors so later cases can still contribute to the report.
        PrintEnvironment
        TestPad
        TestCollapseWhitespace
        TestStartsAndEndsWith
        TestCount
        TestSplit
        TestJoin
        TestIsBlank
        TestJoinPath
        TestSplitPath
        TestSafeFileName
        TestInvalidArguments
        TestRepeatability

    'Require both counts so an early return cannot produce a complete PASS.
        mSuiteCompleted = _
            (mCaseCount = EXPECTED_CASES) And _
            (mAssertionCount = EXPECTED_ASSERTIONS)
        If Not mSuiteCompleted Then
            RecordFailure _
                "suite.completeness", _
                "expected cases=" & CStr(EXPECTED_CASES) & _
                    ", assertions=" & CStr(EXPECTED_ASSERTIONS) & _
                    "; observed cases=" & CStr(mCaseCount) & _
                    ", assertions=" & CStr(mAssertionCount)
        End If

'------------------------------------------------------------------------------
' CLEANUP AND REPORT
'------------------------------------------------------------------------------
CleanExit:
    'Disable the runner handler before cleanup and reporting. A later raise
    'must escape to the caller instead of re-entering the runner handler.
        On Error GoTo 0
        cleanupPassed = CleanupRun( _
            cleanupDetail, _
            stateSnapshotValid, _
            initialCalculation, _
            initialDisplayAlerts, _
            initialEnableEvents, _
            initialScreenUpdating)
        PrintSummary cleanupPassed, cleanupDetail

    'Propagate the original runner error only after cleanup and reporting.
        If savedNumber <> 0 Then
            Err.Raise savedNumber, savedSource, savedDescription
        End If

    'Make assertion, cleanup and completeness failures visible to callers
    'even when no unexpected runner error was saved.
        If mFailureCount <> 0 Or Not cleanupPassed Or Not mSuiteCompleted Then
            Err.Raise _
                TEST_ERROR_FAILURES, _
                "TextTests.RunTextTests", _
                "Regression failed; review the Immediate window report."
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE RUNNER ERROR
'------------------------------------------------------------------------------
RunFailed:
    'Save the escaping error before recording its diagnostic; Resume then
    'leaves the active handler through the common cleanup path.
        savedNumber = Err.Number
        savedSource = Err.Source
        savedDescription = Err.Description
        RecordFailure _
            "runner.unexpected", _
            "error=" & CStr(savedNumber) & _
                "; source=" & savedSource & _
                "; description=" & savedDescription
        Resume CleanExit

End Sub


Public Sub ResetTextTests()
'
'==============================================================================
'                                ResetTextTests
'------------------------------------------------------------------------------
' PURPOSE
'   Recover harness-owned state after an interrupted execution.
'
' USAGE
'   Run only after the previous execution has stopped; then rerun the suite.
'
' SIDE EFFECTS
'   Clear counters, failure text, completion, and the active-run flag.
'   Print confirmation; do not alter Excel application properties.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Recover only harness-owned state after execution has stopped; the
    'confirmation distinguishes an explicit reset from a completed test run.
        ResetRun
        Debug.Print "TEXT TESTS RESET; run_active=no; counters=0"

End Sub


'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Private Sub TestPad()
'
'==============================================================================
'                                   TestPad
'------------------------------------------------------------------------------
' PURPOSE
'   Check padding on each side, the defaults, centering and no truncation.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Centering a three-character gap puts one fill on the left and two on
    'the right; a value longer than the width must come back unchanged.
        On Error GoTo CaseFailed

        BeginCase "text.pad"
        AssertEqualString "TextPad left zeros", "007", TextFacade.TextPad("7", 3, TextPadLeft, "0")
        AssertEqualString "TextPad defaults", "ab   ", TextFacade.TextPad("ab", 5)
        AssertEqualString "TextPad both", "*ab**", TextFacade.TextPad("ab", 5, TextPadBoth, "*")
        AssertEqualString "TextPad no truncation", "abcdef", TextFacade.TextPad("abcdef", 3)
        AssertEqualString "TextPad zero width", "", TextFacade.TextPad("", 0)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.pad"

End Sub


Private Sub TestCollapseWhitespace()
'
'==============================================================================
'                            TestCollapseWhitespace
'------------------------------------------------------------------------------
' PURPOSE
'   Check trimming and run collapsing for tabs, line breaks and U+00A0.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Mix every whitespace class so a missed character leaves a visible
    'double space or an untrimmed end.
        On Error GoTo CaseFailed

        BeginCase "text.collapse-whitespace"
        AssertEqualString _
            "Collapse tabs and line breaks", _
            "a b c", _
            TextFacade.TextCollapseWhitespace("  a" & vbTab & vbTab & "b" & vbCrLf & " c  ")
        AssertEqualString _
            "Collapse no-break spaces", _
            "x y", _
            TextFacade.TextCollapseWhitespace(ChrW(160) & "x" & ChrW(160) & ChrW(160) & "y" & ChrW(160))
        AssertEqualString _
            "Collapse whitespace only", _
            "", _
            TextFacade.TextCollapseWhitespace(vbTab & " " & vbLf)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.collapse-whitespace"

End Sub


Private Sub TestStartsAndEndsWith()
'
'==============================================================================
'                            TestStartsAndEndsWith
'------------------------------------------------------------------------------
' PURPOSE
'   Check exact and case-insensitive prefix and suffix tests and their edges.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The default must be case-sensitive; vbTextCompare must relax only case.
        On Error GoTo CaseFailed

        BeginCase "text.starts-ends-with"
        AssertEqualBoolean "StartsWith exact", True, TextFacade.TextStartsWith("Report.xlsx", "Rep")
        AssertEqualBoolean "StartsWith default is case-sensitive", False, TextFacade.TextStartsWith("Report.xlsx", "rep")
        AssertEqualBoolean "StartsWith text compare", True, TextFacade.TextStartsWith("Report.xlsx", "rep", vbTextCompare)
        AssertEqualBoolean "EndsWith text compare", True, TextFacade.TextEndsWith("Report.xlsx", ".XLSX", vbTextCompare)
        AssertEqualBoolean "EndsWith longer suffix", False, TextFacade.TextEndsWith("a", "abc")
        AssertEqualBoolean "StartsWith empty prefix", True, TextFacade.TextStartsWith("abc", "")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.starts-ends-with"

End Sub


Private Sub TestCount()
'
'==============================================================================
'                                  TestCount
'------------------------------------------------------------------------------
' PURPOSE
'   Check non-overlapping counting, comparison modes and an empty value.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    '"aaaa" holds three overlapping "aa" matches but only two that do not
    'overlap; the expected value fixes the documented rule.
        On Error GoTo CaseFailed

        BeginCase "text.count"
        AssertEqualLong "Count non-overlapping", 2, TextFacade.TextCount("aaaa", "aa")
        AssertEqualLong "Count repeated pair", 2, TextFacade.TextCount("Banana", "an")
        AssertEqualLong "Count text compare", 1, TextFacade.TextCount("Banana", "BAN", vbTextCompare)
        AssertEqualLong "Count in empty value", 0, TextFacade.TextCount("", "x")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.count"

End Sub


Private Sub TestSplit()
'
'==============================================================================
'                                  TestSplit
'------------------------------------------------------------------------------
' PURPOSE
'   Check splitting with and without trimming and empty-item removal.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Expected items are written pipe-separated and expanded with the VBA
    'Split function, so no expectation depends on the code under test.
        On Error GoTo CaseFailed

        BeginCase "text.split"
        AssertStringArray "Split keeps empty items", "a|b||c", TextFacade.TextSplit("a,b,,c")
        AssertStringArray "Split trims items", "a|b||c", TextFacade.TextSplit(" a , b ,, c ", ",", True)
        AssertStringArray "Split removes empty items", "a|b|c", TextFacade.TextSplit(" a , b ,, c ", ",", True, True)
        AssertStringArray "Split empty value", "", TextFacade.TextSplit("", ",")
        AssertStringArray "Split multi-character delimiter", "x|y", TextFacade.TextSplit("x--y", "--")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.split"

End Sub


Private Sub TestJoin()
'
'==============================================================================
'                                   TestJoin
'------------------------------------------------------------------------------
' PURPOSE
'   Check joining one-dimensional, worksheet-shaped and unallocated arrays.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Build each input shape explicitly; integers keep CStr locale-neutral.
    Dim column(1 To 3, 1 To 1)   As Variant    'One-column Range.Value shape
    Dim offset(5 To 6)           As Variant    'One-dimensional array with bounds 5 and 6
    Dim row(0 To 0, 0 To 2)      As Variant    'One-row Range.Value shape
    Dim unallocated()            As String     'Dynamic array never dimensioned

'------------------------------------------------------------------------------
' PREPARE INPUTS
'------------------------------------------------------------------------------
    'Fill the fixed arrays before the case starts so a setup error cannot
    'be mistaken for a join failure.
        On Error GoTo CaseFailed

        column(1, 1) = "x"
        column(2, 1) = "y"
        column(3, 1) = "z"
        row(0, 0) = 1
        row(0, 1) = 2
        row(0, 2) = 3
        offset(5) = "p"
        offset(6) = "q"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Empty and Null join as empty text; skipEmpty removes them entirely.
        BeginCase "text.join"
        AssertEqualString "Join one-dimensional", "a-b-c", TextFacade.TextJoin(Array("a", "b", "c"), "-")
        AssertEqualString "Join skips empty items", "a, d", TextFacade.TextJoin(Array("a", "", Empty, Null, "d"), , True)
        AssertEqualString "Join keeps empty items", "a, , , , d", TextFacade.TextJoin(Array("a", "", Empty, Null, "d"))
        AssertEqualString "Join column", "x;y;z", TextFacade.TextJoin(column, ";")
        AssertEqualString "Join row", "1+2+3", TextFacade.TextJoin(row, "+")
        AssertEqualString "Join unallocated array", "", TextFacade.TextJoin(unallocated)
        AssertEqualString "Join nonzero lower bound", "p/q", TextFacade.TextJoin(offset, "/")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.join"

End Sub


Private Sub TestIsBlank()
'
'==============================================================================
'                                 TestIsBlank
'------------------------------------------------------------------------------
' PURPOSE
'   Check blank classification for every cell-value category.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Zero, False, "0" and a worksheet error are values, not blanks.
        On Error GoTo CaseFailed

        BeginCase "text.is-blank"
        AssertEqualBoolean "IsBlank Empty", True, TextFacade.TextIsBlank(Empty)
        AssertEqualBoolean "IsBlank Null", True, TextFacade.TextIsBlank(Null)
        AssertEqualBoolean "IsBlank empty text", True, TextFacade.TextIsBlank("")
        AssertEqualBoolean "IsBlank whitespace", True, TextFacade.TextIsBlank(" " & vbTab & ChrW(160))
        AssertEqualBoolean "IsBlank zero text", False, TextFacade.TextIsBlank("0")
        AssertEqualBoolean "IsBlank zero", False, TextFacade.TextIsBlank(0)
        AssertEqualBoolean "IsBlank False", False, TextFacade.TextIsBlank(False)
        AssertEqualBoolean "IsBlank error value", False, TextFacade.TextIsBlank(CVErr(2042))
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "text.is-blank"

End Sub


Private Sub TestJoinPath()
'
'==============================================================================
'                                 TestJoinPath
'------------------------------------------------------------------------------
' PURPOSE
'   Check the platform separator, joining, trailing separators, empty sides
'   and a root.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Derive platform expectations here, never from the code under test.
    Dim root        As String    'A root path on this platform
    Dim separator   As String    'Expected separator on this platform

'------------------------------------------------------------------------------
' SELECT PLATFORM EXPECTATIONS
'------------------------------------------------------------------------------
    'State the documented separator and root for the compiled platform.
#If Mac Then
        separator = "/"
        root = "/"
#Else
        separator = "\"
        root = "C:\"
#End If

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Repeated trailing separators collapse to one; a root keeps its own.
        On Error GoTo CaseFailed

        BeginCase "path.join"
        AssertEqualString "Path separator constant", separator, TextFacade.TEXT_PATH_SEPARATOR
        AssertEqualString "JoinPath simple", "reports" & separator & "q3.csv", TextFacade.TextJoinPath("reports", "q3.csv")
        AssertEqualString _
            "JoinPath trailing separators", _
            "reports" & separator & "q3.csv", _
            TextFacade.TextJoinPath("reports" & separator & separator, "q3.csv")
        AssertEqualString "JoinPath empty base", "q3.csv", TextFacade.TextJoinPath("", "q3.csv")
        AssertEqualString "JoinPath empty child", "reports", TextFacade.TextJoinPath("reports", "")
        AssertEqualString "JoinPath root", root & "data", TextFacade.TextJoinPath(root, "data")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "path.join"

End Sub


Private Sub TestSplitPath()
'
'==============================================================================
'                                TestSplitPath
'------------------------------------------------------------------------------
' PURPOSE
'   Check directory, base-name and extension rules, including dot-files,
'   trailing dots, directories and platform separators.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold each result and the platform-specific expectations.
    Dim backslashDirectory   As String           'Expected directory of "a\b.txt"
    Dim parts                As TextPathParts    'Result of the latest split
    Dim separator            As String           'Expected separator on this platform

'------------------------------------------------------------------------------
' SELECT PLATFORM EXPECTATIONS
'------------------------------------------------------------------------------
    'A backslash separates directories on Windows only; on macOS it is an
    'ordinary file-name character.
#If Mac Then
        separator = "/"
        backslashDirectory = ""
#Else
        separator = "\"
        backslashDirectory = "a\"
#End If

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Only the last dot starts an extension, and a leading or trailing dot
    'never does.
        On Error GoTo CaseFailed

        BeginCase "path.split"
        parts = TextFacade.TextSplitPath("dir" & separator & "sub" & separator & "file.tar.gz")
        AssertEqualString "SplitPath directory", "dir" & separator & "sub" & separator, parts.Directory
        AssertEqualString "SplitPath base name", "file.tar", parts.BaseName
        AssertEqualString "SplitPath extension", "gz", parts.Extension

        parts = TextFacade.TextSplitPath(".gitignore")
        AssertEqualString "SplitPath dot-file base", ".gitignore", parts.BaseName
        AssertEqualString "SplitPath dot-file extension", "", parts.Extension

        parts = TextFacade.TextSplitPath("name.")
        AssertEqualString "SplitPath trailing-dot base", "name.", parts.BaseName
        AssertEqualString "SplitPath trailing-dot extension", "", parts.Extension

        parts = TextFacade.TextSplitPath("folder" & separator)
        AssertEqualString "SplitPath directory only", "folder" & separator, parts.Directory
        AssertEqualString "SplitPath directory has no name", "", parts.BaseName

        parts = TextFacade.TextSplitPath("a/b.txt")
        AssertEqualString "SplitPath forward slash", "a/", parts.Directory

        parts = TextFacade.TextSplitPath("a\b.txt")
        AssertEqualString "SplitPath backslash", backslashDirectory, parts.Directory
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "path.split"

End Sub


Private Sub TestSafeFileName()
'
'==============================================================================
'                               TestSafeFileName
'------------------------------------------------------------------------------
' PURPOSE
'   Check invalid characters, deletion, trailing characters and device names
'   against each platform's documented rule.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold the expectations that differ between Windows and macOS.
    Dim expectedDeleted    As String    'Expected result of deleting "<" and ">"
    Dim expectedDevice     As String    'Expected result for "con.txt"
    Dim expectedNumbered   As String    'Expected result for "COM1"
    Dim expectedTitle      As String    'Expected result for the sample title
    Dim expectedTrailing   As String    'Expected result for "report. . "

'------------------------------------------------------------------------------
' SELECT PLATFORM EXPECTATIONS
'------------------------------------------------------------------------------
    'macOS rejects only ":" and "/"; Windows also rejects < > " \ | ? *,
    'drops trailing spaces and dots, and reserves device names.
#If Mac Then
        expectedTitle = "Q3_ North_South?.csv"
        expectedDeleted = "a<b>c"
        expectedTrailing = "report. . "
        expectedDevice = "con.txt"
        expectedNumbered = "COM1"
#Else
        expectedTitle = "Q3_ North_South_.csv"
        expectedDeleted = "abc"
        expectedTrailing = "report"
        expectedDevice = "_con.txt"
        expectedNumbered = "_COM1"
#End If

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'COM0 is not a device name on either platform, so it must be unchanged.
        On Error GoTo CaseFailed

        BeginCase "path.safe-file-name"
        AssertEqualString "SafeFileName title", expectedTitle, TextFacade.TextSafeFileName("Q3: North/South?.csv")
        AssertEqualString "SafeFileName deletion", expectedDeleted, TextFacade.TextSafeFileName("a<b>c", "")
        AssertEqualString "SafeFileName trailing", expectedTrailing, TextFacade.TextSafeFileName("report. . ")
        AssertEqualString "SafeFileName device", expectedDevice, TextFacade.TextSafeFileName("con.txt")
        AssertEqualString "SafeFileName numbered device", expectedNumbered, TextFacade.TextSafeFileName("COM1")
        AssertEqualString "SafeFileName not a device", "COM0", TextFacade.TextSafeFileName("COM0")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "path.safe-file-name"

End Sub


Private Sub TestInvalidArguments()
'
'==============================================================================
'                             TestInvalidArguments
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that every rejected argument raises the public invalid-argument
'   error with the facade procedure as source and the documented rule.
'
' ERROR POLICY
'   Each operation runs inside InvokeInvalid, which captures its error. A
'   call that returns normally is a failure; errors while verifying are
'   unexpected failures of the operation's case.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the captured error and the expected contract for one operation.
    Dim actualDescription     As String    'Error description captured from the facade
    Dim actualNumber          As Long      'Error number captured from the facade
    Dim actualSource          As String    'Error source captured from the facade
    Dim caseName              As String    'Stable case identifier for the operation
    Dim expectedDescription   As String    'Documented rule for the operation
    Dim expectedSource        As String    'Facade procedure expected as the source
    Dim operation             As Long      'Operation number, 1 to INVALID_OPERATIONS
    Dim raised                As Boolean   'True when the operation raised an error

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Register one case per operation so a report names the exact rule that
    'failed; each case contributes the same four assertions.
        For operation = 1 To INVALID_OPERATIONS
            On Error GoTo CaseFailed

            ExpectedInvalid operation, caseName, expectedSource, expectedDescription
            BeginCase caseName
            raised = InvokeInvalid(operation, actualNumber, actualSource, actualDescription)
            AssertEqualBoolean caseName & ".raises", True, raised
            AssertExpectedError _
                caseName, _
                TextFacade.TEXT_ERROR_INVALID_ARGUMENT, _
                expectedSource, _
                expectedDescription, _
                actualNumber, _
                actualSource, _
                actualDescription
NextOperation:
        Next operation
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record the failed operation and continue with the next one; Resume
    'leaves the handler so the loop can install it again.
        RecordUnexpectedCaseError caseName
        Resume NextOperation

End Sub


Private Sub TestRepeatability()
'
'==============================================================================
'                              TestRepeatability
'------------------------------------------------------------------------------
' PURPOSE
'   Verify identical results for two calls with the same inputs.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Retain both results so the comparison observes two separate calls.
    Dim firstName     As String    'Baseline safe file name
    Dim secondName    As String    'Result of the identical second call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The utilities are stateless, so a repeated call must match exactly;
    'a difference would reveal hidden module state.
        On Error GoTo CaseFailed

        BeginCase "api.repeatability"
        firstName = TextFacade.TextSafeFileName("Plan: A/B")
        secondName = TextFacade.TextSafeFileName("Plan: A/B")
        AssertEqualString "Repeated SafeFileName", firstName, secondName
        AssertEqualString _
            "Repeated Pad", _
            TextFacade.TextPad("9", 4, TextPadBoth, "-"), _
            TextFacade.TextPad("9", 4, TextPadBoth, "-")
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Report the failed repeatability case and allow suite finalization.
        RecordUnexpectedCaseError "api.repeatability"

End Sub


'
'------------------------------------------------------------------------------
'
'                          CASE AND ASSERTION HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub BeginCase( _
    ByVal caseName As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Register a started case in both the counter and machine-readable log.
'
' INPUTS
'   caseName: stable identifier used by the evidence contract.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' REGISTER CASE
'------------------------------------------------------------------------------
    'Count a case when it starts, even if its first assertion later fails;
    'the matching log record identifies which case was reached.
        mCaseCount = mCaseCount + 1
        Debug.Print "CASE=" & caseName

End Sub


Private Sub ExpectedInvalid( _
    ByVal operation As Long, _
    ByRef caseName As String, _
    ByRef expectedSource As String, _
    ByRef expectedDescription As String)
'
'==============================================================================
'                               ExpectedInvalid
'------------------------------------------------------------------------------
' PURPOSE
'   Name each invalid-argument operation and state its documented contract.
'
' INPUTS
'   operation: 1 to INVALID_OPERATIONS, matching InvokeInvalid.
'
' RETURNS
'   caseName, expectedSource and expectedDescription for the operation.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DESCRIBE OPERATION
'------------------------------------------------------------------------------
    'Keep the expected texts literal so a changed rule is a visible diff.
        Select Case operation
            Case 1
                caseName = "error.pad-width"
                expectedSource = "TextFacade.TextPad"
                expectedDescription = "width must be zero or greater."
            Case 2
                caseName = "error.pad-side"
                expectedSource = "TextFacade.TextPad"
                expectedDescription = "side must be TextPadRight, TextPadLeft or TextPadBoth."
            Case 3
                caseName = "error.pad-fill"
                expectedSource = "TextFacade.TextPad"
                expectedDescription = "fill must be exactly one character."
            Case 4
                caseName = "error.starts-with-compare"
                expectedSource = "TextFacade.TextStartsWith"
                expectedDescription = "compareMode must be vbBinaryCompare or vbTextCompare."
            Case 5
                caseName = "error.ends-with-compare"
                expectedSource = "TextFacade.TextEndsWith"
                expectedDescription = "compareMode must be vbBinaryCompare or vbTextCompare."
            Case 6
                caseName = "error.count-empty-find"
                expectedSource = "TextFacade.TextCount"
                expectedDescription = "find must not be empty."
            Case 7
                caseName = "error.split-empty-delimiter"
                expectedSource = "TextFacade.TextSplit"
                expectedDescription = "delimiter must not be empty."
            Case 8
                caseName = "error.join-not-array"
                expectedSource = "TextFacade.TextJoin"
                expectedDescription = "items must be an array."
            Case 9
                caseName = "error.join-rank"
                expectedSource = "TextFacade.TextJoin"
                expectedDescription = "items must have one or two dimensions."
            Case 10
                caseName = "error.join-shape"
                expectedSource = "TextFacade.TextJoin"
                expectedDescription = "a two-dimensional items array must have one row or one column."
            Case 11
                caseName = "error.join-item"
                expectedSource = "TextFacade.TextJoin"
                expectedDescription = "items must contain only scalar values."
            Case 12
                caseName = "error.join-path-rooted"
                expectedSource = "TextFacade.TextJoinPath"
                expectedDescription = "childPath must be relative."
            Case 13
                caseName = "error.safe-name-replacement-length"
                expectedSource = "TextFacade.TextSafeFileName"
                expectedDescription = "replacement must be empty or one character."
            Case 14
                caseName = "error.safe-name-replacement-invalid"
                expectedSource = "TextFacade.TextSafeFileName"
                expectedDescription = "replacement must be a valid file-name character."
            Case Else
                caseName = "error.safe-name-empty"
                expectedSource = "TextFacade.TextSafeFileName"
                expectedDescription = "value has no usable file-name characters."
        End Select

End Sub


Private Function InvokeInvalid( _
    ByVal operation As Long, _
    ByRef errorNumber As Long, _
    ByRef errorSource As String, _
    ByRef errorDescription As String) _
    As Boolean
'
'==============================================================================
'                                InvokeInvalid
'------------------------------------------------------------------------------
' PURPOSE
'   Call one facade operation with an invalid argument and capture its error.
'
' INPUTS
'   operation: 1 to INVALID_OPERATIONS, matching ExpectedInvalid.
'
' RETURNS
'   True when the call raised; the three ByRef values then hold the error.
'   False when it returned normally; the values are then zero and empty.
'
' ERROR POLICY
'   The local handler captures the expected error, and leaving the function
'   from the handler clears it before the caller continues.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Receive discarded results and hold the invalid array shapes.
    Dim cube(0 To 0, 0 To 0, 0 To 0)   As Variant    'Three-dimensional array
    Dim ignoredCount                   As Long       'Discarded Long result
    Dim ignoredFlag                    As Boolean    'Discarded Boolean result
    Dim ignoredItems()                 As String     'Discarded array result
    Dim ignoredText                    As String     'Discarded String result
    Dim square(0 To 1, 0 To 1)         As Variant    'Two rows and two columns

'------------------------------------------------------------------------------
' CALL OPERATION
'------------------------------------------------------------------------------
    'Clear the outputs first so a call that fails to raise cannot be scored
    'against the previous operation's error.
        errorNumber = 0
        errorSource = vbNullString
        errorDescription = vbNullString
        On Error GoTo Raised

        Select Case operation
            Case 1
                ignoredText = TextFacade.TextPad("x", -1)
            Case 2
                ignoredText = TextFacade.TextPad("x", 2, 7)
            Case 3
                ignoredText = TextFacade.TextPad("x", 2, TextPadLeft, "ab")
            Case 4
                ignoredFlag = TextFacade.TextStartsWith("a", "a", 2)
            Case 5
                ignoredFlag = TextFacade.TextEndsWith("a", "a", 2)
            Case 6
                ignoredCount = TextFacade.TextCount("abc", "")
            Case 7
                ignoredItems = TextFacade.TextSplit("a,b", "")
            Case 8
                ignoredText = TextFacade.TextJoin("abc")
            Case 9
                ignoredText = TextFacade.TextJoin(cube)
            Case 10
                ignoredText = TextFacade.TextJoin(square)
            Case 11
                ignoredText = TextFacade.TextJoin(Array("a", CVErr(2042)))
            Case 12
                ignoredText = TextFacade.TextJoinPath("base", "/abs")
            Case 13
                ignoredText = TextFacade.TextSafeFileName("a", "ab")
            Case 14
                ignoredText = TextFacade.TextSafeFileName("a", ":")
            Case Else
                ignoredText = TextFacade.TextSafeFileName(":/", "")
        End Select
        InvokeInvalid = False
        Exit Function

'------------------------------------------------------------------------------
' CAPTURE ERROR
'------------------------------------------------------------------------------
Raised:
    'Snapshot the error before the handler exits and clears it.
        errorNumber = Err.Number
        errorSource = Err.Source
        errorDescription = Err.Description
        InvokeInvalid = True

End Function


Private Sub AssertEqualBoolean( _
    ByVal assertionName As String, _
    ByVal expected As Boolean, _
    ByVal actual As Boolean)
'
'==============================================================================
'                              AssertEqualBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require exact Boolean equality.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Count the assertion before evaluating it, then record any mismatch
    'without terminating the suite.
        mAssertionCount = mAssertionCount + 1
        If actual <> expected Then
            RecordFailure _
                assertionName, _
                "expected=" & CStr(expected) & "; actual=" & CStr(actual)
        End If

End Sub


Private Sub AssertStringArray( _
    ByVal assertionName As String, _
    ByVal expectedPipeList As String, _
    ByVal actual As Variant)
'
'==============================================================================
'                              AssertStringArray
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require a zero-based array of exact items.
'
' INPUTS
'   assertionName identifies the check. expectedPipeList lists the expected
'   items separated by "|"; an empty list means no items. actual holds the
'   array under test, passed as a Variant so a function result is accepted.
'
' ERROR POLICY
'   Record the first difference in bounds or items as one failure.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Expand the expectation with VBA Split, independent of the code under test.
    Dim expected()   As String    'Expected items, zero-based
    Dim index        As Long      'Position being compared

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Check the bounds first; item comparison assumes equal bounds.
        mAssertionCount = mAssertionCount + 1
        expected = Split(expectedPipeList, "|")
        If LBound(actual) <> 0 Or UBound(actual) <> UBound(expected) Then
            RecordFailure _
                assertionName, _
                "expected bounds 0.." & CStr(UBound(expected)) & _
                    "; actual bounds " & CStr(LBound(actual)) & ".." & CStr(UBound(actual))
            Exit Sub
        End If
        For index = 0 To UBound(expected)
            If StrComp(actual(index), expected(index), vbBinaryCompare) <> 0 Then
                RecordFailure _
                    assertionName, _
                    "item " & CStr(index) & ": expected=" & expected(index) & "; actual=" & actual(index)
                Exit Sub
            End If
        Next index

End Sub


Private Sub AssertExpectedError( _
    ByVal assertionName As String, _
    ByVal expectedNumber As Long, _
    ByVal expectedSource As String, _
    ByVal expectedDescription As String, _
    ByVal actualNumber As Long, _
    ByVal actualSource As String, _
    ByVal actualDescription As String)
'
'==============================================================================
'                             AssertExpectedError
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the three stable fields of a captured facade error.
'
' INPUTS
'   assertionName prefixes three checks. The expected and actual number,
'   source, and description are supplied as captured scalar values.
'
' SIDE EFFECTS
'   Delegate to three assertions; do not read the live Err object.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT ERROR CONTRACT
'------------------------------------------------------------------------------
    'Keep number, source and description as three independent assertions
    'so a partially correct error cannot satisfy the complete contract.
        AssertEqualLong _
            assertionName & ".number", _
            expectedNumber, _
            actualNumber
        AssertEqualString _
            assertionName & ".source", _
            expectedSource, _
            actualSource
        AssertEqualString _
            assertionName & ".description", _
            expectedDescription, _
            actualDescription

End Sub


Private Sub AssertEqualLong( _
    ByVal assertionName As String, _
    ByVal expected As Long, _
    ByVal actual As Long)
'
'==============================================================================
'                               AssertEqualLong
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require exact Long equality.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Compare the integer contract exactly; a mismatched error number must
    'not be accepted because the source or description happens to match.
        mAssertionCount = mAssertionCount + 1
        If actual <> expected Then
            RecordFailure _
                assertionName, _
                "expected=" & CStr(expected) & "; actual=" & CStr(actual)
        End If

End Sub


Private Sub AssertEqualString( _
    ByVal assertionName As String, _
    ByVal expected As String, _
    ByVal actual As String)
'
'==============================================================================
'                              AssertEqualString
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require binary, case-sensitive string equality.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Use binary comparison so case changes in the error source or message
    'remain observable contract differences, independent of text settings.
        mAssertionCount = mAssertionCount + 1
        If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
            RecordFailure _
                assertionName, _
                "expected=""" & expected & """; actual=""" & actual & """"
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                              FAILURE RECORDING
'
'------------------------------------------------------------------------------
'

Private Sub RecordUnexpectedCaseError( _
    ByVal caseName As String)
'
'==============================================================================
'                          RecordUnexpectedCaseError
'------------------------------------------------------------------------------
' PURPOSE
'   Preserve the active case error before recording its diagnostic fields.
'
' INPUTS
'   caseName: stable case identifier; Err supplies the original failure.
'
' SIDE EFFECTS
'   Add a failure record without incrementing the assertion count.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Retain Err fields before any error-state reset or reporting call.
    Dim errorDescription   As String    'Case diagnostic captured before recording
    Dim errorNumber        As Long      'Case error number captured before recording
    Dim errorSource        As String    'Case error source captured before recording

'------------------------------------------------------------------------------
' CAPTURE ERROR
'------------------------------------------------------------------------------
    'Read the original diagnostic before On Error GoTo 0 can clear Err.
        errorNumber = Err.Number
        errorSource = Err.Source
        errorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' RECORD DIAGNOSTIC
'------------------------------------------------------------------------------
    'Add the case-qualified failure without inventing a completed assertion
    'for an operation that raised before its check could run.
        RecordFailure _
            caseName & ".unexpected", _
            "error=" & CStr(errorNumber) & _
                "; source=" & errorSource & _
                "; description=" & errorDescription

End Sub


Private Sub RecordFailure( _
    ByVal assertionName As String, _
    ByVal detail As String)
'
'==============================================================================
'                                RecordFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Accumulate failure details for the final report.
'
' INPUTS
'   assertionName identifies the failure; detail supplies its diagnostic.
'
' STATE OWNERSHIP
'   Increment the failure count and append text to the owned report buffer.
'   Do not change case or assertion counts.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Retain every diagnostic in arrival order. A separator belongs only
    'between records so the final report has no leading empty entry.
        mFailureCount = mFailureCount + 1
        If Len(mFailureDetails) > 0 Then
            mFailureDetails = mFailureDetails & vbNewLine
        End If
        mFailureDetails = mFailureDetails & assertionName & ": " & detail

End Sub


'
'------------------------------------------------------------------------------
'
'                          CLEANUP AND HOST EVIDENCE
'
'------------------------------------------------------------------------------
'

Private Function CleanupRun( _
    ByRef cleanupDetail As String, _
    ByVal stateSnapshotValid As Boolean, _
    ByVal initialCalculation As XlCalculation, _
    ByVal initialDisplayAlerts As Boolean, _
    ByVal initialEnableEvents As Boolean, _
    ByVal initialScreenUpdating As Boolean) _
    As Boolean
'
'==============================================================================
'                                  CleanupRun
'------------------------------------------------------------------------------
' PURPOSE
'   Release the run flag and verify that observed Excel state is unchanged.
'
' INPUTS
'   stateSnapshotValid indicates whether all four initial properties were
'   read. The initial values belong to the current run only.
'
' RETURNS
'   Boolean success and a ByRef cleanupDetail for the final report.
'
' ERROR POLICY
'   Return False for an unavailable snapshot, changed state, or a cleanup
'   error. Report the reason; do not restore properties the harness never owned.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Track the combined comparison without taking ownership of host state.
    Dim excelStateUnchanged   As Boolean    'All four observed properties match the snapshot

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Release the re-entry flag first, even when no complete host snapshot
    'is available to verify the remaining cleanup contract.
        On Error GoTo CleanupFailed

        mRunActive = False
        If Not stateSnapshotValid Then
            cleanupDetail = "run flag cleared; Excel state snapshot unavailable"
            CleanupRun = False
            Exit Function
        End If

    'Compare only properties that were captured successfully. The harness
    'does not restore them because it never changed or owned them.
        excelStateUnchanged = _
            (Application.Calculation = initialCalculation) And _
            (Application.DisplayAlerts = initialDisplayAlerts) And _
            (Application.EnableEvents = initialEnableEvents) And _
            (Application.ScreenUpdating = initialScreenUpdating)

    'Return the combined result and retain a reason for either outcome.
        CleanupRun = excelStateUnchanged
        If excelStateUnchanged Then
            cleanupDetail = _
                "run flag cleared; verified Excel state unchanged " & _
                "(calculation/display alerts/events/screen updating)"
        Else
            cleanupDetail = "run flag cleared; Excel application state changed during test run"
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE CLEANUP ERROR
'------------------------------------------------------------------------------
CleanupFailed:
    'Describe a failed verification and return False. The runner retains
    'its own saved diagnostic for any later re-raise.
        cleanupDetail = _
            "cleanup error=" & CStr(Err.Number) & _
            "; description=" & Err.Description
        Err.Clear
        CleanupRun = False

End Function


Private Sub PrintEnvironment()
'
'==============================================================================
'                               PrintEnvironment
'------------------------------------------------------------------------------
' PURPOSE
'   Write the report preamble and current Excel environment.
'
' ERROR POLICY
'   Host-property read failures propagate to the runner error path.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT ENVIRONMENT
'------------------------------------------------------------------------------
    'Write the preamble before the host fields so retained output can be
    'recognized and attributed to the environment that produced it.
        Debug.Print "TEXT TESTS"
        Debug.Print "ENVIRONMENT=" & EnvironmentSummary()

End Sub


Private Function EnvironmentSummary() _
    As String
'
'==============================================================================
'                              EnvironmentSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Describe the current Excel host and compiled VBA environment.
'
' RETURNS
'   Machine-readable host, version, operating system, Office bitness, and
'   VBA generation fields. Compile-time branches select the last two values.
'
' ERROR POLICY
'   Host-property read failures propagate to the runner error path.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep compile-time properties distinct from values read from Excel.
    Dim bitness         As String    'Office bitness selected by conditional compilation
    Dim vbaGeneration   As String    'Compiled VBA generation reported as text

'------------------------------------------------------------------------------
' RESOLVE COMPILED ENVIRONMENT
'------------------------------------------------------------------------------
    'Describe the running Office/VBA build from compiler constants; do not
    'infer Office bitness from the Windows operating-system description.
#If Win64 Then
        bitness = "64-bit"
#Else
        bitness = "32-bit"
#End If

#If VBA7 Then
        vbaGeneration = "VBA7+"
#Else
        vbaGeneration = "legacy VBA"
#End If

'------------------------------------------------------------------------------
' BUILD ENVIRONMENT RECORD
'------------------------------------------------------------------------------
    'Combine live host properties with the compiled environment fields,
    'keeping the evidence parser keys and order stable.
        EnvironmentSummary = _
            "host=" & Application.Name & _
            "; version=" & Application.Version & _
            "; os=" & Application.OperatingSystem & _
            "; office=" & bitness & _
            "; runtime=" & vbaGeneration

End Function


Private Sub PrintSummary( _
    ByVal cleanupPassed As Boolean, _
    ByVal cleanupDetail As String)
'
'==============================================================================
'                                 PrintSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Report counts, cleanup, failure details, and the complete run verdict.
'
' INPUTS
'   cleanupPassed and cleanupDetail: the outcome of CleanupRun.
'
' STATE OWNERSHIP
'   Read the retained counters and completion flag without resetting them.
'   PASS requires zero failures, successful cleanup, and a complete suite.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Resolve one verdict before writing the machine-readable report.
    Dim verdict   As String    'Final status derived from failures and cleanup

'------------------------------------------------------------------------------
' RESOLVE VERDICT
'------------------------------------------------------------------------------
    'A zero failure count is insufficient: cleanup must succeed and the
    'suite must have reached both required counts before reporting PASS.
        If mFailureCount = 0 And cleanupPassed And mSuiteCompleted Then
            verdict = "PASS"
        Else
            verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Write the retained counts and cleanup first, then any diagnostics and
    'the final result record consumed by the evidence validator.
        Debug.Print "CASES=" & CStr(mCaseCount)
        Debug.Print "ASSERTIONS=" & CStr(mAssertionCount)
        Debug.Print "FAILURES=" & CStr(mFailureCount)
        Debug.Print "CLEANUP=" & IIf(cleanupPassed, "PASS", "FAIL") & _
            "; detail=" & cleanupDetail

    'Omit the diagnostic record on a clean run; retain all details on failure.
        If Len(mFailureDetails) > 0 Then
            Debug.Print "FAILURE_DETAILS=" & mFailureDetails
        End If

        Debug.Print _
            "RESULT=" & verdict & _
            "; completeness=" & IIf(mSuiteCompleted, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & CStr(mCaseCount) & _
            "; assertions=" & CStr(mAssertionCount) & _
            "; failures=" & CStr(mFailureCount) & _
            "; cleanup=" & IIf(cleanupPassed, "PASS", "FAIL")

End Sub


'
'------------------------------------------------------------------------------
'
'                              OWNED STATE RESET
'
'------------------------------------------------------------------------------
'

Private Sub ResetRun()
'
'==============================================================================
'                                   ResetRun
'------------------------------------------------------------------------------
' PURPOSE
'   Return all harness-owned state to its idle baseline.
'
' STATE OWNERSHIP
'   Clear counters, failure text, active-run state, and completion.
'   This helper never changes Excel application properties.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Clear all harness-owned state together so no count, diagnostic, or
    'completion flag can leak from a previous execution into the next one.
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mFailureDetails = vbNullString
        mRunActive = False
        mSuiteCompleted = False

End Sub
