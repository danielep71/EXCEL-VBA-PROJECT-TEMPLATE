Attribute VB_Name = "ProgressTests"
'==============================================================================
' MODULE: ProgressTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run a deterministic regression suite for the ui-component profile's
'   progress session and report complete evidence to the Immediate window.
'
' PUBLIC SURFACE
'   RunProgressTests is the ui-component suite's entry point.
'   ResetProgressTests is a project-private recovery command for an
'   interrupted run; Option Private Module keeps both procedures out of the
'   external workbook automation API.
'
' DEPENDENCIES
'   ProgressFacade and the built-in VBA/Excel object models only. No external
'   references, workbook fixture, worksheet, donor project, or test framework.
'
' STATE OWNERSHIP
'   Owns private counters, failure text, suite-completeness state, a re-entry
'   flag, and the owned UI properties found at suite start. Cases set caller
'   values and simulate stuck state on purpose; the runner restores the
'   suite-start values before the harness verifies cleanup.
'
' ERROR POLICY
'   Expected facade errors are captured and asserted. A failed case ends any
'   session it left active. Unexpected runner errors are recorded, cleanup
'   runs, and the original number/source/description is re-raised.
'   Assertion failures raise one test-suite error after reporting.
'
' WORKSHEET SAFETY
'   Changes only the status bar, cursor, screen updating and Esc handling,
'   through the component or to set up caller state, and restores them. It
'   never creates, selects, activates, edits, calculates, saves, or closes
'   workbook or worksheet state.
'
' TEST SEAM
'   Reads the live Application properties before, during and after each
'   session. TestCleanupDetection is the negative control: it breaks a
'   restored property and requires the check to report it.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS. Cursor assertions run on Windows only,
'   so the expected assertion count is selected by conditional compilation.
'
' USAGE
'   Import the ui-component production modules first, then run
'   ProgressTests.RunProgressTests from the VBE Immediate window. The status
'   bar flickers while it runs. ProjectTests keeps covering the shared starter.
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
        Private Const TEST_ERROR_DIRTY_START           As Long = vbObjectError + 2260    'Refused re-entry or interrupted run
        Private Const TEST_ERROR_FAILURES              As Long = vbObjectError + 2261    'Failed or incomplete suite outcome
        Private Const EXPECTED_CASES                   As Long = 12                      'Required cases for a complete run
#If Mac Then
        Private Const EXPECTED_ASSERTIONS              As Long = 75                      'Required assertions without cursor checks
#Else
        Private Const EXPECTED_ASSERTIONS              As Long = 79                      'Required assertions with four cursor checks
#End If

    'Number the failing operations shared by AssertRaised and its helpers.
        Private Const OPERATION_BEGIN_WHILE_ACTIVE     As Long = 1                       'Second begin during a session
        Private Const OPERATION_UPDATE_INACTIVE        As Long = 2                       'Update without a session
        Private Const OPERATION_CANCEL_INACTIVE        As Long = 3                       'Cancel request without a session
        Private Const OPERATION_UPDATE_CANCELLED       As Long = 4                       'Update after a cancel request
        Private Const OPERATION_BEGIN_EMPTY_CAPTION    As Long = 5                       'Begin with a blank caption
        Private Const OPERATION_BEGIN_NEGATIVE_TOTAL   As Long = 6                       'Begin with a negative total
        Private Const OPERATION_BEGIN_UNKNOWN_OPTION   As Long = 7                       'Begin with an undefined option flag
        Private Const OPERATION_UPDATE_ABOVE_TOTAL     As Long = 8                       'Update beyond the total
        Private Const OPERATION_UPDATE_NEGATIVE        As Long = 9                       'Update with a negative count

'------------------------------------------------------------------------------
' MODULE TYPES
'------------------------------------------------------------------------------
    'Record the five Application properties a progress session owns.
        Private Type HostUi
            StatusBar          As Variant              'Text, or False when Excel owns the bar
            DisplayStatusBar   As Boolean              'Status-bar visibility
            Cursor             As XlMousePointer       'Mouse pointer; not read on macOS
            ScreenUpdating     As Boolean              'Screen-updating setting
            EnableCancelKey    As XlEnableCancelKey    'Esc/Ctrl+Break handling
        End Type

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Own counters and diagnostics for one run; cleanup releases the active
    'flag while report values remain available until the next reset.
        Private mCaseCount         As Long       'Cases started in the current run
        Private mAssertionCount    As Long       'Assertions evaluated in the current run
        Private mFailureCount      As Long       'Failures recorded, including runner errors
        Private mFailureDetails    As String     'Diagnostics retained until the next reset
        Private mRunActive         As Boolean    'Re-entry guard; cleared by cleanup or reset
        Private mSuiteCompleted    As Boolean    'True only when both expected counts match
        Private mSuiteUi           As HostUi     'Owned UI properties found at suite start
        Private mSuiteUiCaptured   As Boolean    'True once mSuiteUi holds real values


'
'------------------------------------------------------------------------------
'
'                              TEST ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunProgressTests()
'
'==============================================================================
'                               RunProgressTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every progress case and report its assertions with cleanup evidence.
'
' USAGE
'   Run ProgressTests.RunProgressTests from the VBE Immediate window.
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
    Dim uiDetail                As String           'Owned-UI restoration diagnostic, if any
    Dim uiRestored              As Boolean          'True when the suite-start UI was restored

'------------------------------------------------------------------------------
' GUARD ENTRY
'------------------------------------------------------------------------------
    'Refuse an active run before clearing its counters or failure details.
    'An interrupted execution must be reset explicitly by the caller.
        If mRunActive Then
            Debug.Print "RESULT=FAIL_DIRTY_START; cleanup=NOT_RUN"
            Err.Raise _
                TEST_ERROR_DIRTY_START, _
                "ProgressTests.RunProgressTests", _
                "A ProgressTests run is already active. Run ResetProgressTests after an interrupted execution."
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
        mSuiteUi = CaptureHostUi()
        mSuiteUiCaptured = True
        TestBegin
        TestUpdate
        TestRestoresCallerState
        TestKeepScreenUpdating
        TestReentrancy
        TestInactive
        TestCancel
        TestInvalidArguments
        TestElapsed
        TestRepeatedSessions
        TestRecover
        TestCleanupDetection

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
        uiRestored = RestoreSuiteUi(uiDetail)
        cleanupPassed = CleanupRun( _
            cleanupDetail, _
            stateSnapshotValid, _
            initialCalculation, _
            initialDisplayAlerts, _
            initialEnableEvents, _
            initialScreenUpdating)

    'Give the owned UI back before the harness compares host state; a
    'failed restoration fails cleanup and is named in its detail.
        If Not uiRestored Then
            cleanupPassed = False
            cleanupDetail = cleanupDetail & "; " & uiDetail
        End If
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
                "ProgressTests.RunProgressTests", _
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


Public Sub ResetProgressTests()
'
'==============================================================================
'                              ResetProgressTests
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
        Debug.Print "PROGRESS TESTS RESET; run_active=no; counters=0"

End Sub


'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Private Sub TestBegin()
'
'==============================================================================
'                                  TestBegin
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a session takes each owned property with its documented value.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Read the live Application properties inside the session; the component
    'must have changed exactly the ones it owns.
        On Error GoTo CaseFailed

        BeginCase "session.begin"
        ProgressFacade.ProgressBegin "Working", 10
        AssertEqualBoolean "Begin makes the session active", True, ProgressFacade.ProgressIsActive()
        AssertEqualString "Begin shows the first text", "Working: 0% (0 of 10)", CStr(Application.StatusBar)
        AssertEqualBoolean "Begin shows the status bar", True, Application.DisplayStatusBar
        AssertEqualBoolean "Begin turns screen updating off", False, Application.ScreenUpdating
        AssertEqualLong "Begin routes Esc to the error handler", xlErrorHandler, Application.EnableCancelKey
#If Not Mac Then
        AssertEqualLong "Begin sets the wait cursor", xlWait, Application.Cursor
#End If
        ProgressFacade.ProgressEnd
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.begin"
        EndLeakedSession

End Sub


Private Sub TestUpdate()
'
'==============================================================================
'                                  TestUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Check determinate, detailed and indeterminate progress text.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Percentages round down, the detail follows a dash, and a zero total
    'reports a plain count.
        On Error GoTo CaseFailed

        BeginCase "session.update"
        ProgressFacade.ProgressBegin "Copy", 4
        ProgressFacade.ProgressUpdate 1
        AssertEqualString "Update percentage", "Copy: 25% (1 of 4)", CStr(Application.StatusBar)
        ProgressFacade.ProgressUpdate 3, "sheet C"
        AssertEqualString "Update detail", "Copy: 75% (3 of 4) - sheet C", CStr(Application.StatusBar)
        AssertEqualString "ProgressText matches the status bar", CStr(Application.StatusBar), ProgressFacade.ProgressText()
        ProgressFacade.ProgressEnd

        ProgressFacade.ProgressBegin "Scan"
        ProgressFacade.ProgressUpdate 7
        AssertEqualString "Update without a total", "Scan: 7 done", CStr(Application.StatusBar)
        ProgressFacade.ProgressEnd
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.update"
        EndLeakedSession

End Sub


Private Sub TestRestoresCallerState()
'
'==============================================================================
'                           TestRestoresCallerState
'------------------------------------------------------------------------------
' PURPOSE
'   Check that ending a session restores caller-set values, not defaults.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'   The caller values set here are undone by the suite's final restore.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Distinguish the caller's own values from the ones found at suite start.
    Dim callerUi   As HostUi     'Owned properties after the caller set them
    Dim ended      As Boolean    'Result of ProgressEnd

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Use non-default caller values so restoring Excel defaults instead of
    'the snapshot cannot pass.
        On Error GoTo CaseFailed

        BeginCase "session.restores-caller-state"
        Application.StatusBar = "Caller status"
#If Not Mac Then
        Application.Cursor = xlNorthwestArrow
#End If
        callerUi = CaptureHostUi()
        ProgressFacade.ProgressBegin "Job", 2
        ProgressFacade.ProgressUpdate 1
        ended = ProgressFacade.ProgressEnd()
        AssertEqualBoolean "End reports the ended session", True, ended
        AssertEqualString "End restores every owned property", "", HostUiDifference(callerUi)
        AssertEqualString "End restores the caller's status text", "Caller status", CStr(Application.StatusBar)
        AssertEqualBoolean "End leaves no session active", False, ProgressFacade.ProgressIsActive()
#If Not Mac Then
        AssertEqualLong "End restores the caller's cursor", xlNorthwestArrow, Application.Cursor
#End If
        RestoreHostUi mSuiteUi
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.restores-caller-state"
        EndLeakedSession

End Sub


Private Sub TestKeepScreenUpdating()
'
'==============================================================================
'                            TestKeepScreenUpdating
'------------------------------------------------------------------------------
' PURPOSE
'   Check that ProgressKeepScreenUpdating leaves screen updating as found.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Compare against the value found, whatever the caller had set.
    Dim callerUi   As HostUi    'Owned properties before the session

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The option is for callers that must keep repainting during the work.
        On Error GoTo CaseFailed

        BeginCase "session.keep-screen-updating"
        callerUi = CaptureHostUi()
        ProgressFacade.ProgressBegin "Keep", 0, ProgressKeepScreenUpdating
        AssertEqualBoolean "Option keeps screen updating", callerUi.ScreenUpdating, Application.ScreenUpdating
        ProgressFacade.ProgressEnd
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.keep-screen-updating"
        EndLeakedSession

End Sub


Private Sub TestReentrancy()
'
'==============================================================================
'                                TestReentrancy
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a second begin is refused and leaves the first session intact.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A nested begin would overwrite the first snapshot; the refusal must
    'not disturb the outer session's text or state.
        On Error GoTo CaseFailed

        BeginCase "session.reentrancy"
        ProgressFacade.ProgressBegin "Outer"
        AssertRaised "session.reentrancy.begin", OPERATION_BEGIN_WHILE_ACTIVE
        AssertEqualString "Refused begin keeps the outer text", "Outer: 0 done", CStr(Application.StatusBar)
        AssertEqualBoolean "Refused begin keeps the session active", True, ProgressFacade.ProgressIsActive()
        ProgressFacade.ProgressEnd
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.reentrancy"
        EndLeakedSession

End Sub


Private Sub TestInactive()
'
'==============================================================================
'                                 TestInactive
'------------------------------------------------------------------------------
' PURPOSE
'   Check the contract of calls made without an active session.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Prove that the refused calls changed nothing.
    Dim callerUi   As HostUi    'Owned properties before the calls

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Update and cancel need a session; end without one is a no-op so it is
    'safe in every cleanup path.
        On Error GoTo CaseFailed

        BeginCase "session.inactive"
        callerUi = CaptureHostUi()
        AssertRaised "session.inactive.update", OPERATION_UPDATE_INACTIVE
        AssertRaised "session.inactive.cancel", OPERATION_CANCEL_INACTIVE
        AssertEqualBoolean "End without a session returns False", False, ProgressFacade.ProgressEnd()
        AssertEqualString "Inactive calls change nothing", "", HostUiDifference(callerUi)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.inactive"
        EndLeakedSession

End Sub


Private Sub TestCancel()
'
'==============================================================================
'                                  TestCancel
'------------------------------------------------------------------------------
' PURPOSE
'   Check the cancel request, the cancelled update, and cleanup afterwards.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Prove that a cancelled session still restores the caller's state.
    Dim callerUi   As HostUi    'Owned properties before the session

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'RequestCancel stands in for the Esc handler: the next update must raise
    'the cancellation error, and the session stays active until ended.
        On Error GoTo CaseFailed

        BeginCase "session.cancel"
        callerUi = CaptureHostUi()
        ProgressFacade.ProgressBegin "Job", 3
        AssertEqualBoolean "No cancel request at begin", False, ProgressFacade.ProgressIsCancelled()
        ProgressFacade.ProgressRequestCancel
        AssertEqualBoolean "Cancel request recorded", True, ProgressFacade.ProgressIsCancelled()
        AssertRaised "session.cancel.update", OPERATION_UPDATE_CANCELLED
        AssertEqualBoolean "Cancelled session stays active until ended", True, ProgressFacade.ProgressIsActive()
        ProgressFacade.ProgressEnd
        AssertEqualBoolean "End clears the cancel request", False, ProgressFacade.ProgressIsCancelled()
        AssertEqualString "Cancelled session restores every owned property", "", HostUiDifference(callerUi)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.cancel"
        EndLeakedSession

End Sub


Private Sub TestInvalidArguments()
'
'==============================================================================
'                             TestInvalidArguments
'------------------------------------------------------------------------------
' PURPOSE
'   Check each rejected argument and that a rejected begin changes nothing.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Prove that begin validates before it takes any state.
    Dim callerUi   As HostUi    'Owned properties before the rejected calls

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Begin rejections must leave no session and no changed property; update
    'rejections leave the session active for the caller to end.
        On Error GoTo CaseFailed

        BeginCase "session.invalid-arguments"
        callerUi = CaptureHostUi()
        AssertRaised "session.invalid.caption", OPERATION_BEGIN_EMPTY_CAPTION
        AssertRaised "session.invalid.total", OPERATION_BEGIN_NEGATIVE_TOTAL
        AssertRaised "session.invalid.options", OPERATION_BEGIN_UNKNOWN_OPTION
        AssertEqualBoolean "Rejected begin leaves no session", False, ProgressFacade.ProgressIsActive()
        AssertEqualString "Rejected begin changes nothing", "", HostUiDifference(callerUi)
        ProgressFacade.ProgressBegin "Job", 2
        AssertRaised "session.invalid.above-total", OPERATION_UPDATE_ABOVE_TOTAL
        AssertRaised "session.invalid.negative", OPERATION_UPDATE_NEGATIVE
        ProgressFacade.ProgressEnd
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.invalid-arguments"
        EndLeakedSession

End Sub


Private Sub TestElapsed()
'
'==============================================================================
'                                 TestElapsed
'------------------------------------------------------------------------------
' PURPOSE
'   Check the elapsed-time seam: zero without a session, never negative,
'   never decreasing, and shown in the text when requested.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep two readings to prove the clock is monotonic.
    Dim firstReading    As Double    'Elapsed seconds at the first read
    Dim secondReading   As Double    'Elapsed seconds at the later read

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Exact durations depend on the host, so only the documented bounds and
    'the text suffix are asserted.
        On Error GoTo CaseFailed

        BeginCase "session.elapsed"
        AssertEqualBoolean "Elapsed is zero without a session", True, (ProgressFacade.ProgressElapsedSeconds() = 0#)
        ProgressFacade.ProgressBegin "Timed", 0, ProgressShowElapsed
        firstReading = ProgressFacade.ProgressElapsedSeconds()
        secondReading = ProgressFacade.ProgressElapsedSeconds()
        AssertEqualBoolean "Elapsed is never negative", True, (firstReading >= 0#)
        AssertEqualBoolean "Elapsed never decreases", True, (secondReading >= firstReading)
        AssertEqualBoolean "Elapsed appears in the text", True, (Right$(ProgressFacade.ProgressText(), 2) = " s")
        ProgressFacade.ProgressEnd
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.elapsed"
        EndLeakedSession

End Sub


Private Sub TestRepeatedSessions()
'
'==============================================================================
'                             TestRepeatedSessions
'------------------------------------------------------------------------------
' PURPOSE
'   Check that consecutive sessions each leave the owned state unchanged.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Compare each cycle against the state found before the first one.
    Dim callerUi   As HostUi    'Owned properties before the sessions
    Dim cycle      As Long      'Session cycle, 1 to 2

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A second cycle would expose state left behind by the first.
        On Error GoTo CaseFailed

        BeginCase "session.repeat"
        callerUi = CaptureHostUi()
        For cycle = 1 To 2
            ProgressFacade.ProgressBegin "Cycle", 1
            ProgressFacade.ProgressUpdate 1
            ProgressFacade.ProgressEnd
            AssertEqualString "Cycle " & CStr(cycle) & " restores every owned property", "", HostUiDifference(callerUi)
        Next cycle
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.repeat"
        EndLeakedSession

End Sub


Private Sub TestRecover()
'
'==============================================================================
'                                 TestRecover
'------------------------------------------------------------------------------
' PURPOSE
'   Check recovery with a live session and after the snapshot was lost.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'   The stuck values set here are undone by the suite's final restore.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Recovery with a session must use the snapshot, not defaults.
    Dim callerUi   As HostUi    'Owned properties before the live session

'------------------------------------------------------------------------------
' RECOVER A LIVE SESSION
'------------------------------------------------------------------------------
    'With a session active, recovery is exactly an end.
        On Error GoTo CaseFailed

        BeginCase "session.recover"
        callerUi = CaptureHostUi()
        ProgressFacade.ProgressBegin "Stuck", 5
        ProgressFacade.ProgressRecover
        AssertEqualBoolean "Recover ends a live session", False, ProgressFacade.ProgressIsActive()
        AssertEqualString "Recover restores the snapshot", "", HostUiDifference(callerUi)

'------------------------------------------------------------------------------
' RECOVER AFTER A LOST SNAPSHOT
'------------------------------------------------------------------------------
    'Simulate the state an interrupted macro leaves behind once a VBA reset
    'has discarded the session; recovery must fall back to Excel defaults.
        Application.StatusBar = "stuck"
        Application.ScreenUpdating = False
        Application.EnableCancelKey = xlDisabled
#If Not Mac Then
        Application.Cursor = xlWait
#End If
        ProgressFacade.ProgressRecover
        AssertEqualBoolean "Recover returns the status bar to Excel", True, (VarType(Application.StatusBar) = vbBoolean)
        AssertEqualBoolean "Recover turns screen updating on", True, Application.ScreenUpdating
        AssertEqualLong "Recover restores Esc interruption", xlInterrupt, Application.EnableCancelKey
#If Not Mac Then
        AssertEqualLong "Recover restores the default cursor", xlDefault, Application.Cursor
#End If
        RestoreHostUi mSuiteUi
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "session.recover"
        EndLeakedSession

End Sub


Private Sub TestCleanupDetection()
'
'==============================================================================
'                             TestCleanupDetection
'------------------------------------------------------------------------------
' PURPOSE
'   Negative control: prove that the restoration check fails when a
'   restored property is deliberately broken, and passes once repaired.
'
' ERROR POLICY
'   Record unexpected case errors, end any leaked session, and continue.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the reference state and the check's verdict on the broken state.
    Dim callerUi     As HostUi    'Owned properties before the session
    Dim difference   As String    'HostUiDifference result for the broken state

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Break a property after a correct end, exactly as a restoration defect
    'would leave it; every other case relies on this check detecting it.
        On Error GoTo CaseFailed

        BeginCase "negative.cleanup-detection"
        callerUi = CaptureHostUi()
        ProgressFacade.ProgressBegin "Probe"
        ProgressFacade.ProgressEnd
        Application.StatusBar = "broken restoration"
        difference = HostUiDifference(callerUi)
        AssertEqualBoolean "Broken status bar is detected", True, (Len(difference) > 0)
        AssertEqualBoolean "Detection names the status bar", True, (InStr(1, difference, "StatusBar", vbBinaryCompare) > 0)
        RestoreHostUi callerUi
        AssertEqualString "Repaired state passes the check", "", HostUiDifference(callerUi)
#If Not Mac Then
        Application.Cursor = xlIBeam
        AssertEqualBoolean "Broken cursor is detected", True, (InStr(1, HostUiDifference(callerUi), "Cursor", vbBinaryCompare) > 0)
        RestoreHostUi callerUi
#End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give the Excel state back.
        RecordUnexpectedCaseError "negative.cleanup-detection"
        EndLeakedSession

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


Private Sub AssertRaised( _
    ByVal caseLabel As String, _
    ByVal operation As Long)
'
'==============================================================================
'                                 AssertRaised
'------------------------------------------------------------------------------
' PURPOSE
'   Run one operation that must fail and verify its complete error contract.
'
' INPUTS
'   caseLabel: prefix for the four assertion names.
'   operation: one of the OPERATION_ constants.
'
' SIDE EFFECTS
'   Count four assertions: the call raised, and its number, source and
'   description match ExpectedRaise.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the expected contract and the captured error side by side.
    Dim actualDescription     As String     'Error description captured from the facade
    Dim actualNumber          As Long       'Error number captured from the facade
    Dim actualSource          As String     'Error source captured from the facade
    Dim expectedDescription   As String     'Documented rule for the operation
    Dim expectedNumber        As Long       'Public error constant for the operation
    Dim expectedSource        As String     'Facade procedure expected as the source
    Dim raised                As Boolean    'True when the operation raised

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'A call that returns normally leaves zero and empty values, so the three
    'field assertions fail as well as the raise assertion.
        ExpectedRaise operation, expectedNumber, expectedSource, expectedDescription
        raised = InvokeRaise(operation, actualNumber, actualSource, actualDescription)
        AssertEqualBoolean caseLabel & ".raises", True, raised
        AssertExpectedError _
            caseLabel, _
            expectedNumber, _
            expectedSource, _
            expectedDescription, _
            actualNumber, _
            actualSource, _
            actualDescription

End Sub


Private Sub ExpectedRaise( _
    ByVal operation As Long, _
    ByRef expectedNumber As Long, _
    ByRef expectedSource As String, _
    ByRef expectedDescription As String)
'
'==============================================================================
'                                ExpectedRaise
'------------------------------------------------------------------------------
' PURPOSE
'   State the documented error contract of each failing operation.
'
' RETURNS
'   The expected number, source and description through ByRef arguments.
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
            Case OPERATION_BEGIN_WHILE_ACTIVE
                expectedNumber = ProgressFacade.PROGRESS_ERROR_ACTIVE
                expectedSource = "ProgressFacade.ProgressBegin"
                expectedDescription = "A progress session is already active."
            Case OPERATION_UPDATE_INACTIVE
                expectedNumber = ProgressFacade.PROGRESS_ERROR_INACTIVE
                expectedSource = "ProgressFacade.ProgressUpdate"
                expectedDescription = "No progress session is active."
            Case OPERATION_CANCEL_INACTIVE
                expectedNumber = ProgressFacade.PROGRESS_ERROR_INACTIVE
                expectedSource = "ProgressFacade.ProgressRequestCancel"
                expectedDescription = "No progress session is active."
            Case OPERATION_UPDATE_CANCELLED
                expectedNumber = ProgressFacade.PROGRESS_ERROR_CANCELLED
                expectedSource = "ProgressFacade.ProgressUpdate"
                expectedDescription = "The progress session was cancelled."
            Case OPERATION_BEGIN_EMPTY_CAPTION
                expectedNumber = ProgressFacade.PROGRESS_ERROR_INVALID_ARGUMENT
                expectedSource = "ProgressFacade.ProgressBegin"
                expectedDescription = "caption must not be empty."
            Case OPERATION_BEGIN_NEGATIVE_TOTAL
                expectedNumber = ProgressFacade.PROGRESS_ERROR_INVALID_ARGUMENT
                expectedSource = "ProgressFacade.ProgressBegin"
                expectedDescription = "total must be zero or greater."
            Case OPERATION_BEGIN_UNKNOWN_OPTION
                expectedNumber = ProgressFacade.PROGRESS_ERROR_INVALID_ARGUMENT
                expectedSource = "ProgressFacade.ProgressBegin"
                expectedDescription = "options must combine only ProgressKeepScreenUpdating and ProgressShowElapsed."
            Case Else
                expectedNumber = ProgressFacade.PROGRESS_ERROR_INVALID_ARGUMENT
                expectedSource = "ProgressFacade.ProgressUpdate"
                expectedDescription = "completed must be zero or greater and at most total."
        End Select

End Sub


Private Function InvokeRaise( _
    ByVal operation As Long, _
    ByRef errorNumber As Long, _
    ByRef errorSource As String, _
    ByRef errorDescription As String) _
    As Boolean
'
'==============================================================================
'                                 InvokeRaise
'------------------------------------------------------------------------------
' PURPOSE
'   Perform one failing operation and capture its error.
'
' RETURNS
'   True when the call raised, with the error in the ByRef arguments; False
'   when it returned normally, with zero and empty values.
'
' ERROR POLICY
'   The local handler captures the expected error; leaving the function
'   from the handler clears it before the caller continues.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL OPERATION
'------------------------------------------------------------------------------
    'Clear the outputs first so a call that fails to raise cannot be scored
    'against a previous error.
        errorNumber = 0
        errorSource = vbNullString
        errorDescription = vbNullString
        On Error GoTo Raised

        Select Case operation
            Case OPERATION_BEGIN_WHILE_ACTIVE
                ProgressFacade.ProgressBegin "Inner"
            Case OPERATION_UPDATE_INACTIVE, OPERATION_UPDATE_CANCELLED
                ProgressFacade.ProgressUpdate 1
            Case OPERATION_CANCEL_INACTIVE
                ProgressFacade.ProgressRequestCancel
            Case OPERATION_BEGIN_EMPTY_CAPTION
                ProgressFacade.ProgressBegin "   "
            Case OPERATION_BEGIN_NEGATIVE_TOTAL
                ProgressFacade.ProgressBegin "Job", -1
            Case OPERATION_BEGIN_UNKNOWN_OPTION
                ProgressFacade.ProgressBegin "Job", 0, 8
            Case OPERATION_UPDATE_ABOVE_TOTAL
                ProgressFacade.ProgressUpdate 3
            Case Else
                ProgressFacade.ProgressUpdate -1
        End Select
        InvokeRaise = False
        Exit Function

'------------------------------------------------------------------------------
' CAPTURE ERROR
'------------------------------------------------------------------------------
Raised:
    'Snapshot the error before the handler exits and clears it.
        errorNumber = Err.Number
        errorSource = Err.Source
        errorDescription = Err.Description
        InvokeRaise = True

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
'                                HOST UI STATE
'
'------------------------------------------------------------------------------
'

Private Function CaptureHostUi() _
    As HostUi
'
'==============================================================================
'                                CaptureHostUi
'------------------------------------------------------------------------------
' PURPOSE
'   Read the five Application properties a progress session owns.
'
' RETURNS
'   The current values; the cursor is read on Windows only.
'
' ERROR POLICY
'   Property read failures propagate to the calling case.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Assemble the values in a local before returning the whole record.
    Dim ui   As HostUi    'Values read from Excel

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
    'Read the same properties, in the same way, as the component snapshots.
        ui.StatusBar = Application.StatusBar
        ui.DisplayStatusBar = Application.DisplayStatusBar
#If Not Mac Then
        ui.Cursor = Application.Cursor
#End If
        ui.ScreenUpdating = Application.ScreenUpdating
        ui.EnableCancelKey = Application.EnableCancelKey
        CaptureHostUi = ui

End Function


Private Sub RestoreHostUi( _
    ByRef ui As HostUi)
'
'==============================================================================
'                                RestoreHostUi
'------------------------------------------------------------------------------
' PURPOSE
'   Put the five owned Application properties back to recorded values.
'
' INPUTS
'   ui: values recorded by CaptureHostUi.
'
' ERROR POLICY
'   Property write failures propagate to the caller.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Undo test-made changes independently of the component under test.
        Application.EnableCancelKey = ui.EnableCancelKey
        Application.ScreenUpdating = ui.ScreenUpdating
#If Not Mac Then
        Application.Cursor = ui.Cursor
#End If
        Application.StatusBar = ui.StatusBar
        Application.DisplayStatusBar = ui.DisplayStatusBar

End Sub


Private Function HostUiDifference( _
    ByRef expected As HostUi) _
    As String
'
'==============================================================================
'                               HostUiDifference
'------------------------------------------------------------------------------
' PURPOSE
'   Name every owned property whose live value differs from a record.
'
' INPUTS
'   expected: values recorded by CaptureHostUi.
'
' RETURNS
'   A comma-separated list of property names, or empty text when all match.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Compare against a fresh reading taken once.
    Dim actual       As HostUi    'Current values read from Excel
    Dim difference   As String    'Names of the properties that differ

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Name each property separately so a report shows exactly what leaked.
        actual = CaptureHostUi()
        If Not SameStatusBar(expected.StatusBar, actual.StatusBar) Then difference = difference & ", StatusBar"
        If expected.DisplayStatusBar <> actual.DisplayStatusBar Then difference = difference & ", DisplayStatusBar"
#If Not Mac Then
        If expected.Cursor <> actual.Cursor Then difference = difference & ", Cursor"
#End If
        If expected.ScreenUpdating <> actual.ScreenUpdating Then difference = difference & ", ScreenUpdating"
        If expected.EnableCancelKey <> actual.EnableCancelKey Then difference = difference & ", EnableCancelKey"
        If Len(difference) > 0 Then difference = Mid$(difference, 3)
        HostUiDifference = difference

End Function


Private Function SameStatusBar( _
    ByVal expected As Variant, _
    ByVal actual As Variant) _
    As Boolean
'
'==============================================================================
'                                SameStatusBar
'------------------------------------------------------------------------------
' PURPOSE
'   Compare two status-bar values, which are text or False.
'
' RETURNS
'   True when both are the same text, or both are False (Excel owns the bar).
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Compare types first: VBA would raise a type mismatch comparing False
    'with text, and And does not short-circuit.
        If VarType(expected) <> VarType(actual) Then
            SameStatusBar = False
        ElseIf VarType(expected) = vbBoolean Then
            SameStatusBar = (expected = actual)
        Else
            SameStatusBar = (StrComp(CStr(expected), CStr(actual), vbBinaryCompare) = 0)
        End If

End Function


Private Sub EndLeakedSession()
'
'==============================================================================
'                               EndLeakedSession
'------------------------------------------------------------------------------
' PURPOSE
'   End a session that a failed case left active.
'
' SIDE EFFECTS
'   Restore the component's snapshot; no effect without a session.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' END SESSION
'------------------------------------------------------------------------------
    'Use the component's own end so its snapshot, not defaults, is applied.
        If ProgressFacade.ProgressIsActive() Then ProgressFacade.ProgressEnd

End Sub


Private Function RestoreSuiteUi( _
    ByRef detail As String) _
    As Boolean
'
'==============================================================================
'                                RestoreSuiteUi
'------------------------------------------------------------------------------
' PURPOSE
'   Give back the owned properties exactly as found when the suite began.
'
' RETURNS
'   True on success; False with a diagnostic in detail otherwise.
'
' ERROR POLICY
'   Contain any error so the harness can still verify cleanup and report.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Cases that set caller values or simulate stuck state rely on this final
    'restore, which also ends any session a failure left behind.
        On Error GoTo RestoreFailed

        EndLeakedSession
        If mSuiteUiCaptured Then RestoreHostUi mSuiteUi
        mSuiteUiCaptured = False
        detail = vbNullString
        RestoreSuiteUi = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE RESTORE ERROR
'------------------------------------------------------------------------------
RestoreFailed:
    'Report the failure; the harness marks cleanup as failed.
        detail = "ui restore error=" & CStr(Err.Number) & "; description=" & Err.Description
        RestoreSuiteUi = False

End Function


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
        Debug.Print "PROGRESS TESTS"
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
        mSuiteUiCaptured = False

End Sub
