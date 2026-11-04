Attribute VB_Name = "LifecycleTests"
'==============================================================================
' MODULE: LifecycleTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run a deterministic regression suite for the application profile's
'   lifecycle reference and report complete evidence to the Immediate window.
'
' PUBLIC SURFACE
'   RunLifecycleTests is the application suite's entry point.
'   ResetLifecycleTests is a project-private recovery command for an
'   interrupted run; Option Private Module keeps both procedures out of the
'   external workbook automation API.
'
' DEPENDENCIES
'   LifecycleFacade, LifecycleCore's fault-injection seam, and the built-in
'   VBA/Excel object models only. No external references, workbook fixture,
'   donor project, or test framework.
'
' STATE OWNERSHIP
'   Owns private counters, failure text, suite-completeness state, a re-entry
'   flag, the temporary host workbooks it creates, and the workbook count and
'   active workbook found at suite start. Every case stops the lifecycle,
'   disarms the seam and closes its hosts without saving, even when it fails.
'
' ERROR POLICY
'   Expected facade errors are captured and asserted. Unexpected runner
'   errors are recorded, cleanup runs, and the original
'   number/source/description is re-raised. Assertion failures raise one
'   test-suite error after reporting. A lifecycle session already running
'   at suite start refuses the run before anything changes, so every
'   session the harness stops is one its own cases started.
'
' WORKSHEET SAFETY
'   Creates temporary workbooks with Workbooks.Add, gives activation back to
'   the workbook that was active at suite start, writes only to the
'   temporary workbooks, and closes them without saving. It never edits,
'   calculates, saves or closes any other workbook.
'
' TEST SEAM
'   LifecycleCore.SetFault makes a start step fail or re-enter. Two negative
'   controls prove that a leaked session, an open host and a failed recovery
'   are reported by the checks the other cases rely on.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS. Run it with a workbook active, such as
'   the one holding the VBA project, so activation can be given back, and
'   with no lifecycle session running (call LifecycleStop first).
'
' USAGE
'   Import the application production modules first, then run
'   LifecycleTests.RunLifecycleTests from the VBE Immediate window. Excel
'   briefly opens and closes temporary workbooks. ProjectTests keeps
'   covering the shared starter.
'
' UPDATED
'   2026-11-04
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
        Private Const TEST_ERROR_DIRTY_START          As Long = vbObjectError + 2360    'Refused re-entry or interrupted run
        Private Const TEST_ERROR_FAILURES             As Long = vbObjectError + 2361    'Failed or incomplete suite outcome
        Private Const TEST_ERROR_SESSION_RUNNING      As Long = vbObjectError + 2362    'A session was running at suite start
        Private Const TEST_ERROR_ACTIVATION           As Long = vbObjectError + 2363    'Activation could not be given back
        Private Const EXPECTED_CASES                  As Long = 16                      'Required cases for a complete run
        Private Const EXPECTED_ASSERTIONS             As Long = 106                     'Required assertions for a complete run

    'Number the failing operations shared by AssertRaised and its helpers.
        Private Const OPERATION_START_NO_HOST         As Long = 1                       'Start without a host
        Private Const OPERATION_START_OTHER_HOST      As Long = 2                       'Start while another host runs
        Private Const OPERATION_START_NEWER           As Long = 3                       'Start over a newer schema
        Private Const OPERATION_START_INVALID         As Long = 4                       'Start over unreadable state
        Private Const OPERATION_START_UNCLEAN         As Long = 5                       'Start after a session that did not stop
        Private Const OPERATION_WRITE_NOT_RUNNING     As Long = 6                       'Status write without a session
        Private Const OPERATION_WRITE_NO_TARGET       As Long = 7                       'Status write without a target
        Private Const OPERATION_RESET_RUNNING         As Long = 8                       'Reset of the running host
        Private Const OPERATION_START_FAULT_MIGRATE   As Long = 9                       'Start failing in migration 2
        Private Const OPERATION_START_FAULT_STATUS    As Long = 10                      'Start failing at the status step
        Private Const OPERATION_RECOVER_NO_HOST       As Long = 11                      'Recover without a host
        Private Const OPERATION_RECOVER_INVALID       As Long = 12                      'Recover over unreadable state

    'Name the stored values the cases seed and expect.
        Private Const MISSING_NAME                    As String = "(missing)"           'Signature text for an absent name
        Private Const STORED_RUNNING                  As String = "=""running"""        'RefersTo of a running status
        Private Const STORED_STOPPED                  As String = "=""stopped"""        'RefersTo of a stopped status

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Own counters and diagnostics for one run; cleanup releases the active
    'flag while report values remain available until the next reset.
        Private mAssertionCount     As Long          'Assertions evaluated in the current run
        Private mCaseCount          As Long          'Cases started in the current run
        Private mFailureCount       As Long          'Failures recorded, including runner errors
        Private mFailureDetails     As String        'Diagnostics retained until the next reset
        Private mHosts              As Collection    'Temporary host workbooks still open
        Private mRunActive          As Boolean       'Re-entry guard; cleared by cleanup or reset
        Private mSuiteActiveBook    As Workbook      'Workbook active at suite start
        Private mSuiteBookCount     As Long          'Open workbooks at suite start
        Private mSuiteCaptured      As Boolean       'True once the two suite values are real
        Private mSuiteCompleted     As Boolean       'True only when both expected counts match


'
'------------------------------------------------------------------------------
'
'                              TEST ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunLifecycleTests()
'
'==============================================================================
'                              RunLifecycleTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run every lifecycle case and report its assertions with cleanup evidence.
'
' USAGE
'   Run LifecycleTests.RunLifecycleTests from the VBE Immediate window.
'
' STATE OWNERSHIP
'   Reject an active run before resetting counters. Snapshot four Excel
'   properties, the open-workbook count and the active workbook. Cases open
'   and close temporary hosts; cleanup verifies that none is left open, no
'   session is left running and the active workbook is the one found.
'
' ERROR POLICY
'   Record unexpected runner failures, attempt cleanup, print the summary,
'   then re-raise the saved error. Other failed outcomes raise the suite error.
'   A dirty start, or a session already running, raises immediately and
'   does not run cleanup.
'
' UPDATED
'   2026-11-04
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
                "LifecycleTests.RunLifecycleTests", _
                "A LifecycleTests run is already active. Run ResetLifecycleTests after an interrupted execution."
        End If

    'The cases need the single session slot. Refuse rather than stop a
    'session that belongs to another workbook, such as ThisWorkbook.
        If LifecycleFacade.LifecycleIsRunning() Then
            Debug.Print "RESULT=FAIL_SESSION_RUNNING; cleanup=NOT_RUN"
            Err.Raise _
                TEST_ERROR_SESSION_RUNNING, _
                "LifecycleTests.RunLifecycleTests", _
                "A lifecycle session is already running. Call LifecycleFacade.LifecycleStop, then run the suite."
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

    'Record what the cases must give back: the open workbooks and the active one.
        Set mHosts = New Collection
        mSuiteBookCount = Application.Workbooks.Count
        Set mSuiteActiveBook = Application.ActiveWorkbook
        mSuiteCaptured = True

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run cases in the evidence contract order; each case records unexpected
    'errors and closes its hosts so later cases still contribute.
        PrintEnvironment
        TestFreshStart
        TestRepeatStart
        TestRestart
        TestUpgradeFromSchema1
        TestCurrentSchemaNoOp
        TestDiscoveryRefusals
        TestPartialStartFresh
        TestPartialStartUpgrade
        TestReentryRefused
        TestSecondWorkbook
        TestRecoverUncleanStop
        TestRecoverLiveSession
        TestReset
        TestExplicitTarget
        TestCleanupDetection
        TestFailedRecoveryDetection

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
                "LifecycleTests.RunLifecycleTests", _
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


Public Sub ResetLifecycleTests()
'
'==============================================================================
'                             ResetLifecycleTests
'------------------------------------------------------------------------------
' PURPOSE
'   Recover harness-owned state after an interrupted execution.
'
' USAGE
'   Run only after the previous execution has stopped; then rerun the suite.
'
' SIDE EFFECTS
'   Stop a lifecycle session still running, which after an interrupted run
'   is that run's own, disarm the seam, close any temporary host still
'   tracked, and clear counters, failure text, completion, and the
'   active-run flag. Print confirmation.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Give back what an interrupted run may have left, then clear the
    'harness state; the confirmation distinguishes an explicit reset.
        ReleaseCaseState
        ResetRun
        Debug.Print "LIFECYCLE TESTS RESET; run_active=no; counters=0"

End Sub


'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Private Sub TestFreshStart()
'
'==============================================================================
'                                TestFreshStart
'------------------------------------------------------------------------------
' PURPOSE
'   Check a first start on an empty host, then a clean stop.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and both lifecycle results.
    Dim host      As Workbook    'Temporary host for this case
    Dim started   As Boolean     'Result of LifecycleStart
    Dim stopped   As Boolean     'Result of LifecycleStop

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A fresh install runs both migrations, so the stored names are read
    'directly as well as through the facade.
        On Error GoTo CaseFailed

        BeginCase "lifecycle.fresh-start"
        Set host = CreateHost()
        started = LifecycleFacade.LifecycleStart(host)
        AssertEqualBoolean "Start reports a new session", True, started
        AssertEqualBoolean "Start leaves the session running", True, LifecycleFacade.LifecycleIsRunning()
        AssertEqualString "Status reports running", "running", LifecycleFacade.LifecycleStatus()
        AssertEqualLong _
            "Start writes the current schema", _
            LifecycleFacade.LIFECYCLE_SCHEMA_VERSION, _
            LifecycleFacade.LifecycleStoredSchema(host)
        AssertEqualLong "Start counts the first start", 1, LifecycleFacade.LifecycleStoredStarts(host)
        AssertEqualString "Start stores the running status", STORED_RUNNING, StoredRefersTo(host, "LifecycleStatus")
        stopped = LifecycleFacade.LifecycleStop()
        AssertEqualBoolean "Stop reports the stopped session", True, stopped
        AssertEqualString "Stop stores the stopped status", STORED_STOPPED, StoredRefersTo(host, "LifecycleStatus")
        AssertEqualBoolean "Stop leaves no session", False, LifecycleFacade.LifecycleIsRunning()
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "lifecycle.fresh-start"
        ReleaseCaseState

End Sub


Private Sub TestRepeatStart()
'
'==============================================================================
'                               TestRepeatStart
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a repeated start and a repeated stop change nothing.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and the repeated results.
    Dim host          As Workbook    'Temporary host for this case
    Dim secondStart   As Boolean     'Result of the repeated start
    Dim secondStop    As Boolean     'Result of the repeated stop

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Two event handlers may both start the application; the second call must
    'not count a second start.
        On Error GoTo CaseFailed

        BeginCase "lifecycle.repeat-start"
        Set host = CreateHost()
        LifecycleFacade.LifecycleStart host
        secondStart = LifecycleFacade.LifecycleStart(host)
        AssertEqualBoolean "Repeated start is a no-op", False, secondStart
        AssertEqualLong "Repeated start adds no start", 1, LifecycleFacade.LifecycleStoredStarts(host)
        AssertEqualBoolean "Repeated start keeps the session", True, LifecycleFacade.LifecycleIsRunning()
        LifecycleFacade.LifecycleStop
        secondStop = LifecycleFacade.LifecycleStop()
        AssertEqualBoolean "Stop without a session returns False", False, secondStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "lifecycle.repeat-start"
        ReleaseCaseState

End Sub


Private Sub TestRestart()
'
'==============================================================================
'                                 TestRestart
'------------------------------------------------------------------------------
' PURPOSE
'   Check a second session on the same host: counted, with no migration.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host.
    Dim host   As Workbook    'Temporary host for this case

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The second start finds the current schema, so only the status and the
    'count change and no fourth name appears.
        On Error GoTo CaseFailed

        BeginCase "lifecycle.restart"
        Set host = CreateHost()
        LifecycleFacade.LifecycleStart host
        LifecycleFacade.LifecycleStop
        LifecycleFacade.LifecycleStart host
        AssertEqualLong "Restart counts both starts", 2, LifecycleFacade.LifecycleStoredStarts(host)
        AssertEqualLong _
            "Restart keeps the schema", _
            LifecycleFacade.LIFECYCLE_SCHEMA_VERSION, _
            LifecycleFacade.LifecycleStoredSchema(host)
        AssertEqualLong "Restart owns exactly three names", 3, OwnedNameCount(host)
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "lifecycle.restart"
        ReleaseCaseState

End Sub


Private Sub TestUpgradeFromSchema1()
'
'==============================================================================
'                            TestUpgradeFromSchema1
'------------------------------------------------------------------------------
' PURPOSE
'   Check the migration of schema-1 state written by an earlier version.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and the start result.
    Dim host      As Workbook    'Temporary host seeded with schema 1
    Dim started   As Boolean     'Result of LifecycleStart

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Schema 1 had no start count; migration 2 adds it at zero before the
    'start counts itself.
        On Error GoTo CaseFailed

        BeginCase "upgrade.from-schema-1"
        Set host = CreateHost()
        SeedName host, "LifecycleSchema", "=1"
        SeedName host, "LifecycleStatus", STORED_STOPPED
        started = LifecycleFacade.LifecycleStart(host)
        AssertEqualBoolean "Upgrade start succeeds", True, started
        AssertEqualLong _
            "Upgrade writes the current schema", _
            LifecycleFacade.LIFECYCLE_SCHEMA_VERSION, _
            LifecycleFacade.LifecycleStoredSchema(host)
        AssertEqualLong "Upgrade counts from zero", 1, LifecycleFacade.LifecycleStoredStarts(host)
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "upgrade.from-schema-1"
        ReleaseCaseState

End Sub


Private Sub TestCurrentSchemaNoOp()
'
'==============================================================================
'                            TestCurrentSchemaNoOp
'------------------------------------------------------------------------------
' PURPOSE
'   Check that current-schema state is kept: no migration, count continued.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host.
    Dim host   As Workbook    'Temporary host seeded with schema 2

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A seeded count of five proves the stored value is continued, not reset.
        On Error GoTo CaseFailed

        BeginCase "upgrade.current-schema-no-op"
        Set host = CreateHost()
        SeedName host, "LifecycleSchema", "=2"
        SeedName host, "LifecycleStatus", STORED_STOPPED
        SeedName host, "LifecycleStarts", "=5"
        LifecycleFacade.LifecycleStart host
        AssertEqualLong "Current schema continues the count", 6, LifecycleFacade.LifecycleStoredStarts(host)
        AssertEqualLong "Current schema is unchanged", 2, LifecycleFacade.LifecycleStoredSchema(host)
        AssertEqualLong "Current schema adds no name", 3, OwnedNameCount(host)
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "upgrade.current-schema-no-op"
        ReleaseCaseState

End Sub


Private Sub TestDiscoveryRefusals()
'
'==============================================================================
'                            TestDiscoveryRefusals
'------------------------------------------------------------------------------
' PURPOSE
'   Check that discovery refuses newer, unreadable and missing input before
'   changing anything.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep each host's names before the refused call.
    Dim before     As String      'Owned-name signature before the call
    Dim invalid    As Workbook    'Host with an unreadable schema
    Dim newer      As Workbook    'Host with a newer schema

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A refusal during discovery has taken no ownership, so the host must be
    'byte-for-byte as seeded and no session may exist.
        On Error GoTo CaseFailed

        BeginCase "discovery.refusals"
        Set newer = CreateHost()
        SeedName newer, "LifecycleSchema", "=3"
        SeedName newer, "LifecycleStatus", STORED_STOPPED
        before = NameSignature(newer)
        AssertRaised "Newer schema", OPERATION_START_NEWER, newer
        AssertEqualString "Newer schema leaves the names unchanged", before, NameSignature(newer)
        AssertEqualBoolean "Newer schema starts no session", False, LifecycleFacade.LifecycleIsRunning()

        Set invalid = CreateHost()
        SeedName invalid, "LifecycleSchema", "=""abc"""
        before = NameSignature(invalid)
        AssertRaised "Invalid schema", OPERATION_START_INVALID, invalid
        AssertEqualString "Invalid schema leaves the names unchanged", before, NameSignature(invalid)

        AssertRaised "Start without a host", OPERATION_START_NO_HOST, Nothing
        AssertRaised "Recover without a host", OPERATION_RECOVER_NO_HOST, Nothing
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "discovery.refusals"
        ReleaseCaseState

End Sub


Private Sub TestPartialStartFresh()
'
'==============================================================================
'                            TestPartialStartFresh
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a start failing in migration 2 removes everything it wrote.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and the later clean start.
    Dim host      As Workbook    'Temporary host for this case
    Dim started   As Boolean     'Result of the start after the failure

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Migration 1 has already written two names when migration 2 fails, so
    'rollback must delete them; a later start must then succeed normally.
        On Error GoTo CaseFailed

        BeginCase "startup.partial-failure-fresh"
        Set host = CreateHost()
        AssertRaised "Failed fresh start", OPERATION_START_FAULT_MIGRATE, host
        AssertEqualLong "Rollback removes the written names", 0, OwnedNameCount(host)
        AssertEqualBoolean "Failed start leaves no session", False, LifecycleFacade.LifecycleIsRunning()
        AssertEqualString "Failed start reports stopped", "stopped", LifecycleFacade.LifecycleStatus()
        started = LifecycleFacade.LifecycleStart(host)
        AssertEqualBoolean "A clean start follows the failure", True, started
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "startup.partial-failure-fresh"
        ReleaseCaseState

End Sub


Private Sub TestPartialStartUpgrade()
'
'==============================================================================
'                           TestPartialStartUpgrade
'------------------------------------------------------------------------------
' PURPOSE
'   Check that an upgrade failing after its migration restores schema 1.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the seeded names for comparison.
    Dim before   As String      'Owned-name signature as seeded
    Dim host     As Workbook    'Temporary host seeded with schema 1

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Migration 2 has already rewritten the schema and added the count when
    'the status step fails; rollback must undo both and keep the rest.
        On Error GoTo CaseFailed

        BeginCase "startup.partial-failure-upgrade"
        Set host = CreateHost()
        SeedName host, "LifecycleSchema", "=1"
        SeedName host, "LifecycleStatus", STORED_STOPPED
        before = NameSignature(host)
        AssertRaised "Failed upgrade start", OPERATION_START_FAULT_STATUS, host
        AssertEqualString "Rollback restores schema 1 exactly", before, NameSignature(host)
        AssertEqualBoolean "Failed upgrade leaves no session", False, LifecycleFacade.LifecycleIsRunning()
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "startup.partial-failure-upgrade"
        ReleaseCaseState

End Sub


Private Sub TestReentryRefused()
'
'==============================================================================
'                              TestReentryRefused
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a start re-entered from inside a start step is refused.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and the outer result.
    Dim host      As Workbook    'Temporary host for this case
    Dim started   As Boolean     'Result of the outer start

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The seam calls LifecycleCore.StartApplication during the status step,
    'as an event handler re-entering the application would.
        On Error GoTo CaseFailed

        BeginCase "lifecycle.reentry-refused"
        Set host = CreateHost()
        LifecycleCore.SetFault "status", True
        started = LifecycleFacade.LifecycleStart(host)
        LifecycleCore.ClearFault
        AssertEqualBoolean "The outer start completes", True, started
        AssertEqualLong _
            "The re-entrant start is refused as busy", _
            LifecycleFacade.LIFECYCLE_ERROR_BUSY, _
            LifecycleCore.ProbeError()
        AssertEqualBoolean "The session is running once", True, LifecycleFacade.LifecycleIsRunning()
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "lifecycle.reentry-refused"
        ReleaseCaseState

End Sub


Private Sub TestSecondWorkbook()
'
'==============================================================================
'                              TestSecondWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Check single ownership: a second host cannot start, and the running
'   host cannot be reset.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the owner and the refused host.
    Dim other   As Workbook    'Host that must be refused
    Dim owner   As Workbook    'Host of the running session

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The refused host must stay empty, and the owner must keep running.
        On Error GoTo CaseFailed

        BeginCase "ownership.second-workbook"
        Set owner = CreateHost()
        Set other = CreateHost()
        LifecycleFacade.LifecycleStart owner
        AssertRaised "Start of a second host", OPERATION_START_OTHER_HOST, other
        AssertEqualLong "The refused host stays empty", 0, OwnedNameCount(other)
        AssertEqualBoolean "The owner keeps running", True, LifecycleFacade.LifecycleIsRunning()
        AssertRaised "Reset of the running host", OPERATION_RESET_RUNNING, owner
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "ownership.second-workbook"
        ReleaseCaseState

End Sub


Private Sub TestRecoverUncleanStop()
'
'==============================================================================
'                            TestRecoverUncleanStop
'------------------------------------------------------------------------------
' PURPOSE
'   Check recovery from a session that never stopped (crash or VBA reset).
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and both recovery results.
    Dim host             As Workbook    'Host seeded with a running status
    Dim recovered        As Boolean     'Result of the first recovery
    Dim recoveredAgain   As Boolean     'Result of the second recovery

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A stored running status with no session is exactly what a crash leaves;
    'start must refuse it until recovery records the stop.
        On Error GoTo CaseFailed

        BeginCase "recovery.unclean-stop"
        Set host = CreateHost()
        SeedName host, "LifecycleSchema", "=2"
        SeedName host, "LifecycleStatus", STORED_RUNNING
        SeedName host, "LifecycleStarts", "=3"
        AssertRaised "Start after an unclean stop", OPERATION_START_UNCLEAN, host
        recovered = LifecycleFacade.LifecycleRecover(host)
        AssertEqualBoolean "Recovery reports a repair", True, recovered
        AssertEqualString "Recovery stores the stopped status", STORED_STOPPED, StoredRefersTo(host, "LifecycleStatus")
        recoveredAgain = LifecycleFacade.LifecycleRecover(host)
        AssertEqualBoolean "A second recovery has nothing to do", False, recoveredAgain
        LifecycleFacade.LifecycleStart host
        AssertEqualLong "Start after recovery continues the count", 4, LifecycleFacade.LifecycleStoredStarts(host)
        LifecycleFacade.LifecycleStop
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "recovery.unclean-stop"
        ReleaseCaseState

End Sub


Private Sub TestRecoverLiveSession()
'
'==============================================================================
'                            TestRecoverLiveSession
'------------------------------------------------------------------------------
' PURPOSE
'   Check that recovering the running host stops its session cleanly.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and the recovery result.
    Dim host        As Workbook    'Temporary host for this case
    Dim recovered   As Boolean     'Result of LifecycleRecover

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Recovery of a live session is a normal stop.
        On Error GoTo CaseFailed

        BeginCase "recovery.live-session"
        Set host = CreateHost()
        LifecycleFacade.LifecycleStart host
        recovered = LifecycleFacade.LifecycleRecover(host)
        AssertEqualBoolean "Recovery stops the live session", True, recovered
        AssertEqualBoolean "Recovery leaves no session", False, LifecycleFacade.LifecycleIsRunning()
        AssertEqualString "Recovery stores the stopped status", STORED_STOPPED, StoredRefersTo(host, "LifecycleStatus")
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "recovery.live-session"
        ReleaseCaseState

End Sub


Private Sub TestReset()
'
'==============================================================================
'                                  TestReset
'------------------------------------------------------------------------------
' PURPOSE
'   Check that reset removes the three owned names and nothing else.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and both reset results.
    Dim host           As Workbook    'Temporary host for this case
    Dim removed        As Long        'Names removed by the first reset
    Dim removedAgain   As Long        'Names removed by the second reset

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A second reset finds nothing, which proves the first removed all three.
        On Error GoTo CaseFailed

        BeginCase "reset.removes-state"
        Set host = CreateHost()
        LifecycleFacade.LifecycleStart host
        LifecycleFacade.LifecycleStop
        removed = LifecycleFacade.LifecycleReset(host)
        AssertEqualLong "Reset removes three names", 3, removed
        AssertEqualLong "Reset leaves no owned name", 0, OwnedNameCount(host)
        removedAgain = LifecycleFacade.LifecycleReset(host)
        AssertEqualLong "A second reset removes nothing", 0, removedAgain
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "reset.removes-state"
        ReleaseCaseState

End Sub


Private Sub TestExplicitTarget()
'
'==============================================================================
'                              TestExplicitTarget
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the status block goes to the target passed in, while another
'   workbook stays active, and that its preconditions are enforced.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host, the target and the result.
    Dim host       As Workbook    'Temporary host for this case
    Dim rowCount   As Long        'Result of LifecycleWriteStatus
    Dim target     As Range       'Top-left cell of the status block

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The host is not the active workbook, so code that used ActiveSheet or
    'Selection would write somewhere else and fail these checks.
        On Error GoTo CaseFailed

        BeginCase "operation.explicit-target"
        Set host = CreateHost()
        Set target = host.Worksheets(1).Range("B2")
        LifecycleFacade.LifecycleStart host
        rowCount = LifecycleFacade.LifecycleWriteStatus(target)
        AssertEqualLong "Status block has four rows", 4, rowCount
        AssertEqualString "Status block starts with the schema label", "Schema", CStr(target.Value)
        AssertEqualString _
            "Status block records the schema", _
            CStr(LifecycleFacade.LIFECYCLE_SCHEMA_VERSION), _
            CStr(target.Offset(0, 1).Value)
        AssertEqualString "Status block records the status", "running", CStr(target.Offset(1, 1).Value)
        AssertEqualString "Status block records the starts", "1", CStr(target.Offset(2, 1).Value)
        AssertEqualString "Status block records the host", host.Name, CStr(target.Offset(3, 1).Value)
        AssertEqualBoolean "The active workbook did not change", True, (Application.ActiveWorkbook Is mSuiteActiveBook)
        AssertRaised "Status write without a target", OPERATION_WRITE_NO_TARGET, host
        LifecycleFacade.LifecycleStop
        AssertRaised "Status write without a session", OPERATION_WRITE_NOT_RUNNING, host
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "operation.explicit-target"
        ReleaseCaseState

End Sub


Private Sub TestCleanupDetection()
'
'==============================================================================
'                             TestCleanupDetection
'------------------------------------------------------------------------------
' PURPOSE
'   Negative control: prove that the cleanup check reports a host left open
'   and a session left running, and passes once both are released.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the leaked host.
    Dim host   As Workbook    'Host deliberately left open

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Leave exactly the defects cleanup must catch; every other case relies
    'on HarnessDifference reporting them.
        On Error GoTo CaseFailed

        BeginCase "negative.cleanup-detection"
        AssertEqualString "Clean state passes the check", "", HarnessDifference()
        Set host = CreateHost()
        AssertEqualBoolean "An open host is detected", True, (InStr(1, HarnessDifference(), "workbooks", vbBinaryCompare) > 0)
        LifecycleFacade.LifecycleStart host
        AssertEqualBoolean _
            "A leaked session is detected", _
            True, _
            (InStr(1, HarnessDifference(), "lifecycle", vbBinaryCompare) > 0)
        ReleaseCaseState
        AssertEqualString "Released state passes the check", "", HarnessDifference()
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "negative.cleanup-detection"
        ReleaseCaseState

End Sub


Private Sub TestFailedRecoveryDetection()
'
'==============================================================================
'                         TestFailedRecoveryDetection
'------------------------------------------------------------------------------
' PURPOSE
'   Negative control: prove that state recovery cannot repair is reported,
'   not silently accepted, and is left exactly as found.
'
' ERROR POLICY
'   Record unexpected case errors, release the case state, and continue.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the corrupt names for comparison.
    Dim before   As String      'Owned-name signature as seeded
    Dim host     As Workbook    'Host with corrupt state

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A recovery that "succeeded" here would hide corrupt state from the
    'caller; the documented way out is LifecycleReset.
        On Error GoTo CaseFailed

        BeginCase "negative.failed-recovery"
        Set host = CreateHost()
        SeedName host, "LifecycleSchema", "=""abc"""
        before = NameSignature(host)
        AssertRaised "Recovery of corrupt state", OPERATION_RECOVER_INVALID, host
        AssertEqualString "Failed recovery leaves the names unchanged", before, NameSignature(host)
        ReleaseCaseState
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed and give back what it opened.
        RecordUnexpectedCaseError "negative.failed-recovery"
        ReleaseCaseState

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
'   2026-11-03
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
    ByVal operation As Long, _
    ByVal host As Workbook)
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
'   host:      the workbook the operation uses, or Nothing.
'
' SIDE EFFECTS
'   Count four assertions: the call raised, and its number, source and
'   description match ExpectedRaise.
'
' UPDATED
'   2026-11-03
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
        raised = InvokeRaise(operation, host, actualNumber, actualSource, actualDescription)
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
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DESCRIBE OPERATION
'------------------------------------------------------------------------------
    'Keep the expected texts literal so a changed rule is a visible diff.
        Select Case operation
            Case OPERATION_START_NO_HOST
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_INVALID_ARGUMENT
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "host must be a workbook."
            Case OPERATION_START_OTHER_HOST
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_RUNNING
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "Another workbook owns the running application."
            Case OPERATION_START_NEWER
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_STATE_NEWER
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "Stored schema 3 is newer than the supported schema 2."
            Case OPERATION_START_INVALID
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_STATE_INVALID
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "Stored lifecycle state is invalid: LifecycleSchema."
            Case OPERATION_START_UNCLEAN
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_RECOVERY_REQUIRED
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "The previous session did not stop; run LifecycleRecover first."
            Case OPERATION_WRITE_NOT_RUNNING
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_NOT_RUNNING
                expectedSource = "LifecycleFacade.LifecycleWriteStatus"
                expectedDescription = "The application is not running."
            Case OPERATION_WRITE_NO_TARGET
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_INVALID_ARGUMENT
                expectedSource = "LifecycleFacade.LifecycleWriteStatus"
                expectedDescription = "target must be a cell range."
            Case OPERATION_RESET_RUNNING
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_RUNNING
                expectedSource = "LifecycleFacade.LifecycleReset"
                expectedDescription = "Stop the application before resetting its state."
            Case OPERATION_START_FAULT_MIGRATE
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_START_FAILED
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "Start failed at step migrate-2: Injected fault."
            Case OPERATION_START_FAULT_STATUS
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_START_FAILED
                expectedSource = "LifecycleFacade.LifecycleStart"
                expectedDescription = "Start failed at step status: Injected fault."
            Case OPERATION_RECOVER_NO_HOST
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_INVALID_ARGUMENT
                expectedSource = "LifecycleFacade.LifecycleRecover"
                expectedDescription = "host must be a workbook."
            Case Else
                expectedNumber = LifecycleFacade.LIFECYCLE_ERROR_STATE_INVALID
                expectedSource = "LifecycleFacade.LifecycleRecover"
                expectedDescription = "Stored lifecycle state is invalid: LifecycleSchema."
        End Select

End Sub


Private Function InvokeRaise( _
    ByVal operation As Long, _
    ByVal host As Workbook, _
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
'   from the handler clears it before the caller continues. The fault seam
'   is disarmed on both paths.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Receive results of calls that must not return.
    Dim ignoredCount   As Long       'Discarded Long result
    Dim ignoredFlag    As Boolean    'Discarded Boolean result

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
            Case OPERATION_START_NO_HOST
                ignoredFlag = LifecycleFacade.LifecycleStart(Nothing)
            Case OPERATION_WRITE_NOT_RUNNING
                ignoredCount = LifecycleFacade.LifecycleWriteStatus(host.Worksheets(1).Range("A1"))
            Case OPERATION_WRITE_NO_TARGET
                ignoredCount = LifecycleFacade.LifecycleWriteStatus(Nothing)
            Case OPERATION_RESET_RUNNING
                ignoredCount = LifecycleFacade.LifecycleReset(host)
            Case OPERATION_START_FAULT_MIGRATE
                LifecycleCore.SetFault "migrate-2"
                ignoredFlag = LifecycleFacade.LifecycleStart(host)
            Case OPERATION_START_FAULT_STATUS
                LifecycleCore.SetFault "status"
                ignoredFlag = LifecycleFacade.LifecycleStart(host)
            Case OPERATION_RECOVER_NO_HOST
                ignoredFlag = LifecycleFacade.LifecycleRecover(Nothing)
            Case OPERATION_RECOVER_INVALID
                ignoredFlag = LifecycleFacade.LifecycleRecover(host)
            Case Else
                ignoredFlag = LifecycleFacade.LifecycleStart(host)
        End Select
        LifecycleCore.ClearFault
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
        LifecycleCore.ClearFault
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
'   2026-11-03
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
'   2026-11-03
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
'   2026-11-03
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
'   2026-11-03
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
'                               HOST WORKBOOKS
'
'------------------------------------------------------------------------------
'

Private Function CreateHost() _
    As Workbook
'
'==============================================================================
'                                  CreateHost
'------------------------------------------------------------------------------
' PURPOSE
'   Open a temporary one-sheet host workbook and give activation back.
'
' RETURNS
'   The new workbook, tracked for ReleaseCaseState to close unsaved.
'
' ERROR POLICY
'   Raise the suite error if activation cannot be given back; the case
'   handler then releases the host.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Track the workbook before anything else can fail.
    Dim host   As Workbook    'New temporary host

'------------------------------------------------------------------------------
' CREATE
'------------------------------------------------------------------------------
    'Workbooks.Add activates the new workbook; the component must work on
    'a host that is not active, so activation goes straight back.
        Set host = Application.Workbooks.Add(xlWBATWorksheet)
        mHosts.Add host
        Set CreateHost = host
        If Not RestoreActiveBook() Then
            Err.Raise _
                TEST_ERROR_ACTIVATION, _
                "LifecycleTests.CreateHost", _
                "The workbook active at suite start could not be activated again."
        End If

End Function


Private Sub SeedName( _
    ByVal host As Workbook, _
    ByVal nameText As String, _
    ByVal definition As String)
'
'==============================================================================
'                                   SeedName
'------------------------------------------------------------------------------
' PURPOSE
'   Write stored state as an earlier version, a crash or corruption left it.
'
' INPUTS
'   host: temporary host; nameText: owned name; definition: formula text.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' SEED
'------------------------------------------------------------------------------
    'Hidden, like the names the component writes.
        host.Names.Add Name:=nameText, RefersTo:=definition, Visible:=False

End Sub


Private Function StoredRefersTo( _
    ByVal host As Workbook, _
    ByVal nameText As String) _
    As String
'
'==============================================================================
'                                StoredRefersTo
'------------------------------------------------------------------------------
' PURPOSE
'   Read an owned name's definition independently of the component.
'
' RETURNS
'   The RefersTo text, or MISSING_NAME when the name does not exist.
'
' ERROR POLICY
'   A missing name is an answer, not a failure.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Names(...) raises for a missing name.
        On Error GoTo Missing
        StoredRefersTo = host.Names(nameText).RefersTo
        Exit Function

'------------------------------------------------------------------------------
' HANDLE MISSING NAME
'------------------------------------------------------------------------------
Missing:
    'Leaving the handler clears the lookup error.
        StoredRefersTo = MISSING_NAME

End Function


Private Function OwnedNameCount( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                                OwnedNameCount
'------------------------------------------------------------------------------
' PURPOSE
'   Count how many of the three owned names exist in a host.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Accumulate in a local, not in the function name.
    Dim found   As Long    'Owned names present so far

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Test each name the component may own.
        If StoredRefersTo(host, "LifecycleSchema") <> MISSING_NAME Then found = found + 1
        If StoredRefersTo(host, "LifecycleStatus") <> MISSING_NAME Then found = found + 1
        If StoredRefersTo(host, "LifecycleStarts") <> MISSING_NAME Then found = found + 1
        OwnedNameCount = found

End Function


Private Function NameSignature( _
    ByVal host As Workbook) _
    As String
'
'==============================================================================
'                                NameSignature
'------------------------------------------------------------------------------
' PURPOSE
'   Describe the three owned names exactly, for before/after comparison.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
    'Include absent names, so a created or deleted name changes the text.
        NameSignature = _
            "LifecycleSchema=" & StoredRefersTo(host, "LifecycleSchema") & _
            "; LifecycleStatus=" & StoredRefersTo(host, "LifecycleStatus") & _
            "; LifecycleStarts=" & StoredRefersTo(host, "LifecycleStarts")

End Function


Private Function HarnessDifference() _
    As String
'
'==============================================================================
'                              HarnessDifference
'------------------------------------------------------------------------------
' PURPOSE
'   Describe what differs from the state found at suite start.
'
' RETURNS
'   An empty string when the open-workbook count, the active workbook and the
'   lifecycle are as found; otherwise one clause per difference.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Accumulate in a local, not in the function name.
    Dim difference   As String    'Clauses found so far

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Name each difference so a failed cleanup says what leaked.
        If Application.Workbooks.Count <> mSuiteBookCount Then
            difference = difference & _
                "workbooks open=" & CStr(Application.Workbooks.Count) & _
                ", expected=" & CStr(mSuiteBookCount) & "; "
        End If
        If Not (Application.ActiveWorkbook Is mSuiteActiveBook) Then
            difference = difference & "active workbook changed; "
        End If
        If LifecycleFacade.LifecycleIsRunning() Then
            difference = difference & "lifecycle session still running; "
        End If
        HarnessDifference = difference

End Function


Private Sub ReleaseCaseState()
'
'==============================================================================
'                               ReleaseCaseState
'------------------------------------------------------------------------------
' PURPOSE
'   Give back everything a case may hold: the seam, the session, the hosts
'   and activation.
'
' SIDE EFFECTS
'   Disarm the seam, stop a running session, close every tracked host
'   without saving and activate the workbook found at suite start.
'
' ERROR POLICY
'   Each step contains its own errors and reports them as failures, so this
'   is safe inside a case handler and during cleanup.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' RELEASE
'------------------------------------------------------------------------------
    'Stop before closing, so the session never outlives its host.
        LifecycleCore.ClearFault
        If Not StopLeakedSession() Then
            RecordFailure "harness.release", "a running lifecycle session could not be stopped"
        End If
        CloseHosts
        If Not RestoreActiveBook() Then
            RecordFailure "harness.release", "the workbook active at suite start could not be activated again"
        End If

End Sub


Private Function StopLeakedSession() _
    As Boolean
'
'==============================================================================
'                              StopLeakedSession
'------------------------------------------------------------------------------
' PURPOSE
'   Stop a session that a case left running.
'
' RETURNS
'   True when no session is running afterwards.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' STOP
'------------------------------------------------------------------------------
    'Use the component's own stop so its stored status is written.
        On Error GoTo StopFailed
        If LifecycleFacade.LifecycleIsRunning() Then LifecycleFacade.LifecycleStop
        StopLeakedSession = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE STOP ERROR
'------------------------------------------------------------------------------
StopFailed:
    'A failed stop still releases the session; report the failure.
        StopLeakedSession = Not LifecycleFacade.LifecycleIsRunning()

End Function


Private Sub CloseHosts()
'
'==============================================================================
'                                  CloseHosts
'------------------------------------------------------------------------------
' PURPOSE
'   Close every tracked host without saving.
'
' ERROR POLICY
'   A host that cannot be closed is reported and untracked, so one failure
'   cannot stop the others from closing.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Take each host off the list before closing it.
    Dim host   As Workbook    'Host being closed

'------------------------------------------------------------------------------
' CLOSE
'------------------------------------------------------------------------------
    'Nothing is tracked before the first run.
        If mHosts Is Nothing Then Exit Sub
        Do While mHosts.Count > 0
            Set host = mHosts(1)
            mHosts.Remove 1
            If Not CloseHost(host) Then
                RecordFailure "harness.release", "a temporary host workbook could not be closed"
            End If
        Loop

End Sub


Private Function CloseHost( _
    ByVal host As Workbook) _
    As Boolean
'
'==============================================================================
'                                  CloseHost
'------------------------------------------------------------------------------
' PURPOSE
'   Close one temporary host without saving.
'
' RETURNS
'   True when the workbook closed.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CLOSE
'------------------------------------------------------------------------------
    'SaveChanges:=False closes without a prompt, so DisplayAlerts stays as found.
        On Error GoTo CloseFailed
        host.Close SaveChanges:=False
        CloseHost = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE CLOSE ERROR
'------------------------------------------------------------------------------
CloseFailed:
    'Report the failure to the caller.
        CloseHost = False

End Function


Private Function RestoreActiveBook() _
    As Boolean
'
'==============================================================================
'                              RestoreActiveBook
'------------------------------------------------------------------------------
' PURPOSE
'   Activate the workbook that was active at suite start, if it changed.
'
' RETURNS
'   True when that workbook is active afterwards, or none was recorded.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' ACTIVATE
'------------------------------------------------------------------------------
    'Only the harness activates anything; the component never does.
        On Error GoTo ActivateFailed
        If Not mSuiteActiveBook Is Nothing Then
            If Not (Application.ActiveWorkbook Is mSuiteActiveBook) Then mSuiteActiveBook.Activate
        End If
        RestoreActiveBook = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ACTIVATION ERROR
'------------------------------------------------------------------------------
ActivateFailed:
    'Report the failure to the caller.
        RestoreActiveBook = False

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
'   2026-11-03
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
'   2026-11-03
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
'   Release the case state and the run flag, then verify that the observed
'   Excel state is unchanged.
'
' INPUTS
'   stateSnapshotValid indicates whether all four initial properties were
'   read. The initial values belong to the current run only.
'
' RETURNS
'   Boolean success and a ByRef cleanupDetail for the final report.
'
' ERROR POLICY
'   Return False for an unavailable snapshot, changed state, a leaked host
'   or session, or a cleanup error, and report the reason.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Track the combined comparison without taking ownership of host state.
    Dim difference            As String     'HarnessDifference after release
    Dim excelStateUnchanged   As Boolean    'All four observed properties match the snapshot

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Release whatever a failed case left, then the re-entry flag, even when
    'no complete host snapshot is available to verify the rest.
        On Error GoTo CleanupFailed

        ReleaseCaseState
        mRunActive = False
        If Not stateSnapshotValid Or Not mSuiteCaptured Then
            cleanupDetail = "run flag cleared; Excel state snapshot unavailable"
            CleanupRun = False
            Exit Function
        End If

    'The suite never changes calculation, alerts, events or screen updating;
    'workbooks and activation were given back by ReleaseCaseState, so any
    'difference is a real cleanup failure.
        excelStateUnchanged = _
            (Application.Calculation = initialCalculation) And _
            (Application.DisplayAlerts = initialDisplayAlerts) And _
            (Application.EnableEvents = initialEnableEvents) And _
            (Application.ScreenUpdating = initialScreenUpdating)
        difference = HarnessDifference()

    'Return the combined result and retain a reason for either outcome.
        CleanupRun = excelStateUnchanged And (Len(difference) = 0)
        If CleanupRun Then
            cleanupDetail = _
                "run flag cleared; verified Excel state unchanged " & _
                "(calculation/display alerts/events/screen updating), " & _
                "temporary workbooks closed, activation restored, no session running"
        ElseIf Not excelStateUnchanged Then
            cleanupDetail = "run flag cleared; Excel application state changed during test run"
        Else
            cleanupDetail = "run flag cleared; " & difference
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
        mRunActive = False
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
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT ENVIRONMENT
'------------------------------------------------------------------------------
    'Write the preamble before the host fields so retained output can be
    'recognized and attributed to the environment that produced it.
        Debug.Print "LIFECYCLE TESTS"
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
'   2026-11-03
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
    'infer Office bitness from the operating-system description. Office for
    'Mac 2016 and later is 64-bit VBA7 although Win64 is False there.
#If Win64 Or (Mac And VBA7) Then
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
'   2026-11-03
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
'   Clear counters, failure text, active-run state, completion and the
'   suite-start values. This helper never changes Excel state.
'
' UPDATED
'   2026-11-03
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
        mSuiteCaptured = False
        mSuiteBookCount = 0
        Set mSuiteActiveBook = Nothing

End Sub
