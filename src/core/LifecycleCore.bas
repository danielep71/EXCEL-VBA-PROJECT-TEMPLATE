Attribute VB_Name = "LifecycleCore"
'==============================================================================
' MODULE: LifecycleCore
'------------------------------------------------------------------------------
' PURPOSE
'   Own the application lifecycle behind LifecycleFacade: discover the host
'   workbook's stored state without changing it, migrate that state to the
'   current schema, run one application session, stop it, and recover or
'   reset after an interrupted session.
'
' PUBLIC SURFACE
'   None outside this VBA project. The Public procedures and constants below
'   serve LifecycleFacade and focused in-project tests; Option Private Module
'   keeps them off the supported external surface.
'
' DEPENDENCIES
'   VBA runtime and the Excel object model only.
'
' STATE OWNERSHIP
'   Owns one in-memory session: the lifecycle state, the host workbook
'   reference, the start step in progress, the owned-name snapshot taken at
'   start and the fault-injection seam. In the host workbook it owns
'   three hidden workbook-level names: LifecycleSchema, LifecycleStatus and
'   LifecycleStarts. Every other name, sheet, cell, window and Application
'   property is caller-owned and is never read or changed.
'
' ERROR POLICY
'   Discovery and argument checks raise before anything changes. A failure
'   after ownership is taken restores the three names as found and raises
'   ERR_START_FAILED naming the failed step and saying when a name could not
'   be restored. Each failure has its own core-owned number.
'
' WORKSHEET SAFETY
'   Never selects or activates a workbook, worksheet or range and never reads
'   ActiveWorkbook, ActiveSheet, Selection or Application.Caller. The only
'   cell writes go to the target range a caller passes to WriteStatus.
'
' TEST SEAM
'   SetFault makes a named start step fail, or re-enter StartApplication,
'   so tests can prove rollback and re-entry refusal; ProbeError reports the
'   re-entrant call's error number. ClearFault removes the seam.
'
' COMPATIBILITY
'   Excel VBA on Windows (VBA6/VBA7, 32- and 64-bit) and macOS, for workbooks
'   and add-ins alike: hidden names need no window, sheet or activation.
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
    'Own the lifecycle error codes; each facade constant aliases one of them.
        Public Const ERR_RUNNING             As Long = vbObjectError + 2300    'Session owned elsewhere, or reset while running
        Public Const ERR_NOT_RUNNING         As Long = vbObjectError + 2301    'The operation needs a running session
        Public Const ERR_BUSY                As Long = vbObjectError + 2302    'Re-entrant call while starting or stopping
        Public Const ERR_INVALID_ARGUMENT    As Long = vbObjectError + 2303    'Rejected host or target argument
        Public Const ERR_STATE_INVALID       As Long = vbObjectError + 2304    'Stored state cannot be read
        Public Const ERR_STATE_NEWER         As Long = vbObjectError + 2305    'Stored schema is newer than supported
        Public Const ERR_RECOVERY_REQUIRED   As Long = vbObjectError + 2306    'Previous session did not stop
        Public Const ERR_START_FAILED        As Long = vbObjectError + 2307    'A start step failed; rollback attempted
        Public Const ERR_STOP_FAILED         As Long = vbObjectError + 2308    'Stopped status could not be recorded
        Private Const ERR_INJECTED_FAULT     As Long = vbObjectError + 2399    'Test seam failure inside a start step

    'Name the stored state and the schema this version reads and writes.
        Public Const SCHEMA_CURRENT          As Long = 2                       'Schema written by this version
        Public Const NAME_SCHEMA             As String = "LifecycleSchema"     'Stored schema version
        Public Const NAME_STATUS             As String = "LifecycleStatus"     'Stored session status
        Public Const NAME_STARTS             As String = "LifecycleStarts"     'Completed starts, from schema 2
        Public Const STATUS_STOPPED          As String = "stopped"             'No session, or a clean stop
        Public Const STATUS_RUNNING          As String = "running"             'A session is running or did not stop
        Private Const OWNED_NAME_COUNT       As Long = 3                       'Names owned in the host workbook

    'Number the in-memory lifecycle states.
        Private Const STATE_STOPPED          As Long = 0                       'No session
        Private Const STATE_STARTING         As Long = 1                       'Start is changing the host
        Private Const STATE_RUNNING          As Long = 2                       'Session running
        Private Const STATE_STOPPING         As Long = 3                       'Stop is changing the host

'------------------------------------------------------------------------------
' MODULE TYPES
'------------------------------------------------------------------------------
    'Hold the stored state found by discovery, before any change.
        Private Type StoredState
            Schema   As Long      'Stored schema; 0 when nothing is stored
            Status   As String    'Stored status; empty for schema 0
            Starts   As Long      'Stored start count; 0 before schema 2
        End Type

    'Hold one owned name exactly as found, for rollback.
        Private Type NameSnapshot
            Exists     As Boolean    'True when the name existed
            RefersTo   As String     'Original definition
            Visible    As Boolean    'Original visibility
        End Type

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Describe the one session. Stop releases the host, the next start
    'overwrites the snapshot, and only ClearFault disarms the seam.
        Private mFaultReenter   As Boolean                               'Seam: re-enter instead of failing
        Private mFaultStep      As String                                'Seam: step that fails; empty for none
        Private mHost           As Workbook                              'Host of the running session
        Private mProbeError     As Long                                  'Seam: error number of the re-entrant call
        Private mSnapshot(1 To OWNED_NAME_COUNT)   As NameSnapshot       'Owned names as found at start
        Private mState          As Long                                  'One of the STATE_ constants
        Private mStep           As String                                'Start step in progress


'
'------------------------------------------------------------------------------
'
'                              SESSION LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Function StartApplication( _
    ByVal host As Workbook) _
    As Boolean
'
'==============================================================================
'                               StartApplication
'------------------------------------------------------------------------------
' PURPOSE
'   Discover the host's stored state, migrate it to the current schema and
'   start the session.
'
' INPUTS
'   host: the workbook or add-in that stores the application state.
'
' RETURNS
'   True when a session started; False when this host's session was already
'   running, in which case nothing changes.
'
' SIDE EFFECTS
'   Run each migration from the stored schema to SCHEMA_CURRENT, set the
'   stored status to running and add one to the stored start count.
'
' ERROR POLICY
'   Raise ERR_BUSY, ERR_INVALID_ARGUMENT, ERR_RUNNING, ERR_STATE_INVALID,
'   ERR_STATE_NEWER or ERR_RECOVERY_REQUIRED before changing anything. A
'   later failure restores the owned names and raises ERR_START_FAILED; its
'   message ends "Rollback was incomplete." when a name was not restored.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the discovered state and the failure record separate.
    Dim failedStep         As String         'Step that raised, for the diagnostic
    Dim rolledBack         As Boolean        'True when every owned name was restored
    Dim savedDescription   As String         'Original step failure
    Dim stored             As StoredState    'State found by discovery
    Dim version            As Long           'Schema produced by the next migration

'------------------------------------------------------------------------------
' CHECK CALL
'------------------------------------------------------------------------------
    'Refuse re-entry and bad arguments first; a repeated start of the same
    'host is a no-op so it is safe from more than one event handler.
        If mState = STATE_STARTING Or mState = STATE_STOPPING Then
            RaiseCore ERR_BUSY, "StartApplication", "The application is starting or stopping."
        End If
        If host Is Nothing Then
            RaiseCore ERR_INVALID_ARGUMENT, "StartApplication", "host must be a workbook."
        End If
        If mState = STATE_RUNNING Then
            If host Is mHost Then
                StartApplication = False
                Exit Function
            End If
            RaiseCore ERR_RUNNING, "StartApplication", "Another workbook owns the running application."
        End If

'------------------------------------------------------------------------------
' DISCOVER
'------------------------------------------------------------------------------
    'Read and validate without changing anything, so a refusal here needs no
    'cleanup: ownership has not been taken.
        stored = ReadState(host, False)
        If stored.Status = STATUS_RUNNING Then
            RaiseCore ERR_RECOVERY_REQUIRED, "StartApplication", "The previous session did not stop; run LifecycleRecover first."
        End If

'------------------------------------------------------------------------------
' TAKE OWNERSHIP
'------------------------------------------------------------------------------
    'Snapshot the owned names before the first change; from here on a failure
    'must give them back.
        TakeSnapshot host
        Set mHost = host
        mState = STATE_STARTING
        On Error GoTo StepFailed

'------------------------------------------------------------------------------
' MIGRATE AND START
'------------------------------------------------------------------------------
    'Each migration writes its schema last, so a stored schema always
    'describes complete content. The current schema runs no migration.
        For version = stored.Schema + 1 To SCHEMA_CURRENT
            mStep = "migrate-" & CStr(version)
            CheckFault mStep
            ApplyMigration host, version
        Next version
        If stored.Schema < 2 Then stored.Starts = 0

        mStep = "status"
        CheckFault mStep
        WriteText host, NAME_STATUS, STATUS_RUNNING

        mStep = "count"
        CheckFault mStep
        WriteNumber host, NAME_STARTS, stored.Starts + 1

        mStep = vbNullString
        mState = STATE_RUNNING
        StartApplication = True
        Exit Function

'------------------------------------------------------------------------------
' ROLL BACK
'------------------------------------------------------------------------------
StepFailed:
    'Capture the failure, give every owned name back and release ownership
    'before raising, so a failed start leaves no session and no partial state.
        failedStep = mStep
        savedDescription = Err.Description
        rolledBack = RestoreSnapshot(host)
        Set mHost = Nothing
        mStep = vbNullString
        mState = STATE_STOPPED
        If rolledBack Then
            RaiseCore _
                ERR_START_FAILED, _
                "StartApplication", _
                "Start failed at step " & failedStep & ": " & savedDescription
        Else
            RaiseCore _
                ERR_START_FAILED, _
                "StartApplication", _
                "Start failed at step " & failedStep & ": " & savedDescription & " Rollback was incomplete."
        End If

End Function


Public Function StopApplication() _
    As Boolean
'
'==============================================================================
'                               StopApplication
'------------------------------------------------------------------------------
' PURPOSE
'   Stop the running session and record a clean stop in the host.
'
' RETURNS
'   True when a session stopped; False when none was running, so the call
'   is safe in every exit path.
'
' ERROR POLICY
'   Raise ERR_BUSY on re-entry. If the stopped status cannot be written the
'   session is still released, the stored status stays running, and
'   ERR_STOP_FAILED is raised; the next start then requires recovery.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the original diagnostic after ownership is released.
    Dim savedDescription   As String    'Original write failure

'------------------------------------------------------------------------------
' CHECK CALL
'------------------------------------------------------------------------------
    'Without a session there is nothing to stop.
        If mState = STATE_STARTING Or mState = STATE_STOPPING Then
            RaiseCore ERR_BUSY, "StopApplication", "The application is starting or stopping."
        End If
        If mState <> STATE_RUNNING Then
            StopApplication = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' STOP
'------------------------------------------------------------------------------
    'Record the clean stop, then release the in-memory session.
        mState = STATE_STOPPING
        On Error GoTo StopFailed
        WriteText mHost, NAME_STATUS, STATUS_STOPPED
        Set mHost = Nothing
        mState = STATE_STOPPED
        StopApplication = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE STOP ERROR
'------------------------------------------------------------------------------
StopFailed:
    'Never keep a session the caller believes stopped: release it, and leave
    'the stored running status as the signal that recovery is needed.
        savedDescription = Err.Description
        Set mHost = Nothing
        mState = STATE_STOPPED
        RaiseCore ERR_STOP_FAILED, "StopApplication", "Stop could not record the stopped status: " & savedDescription

End Function


Public Function RecoverApplication( _
    ByVal host As Workbook) _
    As Boolean
'
'==============================================================================
'                              RecoverApplication
'------------------------------------------------------------------------------
' PURPOSE
'   Bring the host back to a clean stopped state after an interruption.
'
' INPUTS
'   host: the workbook or add-in whose state is recovered.
'
' RETURNS
'   True when something was recovered: this host's live session stopped, or
'   a running status left by a session that never stopped was cleared.
'   False when the host was already cleanly stopped.
'
' ERROR POLICY
'   Raise ERR_BUSY, ERR_INVALID_ARGUMENT, ERR_RUNNING (another workbook owns
'   the session), ERR_STATE_INVALID or ERR_STATE_NEWER without changing
'   anything; such state needs ResetApplication or a newer version. A live
'   session for this host is stopped by StopApplication, so ERR_STOP_FAILED
'   can follow. If a stale running status cannot be overwritten, Excel's
'   error propagates and the stored status stays running.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold the state found in the host.
    Dim stored   As StoredState    'State found by discovery

'------------------------------------------------------------------------------
' CHECK CALL
'------------------------------------------------------------------------------
    'A live session for this host is recovered by stopping it normally.
        If mState = STATE_STARTING Or mState = STATE_STOPPING Then
            RaiseCore ERR_BUSY, "RecoverApplication", "The application is starting or stopping."
        End If
        If host Is Nothing Then
            RaiseCore ERR_INVALID_ARGUMENT, "RecoverApplication", "host must be a workbook."
        End If
        If mState = STATE_RUNNING Then
            If host Is mHost Then
                RecoverApplication = StopApplication()
                Exit Function
            End If
            RaiseCore ERR_RUNNING, "RecoverApplication", "Another workbook owns the running application."
        End If

'------------------------------------------------------------------------------
' RECOVER
'------------------------------------------------------------------------------
    'A stored running status without a session means the last session ended
    'by a crash or a VBA reset; recording the stop is the whole repair.
        stored = ReadState(host, False)
        If stored.Status = STATUS_RUNNING Then
            WriteText host, NAME_STATUS, STATUS_STOPPED
            RecoverApplication = True
        Else
            RecoverApplication = False
        End If

End Function


Public Function ResetApplication( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                               ResetApplication
'------------------------------------------------------------------------------
' PURPOSE
'   Remove every owned name from the host, returning it to a fresh install.
'
' INPUTS
'   host: the workbook or add-in whose stored state is removed.
'
' RETURNS
'   The number of owned names removed, from 0 to 3.
'
' ERROR POLICY
'   Raise ERR_BUSY, ERR_INVALID_ARGUMENT, or ERR_RUNNING when this host's
'   session is running. Invalid or newer state is removed too: reset is the
'   way out of state that recovery refuses.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Count the names actually removed.
    Dim index     As Long    'Owned-name position, 1 to OWNED_NAME_COUNT
    Dim removed   As Long    'Names deleted so far

'------------------------------------------------------------------------------
' CHECK CALL
'------------------------------------------------------------------------------
    'Never pull state out from under a running session.
        If mState = STATE_STARTING Or mState = STATE_STOPPING Then
            RaiseCore ERR_BUSY, "ResetApplication", "The application is starting or stopping."
        End If
        If host Is Nothing Then
            RaiseCore ERR_INVALID_ARGUMENT, "ResetApplication", "host must be a workbook."
        End If
        If mState = STATE_RUNNING Then
            If host Is mHost Then
                RaiseCore ERR_RUNNING, "ResetApplication", "Stop the application before resetting its state."
            End If
        End If

'------------------------------------------------------------------------------
' REMOVE OWNED NAMES
'------------------------------------------------------------------------------
    'Delete only the three owned names; any other name is caller-owned.
        For index = 1 To OWNED_NAME_COUNT
            If DeleteOwnedName(host, OwnedName(index)) Then removed = removed + 1
        Next index
        ResetApplication = removed

End Function


'
'------------------------------------------------------------------------------
'
'                             OPERATION AND QUERIES
'
'------------------------------------------------------------------------------
'

Public Function WriteStatus( _
    ByVal target As Range) _
    As Long
'
'==============================================================================
'                                 WriteStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Write the running session's status block to an explicit target range.
'
' INPUTS
'   target: the top-left cell of a four-row, two-column block, on any sheet
'           of any open workbook.
'
' RETURNS
'   The number of rows written (4): Schema, Status, Starts and Host.
'
' SIDE EFFECTS
'   Replace the values of the eight cells; formats and every other cell are
'   unchanged. Nothing is selected or activated.
'
' ERROR POLICY
'   Raise ERR_NOT_RUNNING without a session and ERR_INVALID_ARGUMENT for a
'   missing target, before writing anything.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Build the block in memory and write it in one assignment.
    Dim block(1 To 4, 1 To 2)   As Variant        'Status labels and values
    Dim stored                  As StoredState    'Current stored state of the host

'------------------------------------------------------------------------------
' CHECK CALL
'------------------------------------------------------------------------------
    'Report only a running session, and only to a range the caller named.
        If mState <> STATE_RUNNING Then
            RaiseCore ERR_NOT_RUNNING, "WriteStatus", "The application is not running."
        End If
        If target Is Nothing Then
            RaiseCore ERR_INVALID_ARGUMENT, "WriteStatus", "target must be a cell range."
        End If

'------------------------------------------------------------------------------
' WRITE BLOCK
'------------------------------------------------------------------------------
    'Address the block from the target itself, never from the active sheet.
        stored = ReadState(mHost, False)
        block(1, 1) = "Schema"
        block(1, 2) = stored.Schema
        block(2, 1) = "Status"
        block(2, 2) = stored.Status
        block(3, 1) = "Starts"
        block(3, 2) = stored.Starts
        block(4, 1) = "Host"
        block(4, 2) = mHost.Name
        target.Cells(1, 1).Resize(4, 2).Value = block
        WriteStatus = 4

End Function


Public Function IsRunning() _
    As Boolean
'
'==============================================================================
'                                  IsRunning
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether a session is running.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Starting and stopping are transitions, not a running session.
        IsRunning = (mState = STATE_RUNNING)

End Function


Public Function CurrentStatus() _
    As String
'
'==============================================================================
'                                CurrentStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Name the in-memory lifecycle state.
'
' RETURNS
'   "stopped", "starting", "running" or "stopping".
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Use the same words as the stored status where they overlap.
        Select Case mState
            Case STATE_STARTING
                CurrentStatus = "starting"
            Case STATE_RUNNING
                CurrentStatus = STATUS_RUNNING
            Case STATE_STOPPING
                CurrentStatus = "stopping"
            Case Else
                CurrentStatus = STATUS_STOPPED
        End Select

End Function


Public Function StoredSchema( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                                 StoredSchema
'------------------------------------------------------------------------------
' PURPOSE
'   Read the host's stored schema without taking ownership.
'
' RETURNS
'   0 for a host with no stored state; otherwise the stored schema, which
'   may be newer than SCHEMA_CURRENT.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT for a missing host and ERR_STATE_INVALID for
'   unreadable state.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold the state found in the host.
    Dim stored   As StoredState    'State found by discovery

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'A newer schema is reported, not refused: only starting it is unsafe.
        If host Is Nothing Then
            RaiseCore ERR_INVALID_ARGUMENT, "StoredSchema", "host must be a workbook."
        End If
        stored = ReadState(host, True)
        StoredSchema = stored.Schema

End Function


Public Function StoredStarts( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                                 StoredStarts
'------------------------------------------------------------------------------
' PURPOSE
'   Read the host's stored count of completed starts.
'
' RETURNS
'   The stored count; 0 before schema 2 or for a host with no stored state.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT, ERR_STATE_INVALID or ERR_STATE_NEWER.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold the state found in the host.
    Dim stored   As StoredState    'State found by discovery

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Only a schema this version understands has a defined start count.
        If host Is Nothing Then
            RaiseCore ERR_INVALID_ARGUMENT, "StoredStarts", "host must be a workbook."
        End If
        stored = ReadState(host, False)
        StoredStarts = stored.Starts

End Function


'
'------------------------------------------------------------------------------
'
'                                  TEST SEAM
'
'------------------------------------------------------------------------------
'

Public Sub SetFault( _
    ByVal stepName As String, _
    Optional ByVal reenter As Boolean = False)
'
'==============================================================================
'                                   SetFault
'------------------------------------------------------------------------------
' PURPOSE
'   Make one start step fail, or re-enter StartApplication, on the next start.
'
' INPUTS
'   stepName: "migrate-1", "migrate-2", "status" or "count".
'   reenter:  False to raise an injected fault; True to call
'             StartApplication from inside the step and record its error.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' ARM
'------------------------------------------------------------------------------
    'Only tests call this; ClearFault disarms it.
        mFaultStep = stepName
        mFaultReenter = reenter
        mProbeError = 0

End Sub


Public Sub ClearFault()
'
'==============================================================================
'                                  ClearFault
'------------------------------------------------------------------------------
' PURPOSE
'   Disarm the fault-injection seam.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DISARM
'------------------------------------------------------------------------------
    'Keep ProbeError readable until the next SetFault.
        mFaultStep = vbNullString
        mFaultReenter = False

End Sub


Public Function ProbeError() _
    As Long
'
'==============================================================================
'                                  ProbeError
'------------------------------------------------------------------------------
' PURPOSE
'   Report the error number the re-entrant probe received; 0 for none.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Set by CheckFault in re-entry mode.
        ProbeError = mProbeError

End Function


'
'------------------------------------------------------------------------------
'
'                               PRIVATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub ApplyMigration( _
    ByVal host As Workbook, _
    ByVal version As Long)
'
'==============================================================================
'                                ApplyMigration
'------------------------------------------------------------------------------
' PURPOSE
'   Migrate the stored state to one schema version.
'
' INPUTS
'   host:    the workbook being started.
'   version: the schema this step produces, 1 to SCHEMA_CURRENT.
'
' ERROR POLICY
'   Write errors propagate to StartApplication, which rolls back.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' MIGRATE
'------------------------------------------------------------------------------
    'Schema 1 records a status; schema 2 adds the start count. Add a Case
    'and raise SCHEMA_CURRENT for each future schema.
        Select Case version
            Case 1
                WriteText host, NAME_STATUS, STATUS_STOPPED
                WriteNumber host, NAME_SCHEMA, 1
            Case 2
                WriteNumber host, NAME_STARTS, 0
                WriteNumber host, NAME_SCHEMA, 2
        End Select

End Sub


Private Function ReadState( _
    ByVal host As Workbook, _
    ByVal allowNewer As Boolean) _
    As StoredState
'
'==============================================================================
'                                  ReadState
'------------------------------------------------------------------------------
' PURPOSE
'   Read and validate the host's stored state without changing it.
'
' INPUTS
'   host:       the workbook whose owned names are read.
'   allowNewer: True to return a newer schema instead of refusing it.
'
' RETURNS
'   Schema 0 with an empty status when nothing is stored.
'
' ERROR POLICY
'   Raise ERR_STATE_INVALID naming the first unreadable name, and
'   ERR_STATE_NEWER for a newer schema unless allowNewer.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Read each owned name once.
    Dim found    As Boolean        'True when the name exists
    Dim result   As StoredState    'Validated state
    Dim text     As String         'Stored value without its formula wrapper

'------------------------------------------------------------------------------
' READ SCHEMA
'------------------------------------------------------------------------------
    'No schema name means a fresh install; the migrations will write all three.
        text = ReadOwnedValue(host, NAME_SCHEMA, found)
        If Not found Then
            ReadState = result
            Exit Function
        End If
        If Not IsCount(text) Then
            RaiseCore ERR_STATE_INVALID, "ReadState", "Stored lifecycle state is invalid: " & NAME_SCHEMA & "."
        End If
        result.Schema = CLng(text)
        If result.Schema < 1 Then
            RaiseCore ERR_STATE_INVALID, "ReadState", "Stored lifecycle state is invalid: " & NAME_SCHEMA & "."
        End If
        If result.Schema > SCHEMA_CURRENT Then
            If allowNewer Then
                ReadState = result
                Exit Function
            End If
            RaiseCore _
                ERR_STATE_NEWER, _
                "ReadState", _
                "Stored schema " & CStr(result.Schema) & " is newer than the supported schema " & CStr(SCHEMA_CURRENT) & "."
        End If

'------------------------------------------------------------------------------
' READ STATUS AND COUNT
'------------------------------------------------------------------------------
    'Every schema records a status; schema 2 adds the start count.
        result.Status = ReadOwnedValue(host, NAME_STATUS, found)
        If Not found Or (result.Status <> STATUS_STOPPED And result.Status <> STATUS_RUNNING) Then
            RaiseCore ERR_STATE_INVALID, "ReadState", "Stored lifecycle state is invalid: " & NAME_STATUS & "."
        End If
        If result.Schema >= 2 Then
            text = ReadOwnedValue(host, NAME_STARTS, found)
            If Not found Or Not IsCount(text) Then
                RaiseCore ERR_STATE_INVALID, "ReadState", "Stored lifecycle state is invalid: " & NAME_STARTS & "."
            End If
            result.Starts = CLng(text)
        End If
        ReadState = result

End Function


Private Function ReadOwnedValue( _
    ByVal host As Workbook, _
    ByVal nameText As String, _
    ByRef found As Boolean) _
    As String
'
'==============================================================================
'                                ReadOwnedValue
'------------------------------------------------------------------------------
' PURPOSE
'   Read one owned name's constant value.
'
' RETURNS
'   The value without the leading "=" and, for text, without its quotes;
'   found reports whether the name exists.
'
' ERROR POLICY
'   A missing name is not an error: the local handler reports it as not
'   found and returns an empty string.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Unwrap the stored formula text.
    Dim text   As String    'RefersTo, then the unwrapped value

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Names(...) raises for a missing name; that is the not-found answer.
        On Error GoTo Missing
        text = host.Names(nameText).RefersTo
        On Error GoTo 0
        found = True
        If Left$(text, 1) = "=" Then text = Mid$(text, 2)
        If Len(text) >= 2 And Left$(text, 1) = """" And Right$(text, 1) = """" Then
            text = Replace(Mid$(text, 2, Len(text) - 2), """""", """")
        End If
        ReadOwnedValue = text
        Exit Function

'------------------------------------------------------------------------------
' HANDLE MISSING NAME
'------------------------------------------------------------------------------
Missing:
    'Leave the handler with a clean not-found result.
        found = False
        ReadOwnedValue = vbNullString

End Function


Private Sub WriteText( _
    ByVal host As Workbook, _
    ByVal nameText As String, _
    ByVal value As String)
'
'==============================================================================
'                                  WriteText
'------------------------------------------------------------------------------
' PURPOSE
'   Store a text constant in a hidden owned name, creating or replacing it.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Quote the text and double embedded quotes, as Excel formula text needs.
        host.Names.Add Name:=nameText, RefersTo:="=""" & Replace(value, """", """""") & """", Visible:=False

End Sub


Private Sub WriteNumber( _
    ByVal host As Workbook, _
    ByVal nameText As String, _
    ByVal value As Long)
'
'==============================================================================
'                                 WriteNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Store a whole-number constant in a hidden owned name.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'CStr of a Long has no separators, so the text is locale-independent.
        host.Names.Add Name:=nameText, RefersTo:="=" & CStr(value), Visible:=False

End Sub


Private Function DeleteOwnedName( _
    ByVal host As Workbook, _
    ByVal nameText As String) _
    As Boolean
'
'==============================================================================
'                               DeleteOwnedName
'------------------------------------------------------------------------------
' PURPOSE
'   Delete one owned name if it exists.
'
' RETURNS
'   True when a name was deleted; False when it did not exist.
'
' ERROR POLICY
'   Only the lookup is guarded; a failure to delete an existing name
'   propagates.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Hold the name once found.
    Dim owned   As Name    'Owned name to delete

'------------------------------------------------------------------------------
' FIND AND DELETE
'------------------------------------------------------------------------------
    'Look the name up under a local handler, then delete it outside it.
        On Error GoTo Missing
        Set owned = host.Names(nameText)
        On Error GoTo 0
        owned.Delete
        DeleteOwnedName = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE MISSING NAME
'------------------------------------------------------------------------------
Missing:
    'Nothing to delete.
        DeleteOwnedName = False

End Function


Private Sub TakeSnapshot( _
    ByVal host As Workbook)
'
'==============================================================================
'                                 TakeSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Record each owned name exactly as found, for rollback.
'
' ERROR POLICY
'   Called before ownership is taken; a read failure propagates and nothing
'   has changed yet.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Visit the owned names in a fixed order.
    Dim index   As Long    'Owned-name position, 1 to OWNED_NAME_COUNT

'------------------------------------------------------------------------------
' SNAPSHOT
'------------------------------------------------------------------------------
    'Keep the original definition and visibility, not the parsed value.
        For index = 1 To OWNED_NAME_COUNT
            mSnapshot(index) = SnapshotName(host, OwnedName(index))
        Next index

End Sub


Private Function SnapshotName( _
    ByVal host As Workbook, _
    ByVal nameText As String) _
    As NameSnapshot
'
'==============================================================================
'                                 SnapshotName
'------------------------------------------------------------------------------
' PURPOSE
'   Capture one owned name's existence, definition and visibility.
'
' ERROR POLICY
'   A missing name is recorded as not existing.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Build the record from the live name.
    Dim owned    As Name            'Owned name, when it exists
    Dim result   As NameSnapshot    'Captured record

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
    'Only the lookup can legitimately fail.
        On Error GoTo Missing
        Set owned = host.Names(nameText)
        On Error GoTo 0
        result.Exists = True
        result.RefersTo = owned.RefersTo
        result.Visible = owned.Visible
        SnapshotName = result
        Exit Function

'------------------------------------------------------------------------------
' HANDLE MISSING NAME
'------------------------------------------------------------------------------
Missing:
    'Record absence; rollback deletes a name that did not exist.
        result.Exists = False
        SnapshotName = result

End Function


Private Function RestoreSnapshot( _
    ByVal host As Workbook) _
    As Boolean
'
'==============================================================================
'                               RestoreSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Give every owned name back exactly as TakeSnapshot found it.
'
' RETURNS
'   True when all three names were restored.
'
' ERROR POLICY
'   Best effort per name: one failure does not stop the others.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Combine the per-name results.
    Dim allRestored   As Boolean    'False once any name could not be restored
    Dim index         As Long       'Owned-name position, 1 to OWNED_NAME_COUNT

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore in reverse order of the start steps.
        allRestored = True
        For index = OWNED_NAME_COUNT To 1 Step -1
            If Not RestoreName(host, OwnedName(index), mSnapshot(index)) Then allRestored = False
        Next index
        RestoreSnapshot = allRestored

End Function


Private Function RestoreName( _
    ByVal host As Workbook, _
    ByVal nameText As String, _
    ByRef original As NameSnapshot) _
    As Boolean
'
'==============================================================================
'                                 RestoreName
'------------------------------------------------------------------------------
' PURPOSE
'   Restore one owned name from its snapshot.
'
' RETURNS
'   True when the restoring call raised no error; the name is not read back
'   to compare it with the snapshot.
'
' ERROR POLICY
'   Contain any error and return False so the other names are still tried.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Recreate an existing name with its original definition, or delete one
    'that the failed start created.
        On Error GoTo RestoreFailed
        If original.Exists Then
            host.Names.Add Name:=nameText, RefersTo:=original.RefersTo, Visible:=original.Visible
        Else
            DeleteOwnedName host, nameText
        End If
        RestoreName = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE RESTORE ERROR
'------------------------------------------------------------------------------
RestoreFailed:
    'Report the failure; StartApplication appends it to the diagnostic.
        RestoreName = False

End Function


Private Function OwnedName( _
    ByVal index As Long) _
    As String
'
'==============================================================================
'                                  OwnedName
'------------------------------------------------------------------------------
' PURPOSE
'   Map an owned-name position to its name.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' MAP
'------------------------------------------------------------------------------
    'Fixed order: schema, status, start count.
        Select Case index
            Case 1
                OwnedName = NAME_SCHEMA
            Case 2
                OwnedName = NAME_STATUS
            Case Else
                OwnedName = NAME_STARTS
        End Select

End Function


Private Function IsCount( _
    ByVal text As String) _
    As Boolean
'
'==============================================================================
'                                   IsCount
'------------------------------------------------------------------------------
' PURPOSE
'   Accept one to nine decimal digits, so CLng cannot overflow.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Scan each character.
    Dim index   As Long    'One-based scan position

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Like "#" matches one digit and is unaffected by locale.
        If Len(text) < 1 Or Len(text) > 9 Then Exit Function
        For index = 1 To Len(text)
            If Not (Mid$(text, index, 1) Like "#") Then Exit Function
        Next index
        IsCount = True

End Function


Private Sub CheckFault( _
    ByVal stepName As String)
'
'==============================================================================
'                                  CheckFault
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the test seam to one start step.
'
' ERROR POLICY
'   Raise the injected fault when armed for this step; in re-entry mode call
'   StartApplication and record the error it raises instead.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY SEAM
'------------------------------------------------------------------------------
    'Disarmed in production: mFaultStep is empty unless a test armed it.
        If Len(mFaultStep) = 0 Or stepName <> mFaultStep Then Exit Sub
        If mFaultReenter Then
            mProbeError = ReentryError()
        Else
            Err.Raise ERR_INJECTED_FAULT, "LifecycleCore.CheckFault", "Injected fault."
        End If

End Sub


Private Function ReentryError() _
    As Long
'
'==============================================================================
'                                 ReentryError
'------------------------------------------------------------------------------
' PURPOSE
'   Call StartApplication while a start is in progress and capture its error.
'
' RETURNS
'   The error number raised, or 0 if the call wrongly succeeded.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' PROBE
'------------------------------------------------------------------------------
    'The local handler keeps the probe's error out of the outer start.
        On Error GoTo Raised
        StartApplication mHost
        ReentryError = 0
        Exit Function

'------------------------------------------------------------------------------
' CAPTURE ERROR
'------------------------------------------------------------------------------
Raised:
    'Leaving the handler clears the error.
        ReentryError = Err.Number

End Function


Private Sub RaiseCore( _
    ByVal errorNumber As Long, _
    ByVal procedureName As String, _
    ByVal description As String)
'
'==============================================================================
'                                  RaiseCore
'------------------------------------------------------------------------------
' PURPOSE
'   Raise a core-owned lifecycle error with a qualified source.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' RAISE
'------------------------------------------------------------------------------
    'The facade replaces the source with its public procedure name.
        Err.Raise errorNumber, "LifecycleCore." & procedureName, description

End Sub
