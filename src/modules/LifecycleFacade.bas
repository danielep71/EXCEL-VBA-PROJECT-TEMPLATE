Attribute VB_Name = "LifecycleFacade"
'==============================================================================
' MODULE: LifecycleFacade
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the application profile's supported lifecycle: start, stop,
'   recover and reset an application whose versioned state lives in an
'   explicit host workbook or add-in, with a stable caller-facing error
'   contract.
'
' PUBLIC SURFACE
'   LIFECYCLE_ERROR_RUNNING, LIFECYCLE_ERROR_NOT_RUNNING, LIFECYCLE_ERROR_BUSY,
'   LIFECYCLE_ERROR_INVALID_ARGUMENT, LIFECYCLE_ERROR_STATE_INVALID,
'   LIFECYCLE_ERROR_STATE_NEWER, LIFECYCLE_ERROR_RECOVERY_REQUIRED,
'   LIFECYCLE_ERROR_START_FAILED, LIFECYCLE_ERROR_STOP_FAILED,
'   LIFECYCLE_SCHEMA_VERSION
'   LifecycleStart, LifecycleStop, LifecycleRecover, LifecycleReset,
'   LifecycleWriteStatus, LifecycleIsRunning, LifecycleStatus,
'   LifecycleStoredSchema, LifecycleStoredStarts
'
' DEPENDENCIES
'   LifecycleCore only. The dependency direction is facade -> core.
'
' STATE OWNERSHIP
'   Stateless itself; LifecycleCore owns the single session and, in the host
'   workbook, the hidden names LifecycleSchema, LifecycleStatus and
'   LifecycleStarts. Every other workbook, sheet, cell, window and
'   Application property remains caller-owned.
'
' ERROR POLICY
'   Core failures are re-raised with the called facade procedure as source;
'   number, description, help file and help context are preserved. Each
'   public error constant aliases one core-owned code.
'
' WORKSHEET SAFETY
'   Never selects or activates anything and never reads ActiveWorkbook,
'   ActiveSheet, Selection or Application.Caller. The host and the status
'   target are always passed explicitly.
'
' TEST SEAM
'   LifecycleTests drives this surface on temporary workbooks and uses the
'   core's fault-injection seam to prove rollback and re-entry refusal.
'
' COMPATIBILITY
'   Excel VBA on Windows (VBA6/VBA7, 32- and 64-bit) and macOS, in a workbook
'   or an add-in.
'
' USAGE
'   In ThisWorkbook: Workbook_Open calls LifecycleStart ThisWorkbook and
'   Workbook_BeforeClose calls LifecycleStop. After a crash or VBA reset,
'   LifecycleRecover ThisWorkbook clears the unfinished session.
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

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Expose core-owned error codes and the schema without a second source of truth.
        Public Const LIFECYCLE_ERROR_RUNNING             As Long = LifecycleCore.ERR_RUNNING              'Owned elsewhere
        Public Const LIFECYCLE_ERROR_NOT_RUNNING         As Long = LifecycleCore.ERR_NOT_RUNNING          'No running session
        Public Const LIFECYCLE_ERROR_BUSY                As Long = LifecycleCore.ERR_BUSY                 'Re-entrant call
        Public Const LIFECYCLE_ERROR_INVALID_ARGUMENT    As Long = LifecycleCore.ERR_INVALID_ARGUMENT     'Rejected argument
        Public Const LIFECYCLE_ERROR_STATE_INVALID       As Long = LifecycleCore.ERR_STATE_INVALID        'Unreadable state
        Public Const LIFECYCLE_ERROR_STATE_NEWER         As Long = LifecycleCore.ERR_STATE_NEWER          'Newer schema
        Public Const LIFECYCLE_ERROR_RECOVERY_REQUIRED   As Long = LifecycleCore.ERR_RECOVERY_REQUIRED    'Unclean stop
        Public Const LIFECYCLE_ERROR_START_FAILED        As Long = LifecycleCore.ERR_START_FAILED         'Start failed; rollback attempted
        Public Const LIFECYCLE_ERROR_STOP_FAILED         As Long = LifecycleCore.ERR_STOP_FAILED          'Stop not recorded
        Public Const LIFECYCLE_SCHEMA_VERSION            As Long = LifecycleCore.SCHEMA_CURRENT           'Current schema


'
'------------------------------------------------------------------------------
'
'                             SUPPORTED PUBLIC API
'
'------------------------------------------------------------------------------
'

Public Function LifecycleStart( _
    ByVal host As Workbook) _
    As Boolean
'
'==============================================================================
'                                LifecycleStart
'------------------------------------------------------------------------------
' PURPOSE
'   Start the application for a host workbook or add-in.
'
' INPUTS
'   host: the workbook that stores the state, usually ThisWorkbook.
'
' RETURNS
'   True when a session started; False when this host's session was
'   already running, in which case nothing changes.
'
' SIDE EFFECTS
'   Migrate the stored state to LIFECYCLE_SCHEMA_VERSION, set the stored
'   status to running and add one to the stored start count. Only the
'   three hidden lifecycle names change.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_BUSY, _INVALID_ARGUMENT, _RUNNING,
'   _STATE_INVALID, _STATE_NEWER or _RECOVERY_REQUIRED before changing
'   anything. A failed step restores the names as found and raises
'   LIFECYCLE_ERROR_START_FAILED naming the step; the message says when
'   the restore was incomplete.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleStart = LifecycleCore.StartApplication(host)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleStart"

End Function


Public Function LifecycleStop() _
    As Boolean
'
'==============================================================================
'                                LifecycleStop
'------------------------------------------------------------------------------
' PURPOSE
'   Stop the running session and record a clean stop.
'
' RETURNS
'   True when a session stopped; False when none was running, so the
'   call is safe in every exit path, including Workbook_BeforeClose.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_BUSY on re-entry. If the stopped status cannot
'   be written, the session is still released and
'   LIFECYCLE_ERROR_STOP_FAILED is raised; the next start then requires
'   LifecycleRecover.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleStop = LifecycleCore.StopApplication()
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleStop"

End Function


Public Function LifecycleRecover( _
    ByVal host As Workbook) _
    As Boolean
'
'==============================================================================
'                               LifecycleRecover
'------------------------------------------------------------------------------
' PURPOSE
'   Return the host to a clean stopped state after an interruption.
'
' INPUTS
'   host: the workbook or add-in to recover.
'
' RETURNS
'   True when this host's live session was stopped, or when a running
'   status left by a crash or VBA reset was cleared; False when the host
'   was already cleanly stopped.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_BUSY, _INVALID_ARGUMENT, _RUNNING,
'   _STATE_INVALID or _STATE_NEWER without changing anything; such
'   state needs LifecycleReset or a newer version. A live session for this
'   host is stopped as LifecycleStop does, so LIFECYCLE_ERROR_STOP_FAILED
'   can follow. If a stale running status cannot be overwritten, Excel's
'   error propagates and the stored status stays running.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleRecover = LifecycleCore.RecoverApplication(host)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleRecover"

End Function


Public Function LifecycleReset( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                                LifecycleReset
'------------------------------------------------------------------------------
' PURPOSE
'   Remove the stored lifecycle state, returning the host to a fresh
'   install.
'
' INPUTS
'   host: the workbook or add-in whose state is removed.
'
' RETURNS
'   The number of hidden lifecycle names removed, from 0 to 3.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_BUSY, _INVALID_ARGUMENT, or _RUNNING while
'   this host's session runs. Invalid or newer state is removed too.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleReset = LifecycleCore.ResetApplication(host)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleReset"

End Function


Public Function LifecycleWriteStatus( _
    ByVal target As Range) _
    As Long
'
'==============================================================================
'                             LifecycleWriteStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Write the running session's status block to an explicit range.
'
' INPUTS
'   target: top-left cell of a four-row, two-column block on any sheet of
'           any open workbook.
'
' RETURNS
'   4: rows Schema, Status, Starts and Host.
'
' SIDE EFFECTS
'   Replace the eight cells' values; nothing is selected or activated.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_NOT_RUNNING or _INVALID_ARGUMENT before writing.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleWriteStatus = LifecycleCore.WriteStatus(target)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleWriteStatus"

End Function


Public Function LifecycleIsRunning() _
    As Boolean
'
'==============================================================================
'                              LifecycleIsRunning
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether a session is running.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this operation has no failure to normalize.
        LifecycleIsRunning = LifecycleCore.IsRunning()

End Function


Public Function LifecycleStatus() _
    As String
'
'==============================================================================
'                               LifecycleStatus
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
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this operation has no failure to normalize.
        LifecycleStatus = LifecycleCore.CurrentStatus()

End Function


Public Function LifecycleStoredSchema( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                            LifecycleStoredSchema
'------------------------------------------------------------------------------
' PURPOSE
'   Read a host's stored schema without starting the application.
'
' RETURNS
'   0 when nothing is stored; otherwise the stored schema, which may be
'   newer than LIFECYCLE_SCHEMA_VERSION.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_INVALID_ARGUMENT or _STATE_INVALID.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleStoredSchema = LifecycleCore.StoredSchema(host)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleStoredSchema"

End Function


Public Function LifecycleStoredStarts( _
    ByVal host As Workbook) _
    As Long
'
'==============================================================================
'                            LifecycleStoredStarts
'------------------------------------------------------------------------------
' PURPOSE
'   Read a host's stored count of completed starts.
'
' RETURNS
'   The stored count; 0 before schema 2 or when nothing is stored.
'
' ERROR POLICY
'   Raise LIFECYCLE_ERROR_INVALID_ARGUMENT, _STATE_INVALID or
'   _STATE_NEWER.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate to the core; this boundary owns only the caller-facing
    'error source.
        On Error GoTo HandleError

        LifecycleStoredStarts = LifecycleCore.StoredStarts(host)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "LifecycleFacade.LifecycleStoredStarts"

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
'   source: the public procedure name, such as "LifecycleFacade.LifecycleStart".
'
' ERROR POLICY
'   Called only from a facade error handler, so the core error is still the
'   active one. Raising here passes the error to the facade's caller with the
'   same number, description, help file and help context.
'
' UPDATED
'   2026-11-03
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
