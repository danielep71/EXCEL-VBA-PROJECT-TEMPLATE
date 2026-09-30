Attribute VB_Name = "ProgressFacade"
'==============================================================================
' MODULE: ProgressFacade
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the ui-component profile's supported progress session: a scoped
'   owner of the Excel status bar, cursor, screen updating and Esc handling
'   for long operations, with a stable caller-facing error contract.
'
' PUBLIC SURFACE
'   PROGRESS_ERROR_ACTIVE, PROGRESS_ERROR_INACTIVE, PROGRESS_ERROR_CANCELLED,
'   PROGRESS_ERROR_INVALID_ARGUMENT
'   ProgressOptions (Enum)
'   ProgressBegin, ProgressUpdate, ProgressEnd, ProgressRecover,
'   ProgressRequestCancel, ProgressIsCancelled, ProgressIsActive,
'   ProgressElapsedSeconds, ProgressText
'
' DEPENDENCIES
'   ProgressCore only. The dependency direction is facade -> core.
'
' STATE OWNERSHIP
'   Stateless itself; ProgressCore owns the single session. During a session
'   the component owns StatusBar, DisplayStatusBar, Cursor (Windows),
'   ScreenUpdating and EnableCancelKey and restores each at ProgressEnd.
'   Calculation, EnableEvents, DisplayAlerts, selection, workbooks and
'   worksheets remain caller-owned and are never changed.
'
' ERROR POLICY
'   Core failures are re-raised with the called facade procedure as source;
'   number, description, help file and help context are preserved. Each
'   public error constant aliases one core-owned code.
'
' WORKSHEET SAFETY
'   Never selects, activates, edits, calculates or saves a workbook or
'   worksheet and never reads Application.Caller or the selection.
'
' TEST SEAM
'   ProgressTests drives this surface and compares the owned Application
'   properties before, during and after each session, including a negative
'   control that proves a broken restoration is detected.
'
' COMPATIBILITY
'   Excel VBA on Windows (VBA6/VBA7, 32- and 64-bit) and macOS. The text is
'   the whole interface: no layout, so no DPI or scaling assumption; Esc is
'   the keyboard cancel; the cursor is not changed on macOS.
'
' USAGE
'   ProgressBegin "Copying", total
'   ... ProgressUpdate done ... (handle error 18 with ProgressRequestCancel)
'   ProgressEnd in every exit path; ProgressRecover after an interruption.
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
    'Expose core-owned error codes without defining a second source of truth.
        Public Const PROGRESS_ERROR_ACTIVE              As Long = ProgressCore.ERR_SESSION_ACTIVE      'Begin while active
        Public Const PROGRESS_ERROR_INACTIVE            As Long = ProgressCore.ERR_SESSION_INACTIVE    'No active session
        Public Const PROGRESS_ERROR_CANCELLED           As Long = ProgressCore.ERR_CANCELLED           'Update after cancel
        Public Const PROGRESS_ERROR_INVALID_ARGUMENT    As Long = ProgressCore.ERR_INVALID_ARGUMENT    'Rejected argument

'------------------------------------------------------------------------------
' MODULE TYPES
'------------------------------------------------------------------------------
    'Name the option flags; each value aliases the core flag it selects.
        Public Enum ProgressOptions
            ProgressDefault = 0                                             'Standard session
            ProgressKeepScreenUpdating = ProgressCore.OPTION_KEEP_SCREEN    'Leave ScreenUpdating as found
            ProgressShowElapsed = ProgressCore.OPTION_SHOW_ELAPSED          'Append elapsed seconds
        End Enum


'
'------------------------------------------------------------------------------
'
'                             SUPPORTED PUBLIC API
'
'------------------------------------------------------------------------------
'

Public Sub ProgressBegin( _
    ByVal caption As String, _
    Optional ByVal total As Long = 0, _
    Optional ByVal options As ProgressOptions = ProgressDefault)
'
'==============================================================================
'                                ProgressBegin
'------------------------------------------------------------------------------
' PURPOSE
'   Start the progress session and take the owned Excel UI state.
'
' INPUTS
'   caption: nonblank text shown before the progress figures.
'   total:   number of steps; 0 (default) shows a count without a percentage.
'   options: ProgressDefault, or ProgressKeepScreenUpdating and/or
'            ProgressShowElapsed combined with Or.
'
' SIDE EFFECTS
'   Snapshot StatusBar, DisplayStatusBar, Cursor (Windows), ScreenUpdating
'   and EnableCancelKey; show the first progress text, set the wait cursor,
'   turn screen updating off unless kept, and make Esc raise error 18 in the
'   caller's code. Every changed property is restored by ProgressEnd.
'
' ERROR POLICY
'   Raise PROGRESS_ERROR_INVALID_ARGUMENT or PROGRESS_ERROR_ACTIVE before
'   changing anything. A failure while taking state is rolled back first.
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

        ProgressCore.BeginSession caption, total, options
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "ProgressFacade.ProgressBegin"

End Sub


Public Sub ProgressUpdate( _
    ByVal completed As Long, _
    Optional ByVal detail As String = "")
'
'==============================================================================
'                                ProgressUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Report progress for the active session in the status bar.
'
' INPUTS
'   completed: steps done; zero or greater, at most total when total > 0.
'   detail:    optional text shown after the figures.
'
' SIDE EFFECTS
'   Replace the status-bar text; nothing else changes.
'
' ERROR POLICY
'   Raise PROGRESS_ERROR_INACTIVE without a session, PROGRESS_ERROR_CANCELLED
'   after a cancel request, and PROGRESS_ERROR_INVALID_ARGUMENT for an
'   out-of-range count. The session stays active; call ProgressEnd.
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

        ProgressCore.UpdateSession completed, detail
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "ProgressFacade.ProgressUpdate"

End Sub


Public Function ProgressEnd() _
    As Boolean
'
'==============================================================================
'                                 ProgressEnd
'------------------------------------------------------------------------------
' PURPOSE
'   End the active session and restore every owned property.
'
' RETURNS
'   Return True when a session was ended and False when none was active, so
'   the call is safe in every exit path and error handler. Restores the five
'   owned properties in reverse order and clears any cancel request.
'
' ERROR POLICY
'   Every property is restored even if one fails; the first restoration
'   error is then re-raised with this procedure as its source.
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

        ProgressEnd = ProgressCore.EndSession()
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "ProgressFacade.ProgressEnd"

End Function


Public Sub ProgressRecover()
'
'==============================================================================
'                               ProgressRecover
'------------------------------------------------------------------------------
' PURPOSE
'   Make the owned Excel UI state usable after an interrupted run.
'
' SIDE EFFECTS
'   End an active session exactly like ProgressEnd. Without one, for example
'   after the VBA project was reset and the snapshot lost, set StatusBar
'   False, DisplayStatusBar True, Cursor xlDefault (Windows), ScreenUpdating
'   True and EnableCancelKey xlInterrupt.
'
' ERROR POLICY
'   Property errors are re-raised with this procedure as their source.
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

        ProgressCore.RecoverSession
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "ProgressFacade.ProgressRecover"

End Sub


Public Sub ProgressRequestCancel()
'
'==============================================================================
'                            ProgressRequestCancel
'------------------------------------------------------------------------------
' PURPOSE
'   Ask the active session to stop at its next update.
'
' SIDE EFFECTS
'   The next ProgressUpdate raises PROGRESS_ERROR_CANCELLED. Call this from
'   the handler that receives error 18 when the user presses Esc.
'
' ERROR POLICY
'   Raise PROGRESS_ERROR_INACTIVE without a session.
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

        ProgressCore.RequestCancel
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Re-raise the core failure with this procedure as its public source.
        RaiseFromCore "ProgressFacade.ProgressRequestCancel"

End Sub


Public Function ProgressIsCancelled() _
    As Boolean
'
'==============================================================================
'                             ProgressIsCancelled
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether the active session has a pending cancel request.
'
' RETURNS
'   True between ProgressRequestCancel and the end of the session.
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
        ProgressIsCancelled = ProgressCore.IsCancelRequested()

End Function


Public Function ProgressIsActive() _
    As Boolean
'
'==============================================================================
'                               ProgressIsActive
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether a progress session owns the Excel UI state.
'
' RETURNS
'   True from a successful ProgressBegin until ProgressEnd or ProgressRecover.
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
        ProgressIsActive = ProgressCore.IsSessionActive()

End Function


Public Function ProgressElapsedSeconds() _
    As Double
'
'==============================================================================
'                            ProgressElapsedSeconds
'------------------------------------------------------------------------------
' PURPOSE
'   Measure the time since the active session began.
'
' RETURNS
'   Seconds, never negative, correct across midnight; 0 without a session.
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
        ProgressElapsedSeconds = ProgressCore.ElapsedSeconds()

End Function


Public Function ProgressText() _
    As String
'
'==============================================================================
'                                 ProgressText
'------------------------------------------------------------------------------
' PURPOSE
'   Return the status-bar text for the session's current figures.
'
' RETURNS
'   "caption: P% (C of T)" or "caption: C done", then " - detail" and
'   " - N.N s" when set; empty without a session.
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
        ProgressText = ProgressCore.CurrentText()

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
'   source: the public procedure name, such as "ProgressFacade.ProgressBegin".
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
