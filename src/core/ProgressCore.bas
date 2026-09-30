Attribute VB_Name = "ProgressCore"
'==============================================================================
' MODULE: ProgressCore
'------------------------------------------------------------------------------
' PURPOSE
'   Own the single progress session behind ProgressFacade: take the Excel UI
'   state a long operation needs, report progress in the status bar, and give
'   every changed property back exactly as it was found.
'
' PUBLIC SURFACE
'   None outside this VBA project. The Public procedures and constants below
'   serve ProgressFacade and focused in-project tests; Option Private Module
'   keeps them off the supported external surface.
'
' DEPENDENCIES
'   VBA runtime and the Excel Application object. On Windows, kernel32
'   QueryPerformanceCounter/QueryPerformanceFrequency for elapsed time.
'
' STATE OWNERSHIP
'   Owns one session: an active flag, the caption, counts, options, a cancel
'   request, the start time and a snapshot of five Application properties.
'   During a session it owns StatusBar, DisplayStatusBar, Cursor (Windows),
'   ScreenUpdating and EnableCancelKey. Calculation, EnableEvents,
'   DisplayAlerts, selection, workbooks and worksheets stay caller-owned and
'   are never read or changed.
'
' ERROR POLICY
'   Own four error numbers: session already active, no session active,
'   session cancelled, and invalid argument. A failure while taking state
'   restores whatever was already changed before the error is re-raised.
'
' WORKSHEET SAFETY
'   Never selects, activates, edits, calculates or saves a workbook or
'   worksheet. Only the five owned Application properties change.
'
' TEST SEAM
'   ProgressTests drives the session through ProgressFacade and compares the
'   owned Application properties before, during and after each session.
'
' COMPATIBILITY
'   Excel VBA on Windows (VBA6 and VBA7, 32- and 64-bit) and macOS. Windows
'   times sessions with the performance counter; macOS uses Timer with the
'   midnight wrap corrected and leaves the cursor unchanged.
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
' NATIVE DECLARATIONS
'------------------------------------------------------------------------------
    'Timer counts seconds since midnight and wraps, so an overnight session
    'would report a negative time. The Windows performance counter never
    'wraps; Currency receives its 64-bit value on every bitness, and the
    'scaling cancels in the counter/frequency ratio.
#If Mac Then
#ElseIf VBA7 Then
        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
#Else
        Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
        Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
#End If

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Own the session error codes and the option flags used by the facade.
        Public Const ERR_SESSION_ACTIVE     As Long = vbObjectError + 2200    'Begin while a session is active
        Public Const ERR_SESSION_INACTIVE   As Long = vbObjectError + 2201    'Update without an active session
        Public Const ERR_CANCELLED          As Long = vbObjectError + 2202    'Update after a cancel request
        Public Const ERR_INVALID_ARGUMENT   As Long = vbObjectError + 2203    'Rejected caption, count or option
        Public Const OPTION_KEEP_SCREEN     As Long = 1                       'Leave ScreenUpdating as found
        Public Const OPTION_SHOW_ELAPSED    As Long = 2                       'Append elapsed seconds to the text
        Public Const OPTION_ALL             As Long = 3                       'Every defined option flag

'------------------------------------------------------------------------------
' MODULE TYPES
'------------------------------------------------------------------------------
    'Hold the five owned Application properties exactly as found.
        Private Type SessionSnapshot
            StatusBar          As Variant              'Caller text, or False when Excel owns the bar
            DisplayStatusBar   As Boolean              'Caller status-bar visibility
            Cursor             As XlMousePointer       'Caller cursor; not used on macOS
            ScreenUpdating     As Boolean              'Caller screen-updating setting
            EnableCancelKey    As XlEnableCancelKey    'Caller Esc/Ctrl+Break handling
        End Type

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Describe the one active session; everything resets when it ends.
        Private mActive            As Boolean            'True from a successful begin until end
        Private mCaption           As String             'Caption shown before the progress figures
        Private mCancelRequested   As Boolean            'Set by RequestCancel; cleared by end
        Private mCompleted         As Long               'Latest completed count
        Private mDetail            As String             'Latest optional detail text
        Private mOptions           As Long               'Option flags for this session
        Private mSnapshot          As SessionSnapshot    'Owned properties as found at begin
        Private mStartCounter      As Currency           'Performance counter at begin (Windows)
        Private mStartTimer        As Double             'Timer at begin (macOS)
        Private mTotal             As Long               'Total steps; 0 means indeterminate


'
'------------------------------------------------------------------------------
'
'                              SESSION LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Sub BeginSession( _
    ByVal caption As String, _
    ByVal total As Long, _
    ByVal options As Long)
'
'==============================================================================
'                                 BeginSession
'------------------------------------------------------------------------------
' PURPOSE
'   Start the progress session and take the owned Application state.
'
' INPUTS
'   caption: nonblank text shown before the progress figures.
'   total:   number of steps; 0 shows a completed count without a percentage.
'   options: OPTION_KEEP_SCREEN and/or OPTION_SHOW_ELAPSED, or 0.
'
' SIDE EFFECTS
'   Snapshot the five owned properties, then show the status bar with the
'   first progress text, set the wait cursor (Windows), turn screen updating
'   off unless kept, and route Esc to the caller's error handler as error 18.
'
' ERROR POLICY
'   Raise ERR_INVALID_ARGUMENT or ERR_SESSION_ACTIVE before changing any
'   state. If taking state fails, restore what was already changed and
'   re-raise the original error with no session left active.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the failure fields so a partial begin can be rolled back first.
    Dim savedDescription   As String     'Original diagnostic for re-raise
    Dim savedNumber        As Long       'Original error number for re-raise
    Dim savedSource        As String     'Original error source for re-raise
    Dim snapshotComplete   As Boolean    'True once all five properties were read

'------------------------------------------------------------------------------
' VALIDATE REQUEST
'------------------------------------------------------------------------------
    'Refuse a second session before touching state: nested sessions would
    'overwrite the first snapshot and lose the caller's values.
        If Len(Trim$(caption)) = 0 Then
            RaiseCore ERR_INVALID_ARGUMENT, "ProgressCore.BeginSession", "caption must not be empty."
        End If
        If total < 0 Then
            RaiseCore ERR_INVALID_ARGUMENT, "ProgressCore.BeginSession", "total must be zero or greater."
        End If
        If (options And Not OPTION_ALL) <> 0 Or options < 0 Then
            RaiseCore ERR_INVALID_ARGUMENT, "ProgressCore.BeginSession", _
                "options must combine only ProgressKeepScreenUpdating and ProgressShowElapsed."
        End If
        If mActive Then
            RaiseCore ERR_SESSION_ACTIVE, "ProgressCore.BeginSession", "A progress session is already active."
        End If

'------------------------------------------------------------------------------
' SNAPSHOT CALLER STATE
'------------------------------------------------------------------------------
    'Read every owned property before changing any, so restoration always
    'has a complete picture of the caller's state.
        On Error GoTo BeginFailed

        mSnapshot.StatusBar = Application.StatusBar
        mSnapshot.DisplayStatusBar = Application.DisplayStatusBar
#If Not Mac Then
        mSnapshot.Cursor = Application.Cursor
#End If
        mSnapshot.ScreenUpdating = Application.ScreenUpdating
        mSnapshot.EnableCancelKey = Application.EnableCancelKey
        snapshotComplete = True

'------------------------------------------------------------------------------
' TAKE OWNED STATE
'------------------------------------------------------------------------------
    'Record the session before changing Excel so the text can be built,
    'then apply the owned values; the session is active only once all apply.
        mCaption = caption
        mTotal = total
        mCompleted = 0
        mDetail = vbNullString
        mOptions = options
        mCancelRequested = False
        StartClock

        Application.DisplayStatusBar = True
        Application.StatusBar = CurrentText()
#If Not Mac Then
        Application.Cursor = xlWait
#End If
        If (options And OPTION_KEEP_SCREEN) = 0 Then Application.ScreenUpdating = False
        Application.EnableCancelKey = xlErrorHandler
        mActive = True
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE BEGIN ERROR
'------------------------------------------------------------------------------
BeginFailed:
    'Save the failure, give back anything already changed, then re-raise;
    'no session remains active after a failed begin. A failure while still
    'reading changed nothing, and an incomplete snapshot must never be
    'written back.
        savedNumber = Err.Number
        savedSource = Err.Source
        savedDescription = Err.Description
        On Error GoTo 0
        If snapshotComplete Then RestoreSnapshot
        ResetSession
        Err.Raise savedNumber, savedSource, savedDescription

End Sub


Public Sub UpdateSession( _
    ByVal completed As Long, _
    ByVal detail As String)
'
'==============================================================================
'                                UpdateSession
'------------------------------------------------------------------------------
' PURPOSE
'   Report progress for the active session in the status bar.
'
' INPUTS
'   completed: steps done; zero or greater, and at most total when total is
'              greater than zero.
'   detail:    optional text appended after the figures; may be empty.
'
' SIDE EFFECTS
'   Replace the status-bar text. No other property changes.
'
' ERROR POLICY
'   Raise ERR_SESSION_INACTIVE without a session, ERR_CANCELLED after a
'   cancel request, and ERR_INVALID_ARGUMENT for an out-of-range count. The
'   session stays active in every case; the caller still ends it.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE REQUEST
'------------------------------------------------------------------------------
    'Check cancellation before the count, so a cancelled loop stops at its
    'next update whatever value it passes.
        If Not mActive Then
            RaiseCore ERR_SESSION_INACTIVE, "ProgressCore.UpdateSession", "No progress session is active."
        End If
        If mCancelRequested Then
            RaiseCore ERR_CANCELLED, "ProgressCore.UpdateSession", "The progress session was cancelled."
        End If
        If completed < 0 Or (mTotal > 0 And completed > mTotal) Then
            RaiseCore ERR_INVALID_ARGUMENT, "ProgressCore.UpdateSession", _
                "completed must be zero or greater and at most total."
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Keep the figures in module state so CurrentText stays the single
    'formatter for begin, update and callers.
        mCompleted = completed
        mDetail = detail
        Application.StatusBar = CurrentText()

End Sub


Public Function EndSession() _
    As Boolean
'
'==============================================================================
'                                  EndSession
'------------------------------------------------------------------------------
' PURPOSE
'   End the active session and restore every owned property.
'
' RETURNS
'   True when a session was ended; False when none was active, which makes
'   the call safe in any exit path or error handler.
'
' SIDE EFFECTS
'   Restore EnableCancelKey, ScreenUpdating, Cursor (Windows), StatusBar
'   and DisplayStatusBar to their snapshot values, in reverse order of
'   taking them, and clear the session, including any cancel request.
'
' ERROR POLICY
'   Every property is restored even if one fails; the session is cleared
'   and then the first restoration error is re-raised.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Carry the first restoration failure past the session reset.
    Dim failedDescription   As String    'First restoration diagnostic
    Dim failedNumber        As Long      'First restoration error number, or 0
    Dim failedSource        As String    'First restoration error source

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'A call without a session is an intentional no-op for cleanup paths.
        If Not mActive Then
            EndSession = False
            Exit Function
        End If
        RestoreSnapshot failedNumber, failedSource, failedDescription
        ResetSession
        EndSession = True

'------------------------------------------------------------------------------
' REPORT RESTORATION FAILURE
'------------------------------------------------------------------------------
    'The session is already cleared, so a caller can retry with
    'RecoverSession after seeing the error.
        If failedNumber <> 0 Then
            Err.Raise failedNumber, failedSource, failedDescription
        End If

End Function


Public Sub RecoverSession()
'
'==============================================================================
'                                RecoverSession
'------------------------------------------------------------------------------
' PURPOSE
'   Return the owned Application state to a usable condition after an
'   interrupted run.
'
' SIDE EFFECTS
'   With an active session, behave exactly like EndSession. Without one
'   (for example after the VBA project was reset, which discards the
'   snapshot), set the documented Excel defaults: StatusBar False,
'   DisplayStatusBar True, Cursor xlDefault (Windows), ScreenUpdating True
'   and EnableCancelKey xlInterrupt. The cancel request is cleared.
'
' ERROR POLICY
'   Property errors propagate; there is no session state left to protect.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RECOVER
'------------------------------------------------------------------------------
    'Prefer the exact snapshot; defaults are the fallback when it is gone.
        If mActive Then
            EndSession
            Exit Sub
        End If
        Application.EnableCancelKey = xlInterrupt
        Application.ScreenUpdating = True
#If Not Mac Then
        Application.Cursor = xlDefault
#End If
        Application.StatusBar = False
        Application.DisplayStatusBar = True
        ResetSession

End Sub


'
'------------------------------------------------------------------------------
'
'                             SESSION OBSERVATION
'
'------------------------------------------------------------------------------
'

Public Sub RequestCancel()
'
'==============================================================================
'                                RequestCancel
'------------------------------------------------------------------------------
' PURPOSE
'   Ask the active session to stop at its next update.
'
' SIDE EFFECTS
'   Set the cancel request; the next UpdateSession raises ERR_CANCELLED.
'   Call it from the caller's handler for error 18 (Esc) or from any UI.
'
' ERROR POLICY
'   Raise ERR_SESSION_INACTIVE without a session.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' REQUEST
'------------------------------------------------------------------------------
    'A request without a session would be silently lost, so it is refused.
        If Not mActive Then
            RaiseCore ERR_SESSION_INACTIVE, "ProgressCore.RequestCancel", "No progress session is active."
        End If
        mCancelRequested = True

End Sub


Public Function IsCancelRequested() _
    As Boolean
'
'==============================================================================
'                              IsCancelRequested
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether the active session has a pending cancel request.
'
' RETURNS
'   True between RequestCancel and the end of the session; otherwise False.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'The flag is cleared whenever the session ends.
        IsCancelRequested = mCancelRequested

End Function


Public Function IsSessionActive() _
    As Boolean
'
'==============================================================================
'                               IsSessionActive
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether a progress session currently owns the Application state.
'
' RETURNS
'   True from a successful begin until the session ends or is recovered.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Expose the lifecycle flag without letting callers change it.
        IsSessionActive = mActive

End Function


Public Function ElapsedSeconds() _
    As Double
'
'==============================================================================
'                                ElapsedSeconds
'------------------------------------------------------------------------------
' PURPOSE
'   Measure the time since the active session began.
'
' RETURNS
'   Seconds as a Double, never negative; 0 without an active session.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Read the platform clock into locals before computing the difference.
#If Mac Then
    Dim seconds      As Double      'Timer difference, corrected for midnight
#Else
    Dim counterNow   As Currency    'Current counter value, scaled by Currency
    Dim frequency    As Currency    'Counter ticks per second, scaled by Currency
#End If

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Without a session there is no start time to measure from.
        If Not mActive Then
            ElapsedSeconds = 0#
            Exit Function
        End If
#If Mac Then
        seconds = Timer - mStartTimer
        If seconds < 0# Then seconds = seconds + 86400#
        ElapsedSeconds = seconds
#Else
        QueryPerformanceCounter counterNow
        QueryPerformanceFrequency frequency
        If frequency > 0 Then
            ElapsedSeconds = CDbl(counterNow - mStartCounter) / CDbl(frequency)
        Else
            ElapsedSeconds = 0#
        End If
#End If

End Function


Public Function CurrentText() _
    As String
'
'==============================================================================
'                                 CurrentText
'------------------------------------------------------------------------------
' PURPOSE
'   Build the status-bar text for the session's current figures.
'
' RETURNS
'   "caption: P% (C of T)" when total is greater than zero, otherwise
'   "caption: C done"; then " - detail" when a detail is set and
'   " - N.N s" when elapsed time is shown. Empty without a session.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Build the text in one place so begin, update and callers agree.
    Dim percent      As Long      'Whole percentage, rounded down
    Dim statusText   As String    'Status text built so far

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Compute the percentage in Double so large totals cannot overflow Long;
    'rounding down means 100% appears only when every step is done.
        If Len(mCaption) = 0 Then
            CurrentText = vbNullString
            Exit Function
        End If
        If mTotal > 0 Then
            percent = Int(CDbl(mCompleted) * 100# / CDbl(mTotal))
            statusText = mCaption & ": " & CStr(percent) & "% (" & CStr(mCompleted) & " of " & CStr(mTotal) & ")"
        Else
            statusText = mCaption & ": " & CStr(mCompleted) & " done"
        End If
        If Len(mDetail) > 0 Then statusText = statusText & " - " & mDetail
        If (mOptions And OPTION_SHOW_ELAPSED) <> 0 Then
            statusText = statusText & " - " & Format$(ElapsedSeconds(), "0.0") & " s"
        End If
        CurrentText = statusText

End Function


'
'------------------------------------------------------------------------------
'
'                               PRIVATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub StartClock()
'
'==============================================================================
'                                  StartClock
'------------------------------------------------------------------------------
' PURPOSE
'   Record the session start on this platform's clock.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RECORD START
'------------------------------------------------------------------------------
    'Only the clock that ElapsedSeconds reads on this platform is set.
#If Mac Then
        mStartTimer = Timer
#Else
        QueryPerformanceCounter mStartCounter
#End If

End Sub


Private Sub RestoreSnapshot( _
    Optional ByRef failedNumber As Long, _
    Optional ByRef failedSource As String, _
    Optional ByRef failedDescription As String)
'
'==============================================================================
'                               RestoreSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Put back every owned property from the snapshot, best effort.
'
' RETURNS
'   The first restoration error through the optional ByRef arguments; 0
'   and empty text when every property was restored.
'
' ERROR POLICY
'   Resume Next keeps one failed property from blocking the others; the
'   first failure is captured, never lost.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RESTORE IN REVERSE ORDER
'------------------------------------------------------------------------------
    'Undo in the reverse order of BeginSession so the status bar is the last
    'thing the caller sees change.
        On Error Resume Next
        Application.EnableCancelKey = mSnapshot.EnableCancelKey
        CaptureFirstError failedNumber, failedSource, failedDescription
        Application.ScreenUpdating = mSnapshot.ScreenUpdating
        CaptureFirstError failedNumber, failedSource, failedDescription
#If Not Mac Then
        Application.Cursor = mSnapshot.Cursor
        CaptureFirstError failedNumber, failedSource, failedDescription
#End If
        Application.StatusBar = mSnapshot.StatusBar
        CaptureFirstError failedNumber, failedSource, failedDescription
        Application.DisplayStatusBar = mSnapshot.DisplayStatusBar
        CaptureFirstError failedNumber, failedSource, failedDescription
        On Error GoTo 0

End Sub


Private Sub CaptureFirstError( _
    ByRef failedNumber As Long, _
    ByRef failedSource As String, _
    ByRef failedDescription As String)
'
'==============================================================================
'                              CaptureFirstError
'------------------------------------------------------------------------------
' PURPOSE
'   Keep the first error raised under Resume Next, then clear Err.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
    'Later failures are cleared too, so each restore step starts clean.
        If Err.Number <> 0 Then
            If failedNumber = 0 Then
                failedNumber = Err.Number
                failedSource = Err.Source
                failedDescription = Err.Description
            End If
            Err.Clear
        End If

End Sub


Private Sub ResetSession()
'
'==============================================================================
'                                 ResetSession
'------------------------------------------------------------------------------
' PURPOSE
'   Return all session state to idle.
'
' STATE OWNERSHIP
'   Clear the active flag, caption, counts, detail, options and cancel
'   request. The snapshot is left in place until the next begin overwrites
'   it; it is never read without an active session.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Clear everything together so no value leaks into the next session.
        mActive = False
        mCaption = vbNullString
        mCancelRequested = False
        mCompleted = 0
        mDetail = vbNullString
        mOptions = 0
        mTotal = 0

End Sub


Private Sub RaiseCore( _
    ByVal number As Long, _
    ByVal source As String, _
    ByVal description As String)
'
'==============================================================================
'                                  RaiseCore
'------------------------------------------------------------------------------
' PURPOSE
'   Raise one of the core session errors with a stated rule.
'
' ERROR POLICY
'   Always raises; ProgressFacade replaces the source.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' RAISE
'------------------------------------------------------------------------------
    'Keep the numeric contract in the constants above.
        Err.Raise number, source, description

End Sub
