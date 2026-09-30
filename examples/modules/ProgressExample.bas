Attribute VB_Name = "ProgressExample"
'==============================================================================
' MODULE: ProgressExample
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate the complete caller pattern for a progress session: begin,
'   update in a loop, turn Esc into a cancel request, and end in every exit
'   path.
'
' PUBLIC SURFACE
'   RunProgressExample is an in-project example macro, not production API.
'
' DEPENDENCIES
'   ProgressFacade only.
'
' STATE OWNERSHIP
'   No module state. The progress session owns the status bar, cursor,
'   screen updating and Esc handling only between begin and end; the result
'   is written to the VBE Immediate window.
'
' ERROR POLICY
'   Esc (error 18) becomes a cancel request and the loop stops at its next
'   update. Any other error ends the session before it is re-raised.
'
' WORKSHEET SAFETY
'   Does not read or modify workbooks, worksheets, ranges, selection,
'   calculation, events or alerts.
'
' TEST SEAM
'   ProgressTests covers the same lifecycle with deterministic assertions.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS; the cursor changes on Windows only.
'
' USAGE
'   Import the ui-component production modules first, then run
'   ProgressExample.RunProgressExample from the VBE Immediate window. Press
'   Esc while it runs to see cancellation.
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

Public Sub RunProgressExample()
'
'==============================================================================
'                              RunProgressExample
'------------------------------------------------------------------------------
' PURPOSE
'   Sum the squares of 1 to 25 while reporting each step in the status bar.
'
' USAGE
'   Run ProgressExample.RunProgressExample from the VBE Immediate window.
'
' SIDE EFFECTS
'   The status bar shows the progress, the elapsed time and the current step
'   while the loop runs; the caller's status bar, cursor, screen updating and
'   Esc handling are restored afterwards. One result line is printed.
'
' ERROR POLICY
'   Esc cancels cleanly. Other errors end the session, then propagate.
'
' UPDATED
'   2026-09-30
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the work deterministic: an uncancelled run always prints 5525.
    Const STEPS              As Long = 25       'Loop length and progress total
    Dim completed            As Long            'Steps finished before any cancel
    Dim savedDescription     As String          'Unexpected error diagnostic for re-raise
    Dim savedNumber          As Long            'Unexpected error number for re-raise
    Dim savedSource          As String          'Unexpected error source for re-raise
    Dim stepIndex            As Long            'Current step, 1 to STEPS
    Dim sumOfSquares         As Long            'Running result of the demonstration work

'------------------------------------------------------------------------------
' RUN WITH PROGRESS
'------------------------------------------------------------------------------
    'Install the handler before the session starts, so every failure below
    'passes through the path that ends it. DoEvents lets Excel repaint the
    'status bar and notice Esc.
        On Error GoTo HandleError

        ProgressFacade.ProgressBegin "Summing squares", STEPS, ProgressShowElapsed
        For stepIndex = 1 To STEPS
            sumOfSquares = sumOfSquares + stepIndex * stepIndex
            completed = stepIndex
            ProgressFacade.ProgressUpdate stepIndex, "step " & CStr(stepIndex)
            DoEvents
        Next stepIndex

'------------------------------------------------------------------------------
' END AND REPORT
'------------------------------------------------------------------------------
CleanExit:
    'End the session on success and after cancellation alike; ending an
    'already ended session is a harmless no-op.
        ProgressFacade.ProgressEnd
        Debug.Print "Completed " & CStr(completed) & " of " & CStr(STEPS) & _
            "; sum of squares = " & CStr(sumOfSquares)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Esc raises error 18 wherever the loop happens to be; record the request
    'and continue, so the next update stops the loop through the documented
    'cancellation error.
        If Err.Number = 18 Then
            ProgressFacade.ProgressRequestCancel
            Resume Next
        ElseIf Err.Number = ProgressFacade.PROGRESS_ERROR_CANCELLED Then
            Debug.Print "Cancelled by the user."
            Resume CleanExit
        End If

    'Anything else is unexpected: give the Excel state back first, then
    'let the caller see the original error.
        savedNumber = Err.Number
        savedSource = Err.Source
        savedDescription = Err.Description
        On Error GoTo 0
        ProgressFacade.ProgressEnd
        Err.Raise savedNumber, savedSource, savedDescription

End Sub
