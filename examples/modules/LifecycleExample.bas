Attribute VB_Name = "LifecycleExample"
'==============================================================================
' MODULE: LifecycleExample
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate the application lifecycle end to end on a temporary host:
'   start, operate through an explicit target range, repeat the start, stop,
'   and close the host without saving. It is the reference smoke check.
'
' PUBLIC SURFACE
'   RunLifecycleExample is an in-project example macro, not production API.
'
' DEPENDENCIES
'   LifecycleFacade only.
'
' STATE OWNERSHIP
'   No module state. The example owns one temporary workbook from creation
'   to close; the lifecycle owns three hidden names inside it. Results are
'   written to the VBE Immediate window.
'
' ERROR POLICY
'   Any error stops the session this example started and closes the
'   temporary workbook before it is re-raised. A session another workbook
'   started is never stopped here.
'
' WORKSHEET SAFETY
'   Writes only to its temporary workbook, gives activation back to the
'   workbook that was active, and never saves anything.
'
' TEST SEAM
'   LifecycleTests covers the same lifecycle with deterministic assertions.
'
' COMPATIBILITY
'   Excel VBA on Windows and macOS.
'
' USAGE
'   Import the application production modules first, then run
'   LifecycleExample.RunLifecycleExample from the VBE Immediate window.
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


'
'------------------------------------------------------------------------------
'
'                             EXAMPLE ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunLifecycleExample()
'
'==============================================================================
'                             RunLifecycleExample
'------------------------------------------------------------------------------
' PURPOSE
'   Run one complete application session on a temporary host.
'
' USAGE
'   Run LifecycleExample.RunLifecycleExample from the VBE Immediate window.
'
' SIDE EFFECTS
'   Open and close one temporary workbook; print five lines.
'
' ERROR POLICY
'   Errors stop the session this example started and close the host, then
'   propagate. If another workbook's session is running, the start is
'   refused and that session is left alone.
'
' UPDATED
'   2026-11-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host and the workbook to give activation back to.
    Dim callerBook         As Workbook    'Workbook active before the example
    Dim host               As Workbook    'Temporary host for the session
    Dim savedDescription   As String      'Unexpected error diagnostic for re-raise
    Dim savedNumber        As Long        'Unexpected error number for re-raise
    Dim savedSource        As String      'Unexpected error source for re-raise
    Dim started            As Boolean     'True while this example's session runs
    Dim stopped            As Boolean     'Result of LifecycleStop
    Dim target             As Range       'Top-left cell of the status block

'------------------------------------------------------------------------------
' RUN SESSION
'------------------------------------------------------------------------------
    'Install the handler before the host exists, so every failure below
    'passes through the path that stops the session and closes the host.
        On Error GoTo HandleError

        Set callerBook = Application.ActiveWorkbook
        Set host = Application.Workbooks.Add(xlWBATWorksheet)
        If Not callerBook Is Nothing Then callerBook.Activate
        Set target = host.Worksheets(1).Range("A1")

        Debug.Print "LIFECYCLE EXAMPLE"
        started = LifecycleFacade.LifecycleStart(host)
        Debug.Print "Start: " & YesNo(started) & _
            "; status=" & LifecycleFacade.LifecycleStatus() & _
            "; schema=" & CStr(LifecycleFacade.LifecycleStoredSchema(host)) & _
            "; starts=" & CStr(LifecycleFacade.LifecycleStoredStarts(host))
        LifecycleFacade.LifecycleWriteStatus target
        Debug.Print "Status block: " & CStr(target.Value) & "=" & CStr(target.Offset(0, 1).Value) & _
            "; " & CStr(target.Offset(1, 0).Value) & "=" & CStr(target.Offset(1, 1).Value) & _
            "; " & CStr(target.Offset(2, 0).Value) & "=" & CStr(target.Offset(2, 1).Value)
        Debug.Print "Repeat start: " & YesNo(LifecycleFacade.LifecycleStart(host))
        stopped = LifecycleFacade.LifecycleStop()
        started = False
        Debug.Print "Stop: " & YesNo(stopped) & _
            "; status=" & LifecycleFacade.LifecycleStatus()

'------------------------------------------------------------------------------
' CLOSE HOST
'------------------------------------------------------------------------------
    'The host was only a stage for the session; nothing is saved.
        host.Close SaveChanges:=False
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
HandleError:
    'Give everything back first, then let the caller see the original error.
        savedNumber = Err.Number
        savedSource = Err.Source
        savedDescription = Err.Description
        On Error GoTo 0
        If started Then LifecycleFacade.LifecycleStop
        If Not host Is Nothing Then host.Close SaveChanges:=False
        Err.Raise savedNumber, savedSource, savedDescription

End Sub


'
'------------------------------------------------------------------------------
'
'                               PRIVATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Function YesNo( _
    ByVal value As Boolean) _
    As String
'
'==============================================================================
'                                    YesNo
'------------------------------------------------------------------------------
' PURPOSE
'   Spell a Boolean result in fixed words.
'
' RETURNS
'   "yes" or "no". CStr would follow the Office language, and the printed
'   lines must be identical on every installation.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' SPELL
'------------------------------------------------------------------------------
    'Fixed words keep the smoke output comparable across hosts.
        If value Then
            YesNo = "yes"
        Else
            YesNo = "no"
        End If

End Function
