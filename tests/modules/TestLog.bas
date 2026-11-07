Attribute VB_Name = "TestLog"
'==============================================================================
' MODULE: TestLog
'------------------------------------------------------------------------------
' PURPOSE
'   Print every line a test suite reports to the Immediate window and, while
'   a certification log is open, write the same line to that file. The suites
'   call Say instead of Debug.Print, so an automated run keeps each complete
'   report as evidence without reading the Immediate window.
'
' PUBLIC SURFACE
'   Say prints one line. OpenLog, CloseLog and CertificationReady are called
'   by tools\Certify-Excel.ps1 through Application.Run; a maintainer never
'   needs them. Option Private Module keeps all four out of the external
'   workbook automation API.
'
' DEPENDENCIES
'   The VBA runtime only: Debug.Print and the Open, Print # and Close
'   statements. No reference, worksheet or workbook state.
'
' STATE OWNERSHIP
'   Owns the open log: its file number, its path and the count of lines it
'   could not write. Nothing else; the suites own their own counters.
'
' ERROR POLICY
'   Say never raises: a line that cannot be written is counted and the
'   Immediate window still shows it. OpenLog refuses, instead of raising, an
'   existing file or a second log, so a failing report is never replaced.
'   CloseLog reports the count of lines that were lost.
'
' WORKSHEET SAFETY
'   Never reads or changes workbook, worksheet, selection or Application
'   state. The only side effect outside VBA is the log file the caller names.
'
' USAGE
'   Import with the test modules; they call TestLog.Say. For an evidence
'   bundle, run tools\Certify-Excel.ps1, which opens one log per suite
'   (docs/EXCEL_EVIDENCE.md).
'
' UPDATED
'   2026-11-07
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the log out of the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Describe the one open log; CloseLog clears all three.
        Private mFile       As Integer    'File number of the open log; 0 for none
        Private mPath       As String     'Path of the open log; empty for none
        Private mLost       As Long       'Lines that could not be written since OpenLog


'
'------------------------------------------------------------------------------
'
'                                  REPORTING
'
'------------------------------------------------------------------------------
'

Public Sub Say( _
    ByVal text As String)
'
'==============================================================================
'                                     Say
'------------------------------------------------------------------------------
' PURPOSE
'   Print one report line to the Immediate window and to the open log.
'
' INPUTS
'   text: the line, without a line end.
'
' SIDE EFFECTS
'   Append the line, with CRLF, to the open log when there is one.
'
' ERROR POLICY
'   Never raise: a line that cannot be written is counted in the lost
'   count that CloseLog reports.
'
' UPDATED
'   2026-11-07
'==============================================================================
'

'------------------------------------------------------------------------------
' PRINT
'------------------------------------------------------------------------------
    'The Immediate window always gets the line, log or no log.
        Debug.Print text
        If mFile = 0 Then Exit Sub

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Print # ends the line with CRLF; a write failure must not stop a suite.
        On Error GoTo WriteFailed
        Print #mFile, text
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE WRITE ERROR
'------------------------------------------------------------------------------
WriteFailed:
    'Count the loss; the validator rejects a log with lost lines.
        mLost = mLost + 1

End Sub


'
'------------------------------------------------------------------------------
'
'                                CERTIFICATION
'
'------------------------------------------------------------------------------
'

Public Function CertificationReady() _
    As String
'
'==============================================================================
'                              CertificationReady
'------------------------------------------------------------------------------
' PURPOSE
'   Answer READY, so the certification script can prove the project's code
'   runs before it opens a log.
'
' RETURNS
'   The text READY.
'
' UPDATED
'   2026-11-07
'==============================================================================
'

'------------------------------------------------------------------------------
' ANSWER
'------------------------------------------------------------------------------
    'A call that returns at all proves the macros run.
        CertificationReady = "READY"

End Function


Public Function OpenLog( _
    ByVal path As String, _
    ByVal candidate As String) _
    As String
'
'==============================================================================
'                                   OpenLog
'------------------------------------------------------------------------------
' PURPOSE
'   Create the log of one suite and write its first line, CANDIDATE= and the
'   commit SHA, so every later Say line is also written there.
'
' INPUTS
'   path: full path of the log to create; it must not exist.
'   candidate: the 40-character SHA of the checked-out commit.
'
' RETURNS
'   OPEN when the log was created; otherwise REFUSED; reason=... and no
'   log is open.
'
' ERROR POLICY
'   Refuse, never raise: a log that exists already, a log already open,
'   and a path that cannot be created are all reported in the result.
'
' UPDATED
'   2026-11-07
'==============================================================================
'

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'A report is never replaced, and two logs would mix two suites. Dir$
    'raises on a malformed path, so the handler is armed first.
        On Error GoTo OpenFailed
        If mFile <> 0 Then
            OpenLog = "REFUSED; reason=a log is already open: " & mPath
            Exit Function
        End If
        If Len(Dir$(path)) > 0 Then
            OpenLog = "REFUSED; reason=the log exists already: " & path
            Exit Function
        End If

'------------------------------------------------------------------------------
' OPEN
'------------------------------------------------------------------------------
    'Write the candidate line here, so the log never claims a typed SHA.
        mFile = FreeFile
        Open path For Output As #mFile
        Print #mFile, "CANDIDATE=" & candidate
        mPath = path
        mLost = 0
        OpenLog = "OPEN"
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OPEN ERROR
'------------------------------------------------------------------------------
OpenFailed:
    'Leave no half-open log behind.
        OpenLog = "REFUSED; reason=" & Err.Description
        On Error Resume Next
        If mFile <> 0 Then Close #mFile
        mFile = 0
        mPath = vbNullString

End Function


Public Function CloseLog() _
    As String
'
'==============================================================================
'                                   CloseLog
'------------------------------------------------------------------------------
' PURPOSE
'   Close the open log and report how many lines were lost.
'
' RETURNS
'   CLOSED; lost=N, where N counts the lines Say could not write; NONE when
'   no log was open.
'
' ERROR POLICY
'   Never raise: a close error is contained and the log is forgotten.
'
' UPDATED
'   2026-11-07
'==============================================================================
'

'------------------------------------------------------------------------------
' CLOSE
'------------------------------------------------------------------------------
    'Forget the log even when the close fails, so the next OpenLog works.
        If mFile = 0 Then
            CloseLog = "NONE"
            Exit Function
        End If
        On Error Resume Next
        Close #mFile
        On Error GoTo 0
        mFile = 0
        mPath = vbNullString
        CloseLog = "CLOSED; lost=" & CStr(mLost)

End Function
