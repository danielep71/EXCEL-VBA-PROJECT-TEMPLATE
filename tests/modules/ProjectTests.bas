Attribute VB_Name = "ProjectTests"
Option Explicit
Option Private Module

'===============================================================================
' MODULE: ProjectTests
'-------------------------------------------------------------------------------
' RESPONSIBILITY
'   Run a deterministic, dependency-free regression suite for the neutral
'   facade/core starter and report complete evidence to the Immediate window.
'
' PUBLIC SURFACE
'   RunProjectTests is the single documented test entry point.
'
' DEPENDENCIES
'   ProjectFacade and the built-in VBA/Excel object models only. No external
'   references, workbook fixture, worksheet, donor project, or test framework.
'
' STATE OWNERSHIP
'   Owns only private counters, failure text, and a re-entry flag. The cleanup
'   path resets all owned state and verifies the flag. No Excel state is changed.
'
' ERROR POLICY
'   Expected facade errors are captured and asserted. Unexpected runner errors
'   are recorded, cleanup runs, and the original number/source/description is
'   re-raised. Assertion failures raise one test-suite error after reporting.
'
' WORKSHEET SAFETY
'   Reads Application properties only to report the environment. It never
'   creates, selects, activates, edits, calculates, saves, or closes Excel state.
'
' TEST SEAM
'   Tests the supported ProjectFacade surface. The fixed core boundary can be
'   exercised by future focused tests without adding production API.
'===============================================================================

Private Const TEST_ERROR_DIRTY_START As Long = vbObjectError + 2060
Private Const TEST_ERROR_FAILURES As Long = vbObjectError + 2061
Private Const EXPECTED_ASSERTIONS As Long = 6
Private Const EXPECTED_CASES As Long = 4

Private mCaseCount As Long
Private mAssertionCount As Long
Private mFailureCount As Long
Private mFailureDetails As String
Private mRunActive As Boolean
Private mSuiteCompleted As Boolean

Public Sub RunProjectTests()
    Dim cleanupDetail As String
    Dim cleanupPassed As Boolean
    Dim savedDescription As String
    Dim savedNumber As Long
    Dim savedSource As String

    If mRunActive Then
        Debug.Print "RESULT=FAIL_DIRTY_START; cleanup=NOT_RUN"
        Err.Raise _
            TEST_ERROR_DIRTY_START, _
            "ProjectTests.RunProjectTests", _
            "A ProjectTests run is already active."
    End If

    ResetRun
    mRunActive = True
    On Error GoTo RunFailed

    PrintEnvironment
    TestExactEquality
    TestTolerance
    TestExpectedError
    TestRepeatability
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

CleanExit:
    On Error GoTo 0
    cleanupPassed = CleanupRun(cleanupDetail)
    PrintSummary cleanupPassed, cleanupDetail

    If savedNumber <> 0 Then
        Err.Raise savedNumber, savedSource, savedDescription
    End If

    If mFailureCount <> 0 Or Not cleanupPassed Or Not mSuiteCompleted Then
        Err.Raise _
            TEST_ERROR_FAILURES, _
            "ProjectTests.RunProjectTests", _
            "Regression failed; review the Immediate window report."
    End If
    Exit Sub

RunFailed:
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

Private Sub TestExactEquality()
    On Error GoTo CaseFailed

    BeginCase "ratio.exact"
    AssertEqualDouble _
        "ProjectRatio(10, 4)", _
        2.5, _
        ProjectFacade.ProjectRatio(10#, 4#)
    Exit Sub

CaseFailed:
    RecordUnexpectedCaseError "ratio.exact"
End Sub

Private Sub TestTolerance()
    On Error GoTo CaseFailed

    BeginCase "ratio.tolerance"
    AssertNear _
        "ProjectRatio(1, 3)", _
        0.333333333333333, _
        ProjectFacade.ProjectRatio(1#, 3#), _
        0.000000000000001
    Exit Sub

CaseFailed:
    RecordUnexpectedCaseError "ratio.tolerance"
End Sub

Private Sub TestExpectedError()
    Dim actualDescription As String
    Dim actualNumber As Long
    Dim actualSource As String
    Dim ignored As Double

    BeginCase "ratio.zero-denominator"
    On Error GoTo ExpectedError

    ignored = ProjectFacade.ProjectRatio(1#, 0#)
    RecordFailure _
        "ratio.zero-denominator.raises", _
        "Expected an error, but the call returned " & CStr(ignored) & "."
    Exit Sub

ExpectedError:
    actualNumber = Err.Number
    actualSource = Err.Source
    actualDescription = Err.Description
    On Error GoTo 0

    AssertExpectedError _
        "ratio.zero-denominator", _
        ProjectFacade.PROJECT_ERROR_ZERO_DENOMINATOR, _
        "ProjectFacade.ProjectRatio", _
        "Denominator must not be zero.", _
        actualNumber, _
        actualSource, _
        actualDescription
End Sub

Private Sub TestRepeatability()
    Dim firstResult As Double
    Dim secondResult As Double

    On Error GoTo CaseFailed

    BeginCase "ratio.repeatability"
    firstResult = ProjectFacade.ProjectRatio(-9#, 4#)
    secondResult = ProjectFacade.ProjectRatio(-9#, 4#)
    AssertEqualDouble "Repeated calls", firstResult, secondResult
    Exit Sub

CaseFailed:
    RecordUnexpectedCaseError "ratio.repeatability"
End Sub

Private Sub BeginCase(ByVal caseName As String)
    mCaseCount = mCaseCount + 1
    Debug.Print "CASE=" & caseName
End Sub

Private Sub AssertEqualDouble( _
    ByVal assertionName As String, _
    ByVal expected As Double, _
    ByVal actual As Double)

    mAssertionCount = mAssertionCount + 1
    If actual <> expected Then
        RecordFailure _
            assertionName, _
            "expected=" & CStr(expected) & "; actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertNear( _
    ByVal assertionName As String, _
    ByVal expected As Double, _
    ByVal actual As Double, _
    ByVal tolerance As Double)

    mAssertionCount = mAssertionCount + 1
    If tolerance < 0# Then
        RecordFailure assertionName, "Tolerance must not be negative."
    ElseIf Abs(actual - expected) > tolerance Then
        RecordFailure _
            assertionName, _
            "expected=" & CStr(expected) & _
                "; actual=" & CStr(actual) & _
                "; tolerance=" & CStr(tolerance)
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

    mAssertionCount = mAssertionCount + 1
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
        RecordFailure _
            assertionName, _
            "expected=""" & expected & """; actual=""" & actual & """"
    End If
End Sub

Private Sub RecordUnexpectedCaseError(ByVal caseName As String)
    Dim errorDescription As String
    Dim errorNumber As Long
    Dim errorSource As String

    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error GoTo 0

    RecordFailure _
        caseName & ".unexpected", _
        "error=" & CStr(errorNumber) & _
            "; source=" & errorSource & _
            "; description=" & errorDescription
End Sub

Private Sub RecordFailure( _
    ByVal assertionName As String, _
    ByVal detail As String)

    mFailureCount = mFailureCount + 1
    If Len(mFailureDetails) > 0 Then
        mFailureDetails = mFailureDetails & vbNewLine
    End If
    mFailureDetails = mFailureDetails & assertionName & ": " & detail
End Sub

Private Function CleanupRun(ByRef cleanupDetail As String) As Boolean
    On Error GoTo CleanupFailed

    mRunActive = False
    CleanupRun = Not mRunActive
    If CleanupRun Then
        cleanupDetail = "owned module state restored; Excel state changed=no"
    Else
        cleanupDetail = "owned re-entry flag remained set"
    End If
    Exit Function

CleanupFailed:
    cleanupDetail = _
        "cleanup error=" & CStr(Err.Number) & _
        "; description=" & Err.Description
    Err.Clear
    CleanupRun = False
End Function

Private Sub PrintEnvironment()
    Debug.Print "PROJECT TESTS"
    Debug.Print "ENVIRONMENT=" & EnvironmentSummary()
End Sub

Private Function EnvironmentSummary() As String
    Dim bitness As String
    Dim vbaGeneration As String

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

    Dim verdict As String

    If mFailureCount = 0 And cleanupPassed And mSuiteCompleted Then
        verdict = "PASS"
    Else
        verdict = "FAIL"
    End If

    Debug.Print "CASES=" & CStr(mCaseCount)
    Debug.Print "ASSERTIONS=" & CStr(mAssertionCount)
    Debug.Print "FAILURES=" & CStr(mFailureCount)
    Debug.Print "CLEANUP=" & IIf(cleanupPassed, "PASS", "FAIL") & _
        "; detail=" & cleanupDetail

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

Private Sub ResetRun()
    mCaseCount = 0
    mAssertionCount = 0
    mFailureCount = 0
    mFailureDetails = vbNullString
    mSuiteCompleted = False
End Sub
