Attribute VB_Name = "ProjectFacade"
Option Explicit

'===============================================================================
' MODULE: ProjectFacade
'-------------------------------------------------------------------------------
' RESPONSIBILITY
'   Provide the neutral supported entry point and translate core failures into
'   a stable caller-facing error contract.
'
' PUBLIC SURFACE
'   PROJECT_ERROR_ZERO_DENOMINATOR
'   ProjectRatio(numerator, denominator) As Double
'
' DEPENDENCIES
'   ProjectCore only. The dependency direction is facade -> core.
'
' STATE OWNERSHIP
'   Stateless. This module owns no Application, workbook, worksheet, Range, UI,
'   callback, file-system, or module-level mutable state.
'
' ERROR POLICY
'   A zero denominator raises PROJECT_ERROR_ZERO_DENOMINATOR with this facade as
'   the source. Other core errors retain their number, description, help file,
'   and help context while the public source is normalized to this facade.
'
' WORKSHEET SAFETY
'   Uses only explicit scalar arguments. It never reads Application.Caller,
'   ActiveWorkbook, ActiveSheet, Selection, or other ambient Excel state.
'
' TEST SEAM
'   ProjectTests exercises this public surface. ProjectCore remains separately
'   addressable inside the VBA project without becoming supported public API.
'===============================================================================

Public Const PROJECT_ERROR_ZERO_DENOMINATOR As Long = vbObjectError + 2048

Public Function ProjectRatio( _
    ByVal numerator As Double, _
    ByVal denominator As Double) As Double

    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedHelpContext As Long
    Dim savedHelpFile As String

    On Error GoTo HandleError

    ProjectRatio = ProjectCore.DivideChecked(numerator, denominator)
    Exit Function

HandleError:
    savedNumber = Err.Number
    savedDescription = Err.Description
    savedHelpFile = Err.HelpFile
    savedHelpContext = Err.HelpContext

    Err.Raise _
        savedNumber, _
        "ProjectFacade.ProjectRatio", _
        savedDescription, _
        savedHelpFile, _
        savedHelpContext
End Function
