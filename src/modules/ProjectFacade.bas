Attribute VB_Name = "ProjectFacade"
'==============================================================================
' MODULE: ProjectFacade
'------------------------------------------------------------------------------
' PURPOSE
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
'   the source. The public constant aliases the core-owned internal error code,
'   so the numeric contract has one source of truth. Other core errors retain
'   their number, description, help file, and help context while the public
'   source is normalized to this facade.
'
' WORKSHEET SAFETY
'   Uses only explicit scalar arguments. It never reads Application.Caller,
'   ActiveWorkbook, ActiveSheet, Selection, or other ambient Excel state.
'
' TEST SEAM
'   ProjectTests exercises this public surface. ProjectCore remains separately
'   addressable inside the VBA project without becoming supported public API.
'
' COMPATIBILITY
'   Excel VBA; scalar VBA arithmetic requires no optional references.
'
' UPDATED
'   2026-09-09
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Public Const PROJECT_ERROR_ZERO_DENOMINATOR   As Long = ProjectCore.ERR_ZERO_DENOMINATOR


'
'------------------------------------------------------------------------------
'
'                             SUPPORTED PUBLIC API
'
'------------------------------------------------------------------------------
'

Public Function ProjectRatio( _
    ByVal numerator As Double, _
    ByVal denominator As Double) _
    As Double
'
'==============================================================================
'                                 ProjectRatio
'------------------------------------------------------------------------------
' PURPOSE
'   Expose checked division through the supported facade.
'
' INPUTS
'   numerator, denominator: explicit scalar Double values.
'
' RETURNS
'   Double quotient when the core call succeeds.
'
' ERROR POLICY
'   Normalize the error source to ProjectFacade.ProjectRatio; preserve the
'   number, description, help file, and help context supplied by the core.
'
' DEPENDENCIES
'   ProjectCore.DivideChecked. No host-state access.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim savedNumber        As Long
    Dim savedDescription   As String
    Dim savedHelpContext   As Long
    Dim savedHelpFile      As String

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
        On Error GoTo HandleError

        ProjectRatio = ProjectCore.DivideChecked(numerator, denominator)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
    'Capture every field before raising through the supported facade.
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
