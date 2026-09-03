Attribute VB_Name = "ProjectExample"
Option Explicit
Option Private Module

'===============================================================================
' MODULE: ProjectExample
'-------------------------------------------------------------------------------
' RESPONSIBILITY
'   Demonstrate one supported facade call from explicit scalar inputs.
'
' PUBLIC SURFACE
'   RunProjectExample is an in-project example macro, not production API.
'
' DEPENDENCIES
'   ProjectFacade only.
'
' STATE OWNERSHIP
'   No mutable state. Output is written only to the VBE Immediate window.
'
' ERROR POLICY
'   Does not suppress facade errors; callers see the supported error contract.
'
' WORKSHEET SAFETY
'   Does not read or modify Application, workbook, worksheet, Range, selection,
'   calculation, events, display settings, or other host state.
'
' TEST SEAM
'   ProjectTests covers the same facade behavior with deterministic assertions.
'===============================================================================

Public Sub RunProjectExample()
    Const NUMERATOR As Double = 12#
    Const DENOMINATOR As Double = 4#

    Debug.Print "ProjectRatio(12, 4) = " & _
        CStr(ProjectFacade.ProjectRatio(NUMERATOR, DENOMINATOR))
End Sub
