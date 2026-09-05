Attribute VB_Name = "ProjectCore"
Option Explicit
Option Private Module

'===============================================================================
' MODULE: ProjectCore
'-------------------------------------------------------------------------------
' RESPONSIBILITY
'   Implement the neutral arithmetic example behind the supported facade.
'
' PUBLIC SURFACE
'   None outside this VBA project. DivideChecked and ERR_ZERO_DENOMINATOR are
'   Public only for in-project use; Option Private Module keeps them off the
'   supported external surface.
'
' DEPENDENCIES
'   VBA runtime only. This core never depends on the facade, tests, examples,
'   workbook objects, or optional references.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on explicit scalar arguments.
'
' ERROR POLICY
'   Own the single internal zero-denominator error number and raise it with a
'   stable description. The facade exposes the same value and normalizes the
'   public error source.
'
' WORKSHEET SAFETY
'   Performs no Excel object-model access and changes no caller-owned state.
'
' TEST SEAM
'   ProjectTests verifies behavior through ProjectFacade. Direct core access is
'   available inside the project for focused future tests without widening API.
'===============================================================================

Public Const ERR_ZERO_DENOMINATOR As Long = vbObjectError + 2048

Public Function DivideChecked( _
    ByVal numerator As Double, _
    ByVal denominator As Double) As Double

    If denominator = 0# Then
        Err.Raise _
            ERR_ZERO_DENOMINATOR, _
            "ProjectCore.DivideChecked", _
            "Denominator must not be zero."
    End If

    DivideChecked = numerator / denominator
End Function
