VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgDlrActWarr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==========================================================================================================
' Module Name  : dlgDlrActWarr
' Description  : Dialog form for confirming dealer activity data before generating the report dlgDlrActWarr.
'                Part of the process for handling warranty compensation.
'
' Parent Form  : frClaims
' Child Forms  : None
'
' Key Features:
' 1. Validates dealer selection before proceeding with report generation.
' 2. Provides user interaction enhancements like dropdown auto-opening.
'
' Events:
' 1. cmdMakeReport_Click    : Validates inputs, hides the form, and triggers
'                             price recalculation queries.
' 2. fldDlrSlct_GotFocus    : Automatically opens the dropdown menu for easier
'                             dealer selection.
' 3. cmdCancelReport_Click  : Cancels the operation and closes the form.
'
' Developer    : Oleh Bondarenko
' Created      : 2016-04-14
' Last Updated : 2025-01-09 by Oleh Bondarenko - Added comments for GitHub upload
' ==========================================================================================================
Option Compare Database
Option Explicit
' -----------------------------------------------------------------------------------------
' Event #1          : cmdMakeReport_Click
' Purpose           : Generates a report for the selected dealer with the specified date.
'                     Validates that a dealer is selected before proceeding.
' -----------------------------------------------------------------------------------------
Private Sub cmdMakeReport_Click()

On Error GoTo Err_cmdOKbt_Click
    
    If IsNull(fldDlrSlct) Then              ' - Ensure that a dealer has been selected before proceeding
        MsgBox "Выберите дилера"            ' - Notify the user to select a dealer
        fldDlrSlct.SetFocus                 ' - Set focus on the dealer selection dropdown field
    Else
        Me.Visible = False                  ' - Hide the form if validation passes
    End If

Exit_cmdOKbt_Click:
    Exit Sub

Err_cmdOKbt_Click:
    MsgBox Err.Description
    Resume Exit_cmdOKbt_Click
    
End Sub
' ------------------------------------------------------------------------------
' Event #2          : fldDlrSlct_GotFocus
' Purpose           : Automatically opens the dropdown menu for dealer selection
'                     when the field gains focus, improving user interaction.
' ------------------------------------------------------------------------------
Private Sub fldDlrSlct_GotFocus()

    fldDlrSlct.Dropdown                     ' - Open the dropdown menu for easier user selection
    
End Sub
' ------------------------------------------------------------------------------
' Event #3          : cmdCancelReport_Click
' Purpose           : Cancels the report creation process and closes the current form.
' ------------------------------------------------------------------------------
Private Sub cmdCancelReport_Click()

    DoCmd.Close acForm, "dlgDlrActWarr"     ' - Close the form without generating the report

End Sub


