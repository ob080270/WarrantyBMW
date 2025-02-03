VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgDlrAct_InsVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name  : dlgDlrAct_InsVar
' Description  : Form module for dealer action dialogs. Includes logic for
'                generating warranty compensation reports and managing
'                user interactions for dealer selection.
'
' Parent Form  : Ќе могу найти откуда форма вызываетс€ в проекте
'                похоже, эта форма не открываетс€ с главной формы
' Child Forms  : None
'
' Key Features:
' 1. Validates dealer selection before proceeding with report generation.
' 2. Automates the opening of dropdown lists for improved usability.
' 3. Provides error handling for user actions.
'
' Events:
' 1. cmdMakeReport_Click  : Handles the click event to generate reports.
' 2. fldDlrSlct_GotFocus  : Triggers dropdown opening when field gets focus.
' 3. cmdCancelReport_Click: Handles the click event to cancel the process.
'
' Developer    : Oleh Bondarenko
' Created      : 2011-03-22
' Last Updated : 2025-01-09 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Subroutine Name: cmdMakeReport_Click
' Purpose         : Generates a warranty compensation report for the selected dealer
'                   and specified date.
' Behavior        : Ensures that a dealer is selected before making the form invisible.
' Error Handling  : Displays an error message if an exception occurs.
' -------------------------------------------------------------------
Private Sub cmdMakeReport_Click()

On Error GoTo Err_cmdOKbt_Click
    
    If IsNull(fldDlrSlct) Then                          ' Validate that a dealer has been selected:
        MsgBox "Please select a dealer."                ' - Notify the user to select a dealer
        fldDlrSlct.SetFocus                             ' - Set focus on the dealer selection field
    Else
        Me.Visible = False                              ' - Hide the form if validation passes
    End If

Exit_cmdOKbt_Click:
    Exit Sub

Err_cmdOKbt_Click:
    MsgBox Err.Description                              ' - Display error details in a message box
    Resume Exit_cmdOKbt_Click                           ' - Resume normal flow after handling the error
    
End Sub
' -------------------------------------------------------------------
' Subroutine Name: fldDlrSlct_GotFocus
' Purpose         : Automatically opens the dropdown list when the dealer
'                   selection field gains focus.
' -------------------------------------------------------------------
Private Sub fldDlrSlct_GotFocus()

    fldDlrSlct.Dropdown                                 ' - Trigger the dropdown to open for easier user interaction
    
End Sub
' -------------------------------------------------------------------
' Subroutine Name: cmdCancelReport_Click
' Purpose         : Cancels the report creation and closes the form.
' Behavior        : Closes the "dlgDlrAct" form without generating the report.
' -------------------------------------------------------------------
Private Sub cmdCancelReport_Click()

    DoCmd.Close acForm, "dlgDlrAct"                     ' Close the current form

End Sub


