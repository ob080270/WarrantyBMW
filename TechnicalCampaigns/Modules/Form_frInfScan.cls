VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frInfScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ========================================================================================
' Module Name      : Form_frInfScan
' Purpose          : Provides functionality for managing customer notifications
'                    related to technical campaigns, including data entry and
'                    interaction with parent forms.
'
' Parent Form       : frTA ("Letters" tab, subform with letters list, camera button)
' Child Forms       : None

' Key Functions    :
'   - Copies fields from parent form and subform (cmdCopyFields_Click).
'   - Tracks the loading and unloading state of the form (Form_Load, Form_Unload).
'   - Checks if the parent form is open (ParentFormIsOpen).
' Notes            :
'   - This form is designed to add customer notification data into the database.
'   - Interaction is managed through parent form "frTA" and subform "sfLetters".
' Revision History :
'   - Created on: 2008-07-25
'   - Updated on: 2025-01-03 by Oleh Bondarenko - Added comments for uploading to GitHub.
' =======================================================================================
Option Compare Database
Option Explicit
' ----------------------------------------------------------------------------
' Subroutine Name : cmdCopyFields_Click
' Purpose         : Populates the "Stock Number" and "VIN" fields
'                   in the current form with data from the "frTA" main form.
'                   The stock number is retrieved from the main form field,
'                   and the VIN is fetched from the subform "sfLetters".
' Preconditions   : "frTA" main form and "sfLetters" subform must be open.
' ----------------------------------------------------------------------------
Private Sub cmdCopyFields_Click()

    Me!infTA = Forms!frTA!aNo                           ' - Populate the "Stock Number" field with the current value from "frTA".
    Me!infVIN = Forms!frTA!sfLetters.Form!avVIN         ' - Populate the "VIN" field with the current value from the "sfLetters" subform.
    
End Sub
' -------------------------------------------------------------------
' Subroutine Name : Form_Load
' Purpose         : Initializes the form when it loads.
'                   Ensures the ToggleLink control in the parent form is updated.
' Error Handling  : Displays an error message and resumes execution gracefully.
' -------------------------------------------------------------------
Sub Form_Load()
On Error GoTo Form_Load_Err

    ' Check if the parent form is open; update ToggleLink control if true.
    If ParentFormIsOpen() Then Me.Parent!ToggleLink = True

Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox Error$
    Resume Form_Load_Exit

End Sub
' -------------------------------------------------------------------
' Subroutine Name: Form_Unload
' Purpose         : Cleans up the form state when it is unloaded.
'                   Ensures the ToggleLink control in the parent form is reset.
' Error Handling  : Displays an error message and resumes execution gracefully.
' -------------------------------------------------------------------
Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Err

    ' Check if the parent form is open; reset ToggleLink control if true.
    If ParentFormIsOpen() Then Me.Parent!ToggleLink = False

Form_Unload_Exit:
    Exit Sub

Form_Unload_Err:
    MsgBox Error$
    Resume Form_Unload_Exit

End Sub
' -------------------------------------------------------------------
' Function Name  : ParentFormIsOpen
' Purpose        : Checks if the parent form "sfLetters" is currently open.
' Returns        : Boolean - True if the parent form is open; False otherwise.
' -------------------------------------------------------------------
Private Function ParentFormIsOpen()

    ParentFormIsOpen = (SysCmd(acSysCmdGetObjectState, acForm, "sfLetters") And acObjStateOpen) <> False

End Function

