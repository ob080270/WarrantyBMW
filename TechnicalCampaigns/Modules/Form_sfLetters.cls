VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name  : sfLetters
' Description  : Subform for managing the process of sending letters to customers
'                about active Technical Campaigns. Includes logic for toggling and
'                filtering the related child form "frInfScan".
'
' Parent Form  : frTA ("Letters" tab)
' Child Form   : frInfScan
'
' Key Features:
' 1. Automatically updates the sending date based on the marker status (ltSentFlag).
' 2. Opens, closes, and filters the related child form dynamically.
' 3. Implements error handling for user actions and data synchronization.
'
' Events:
' 1. ltSentFlag_AfterUpdate : Updates sending date when marker status changes.
' 2. Form_Current           : Handles actions when the current record changes.
' 3. ToggleLink_Click       : Toggles the state of the child form.
'
' Methods:
' 4. FilterChildForm        : Filters the child form based on the current record.
' 5. OpenChildForm          : Opens the child form "frInfScan".
' 6. CloseChildForm         : Closes the child form "frInfScan".
' 7. ChildFormIsOpen        : Checks if the child form is currently open.
'
' Notes        : Ensure the parent form "frTA" is open for proper interaction
'                between forms.
'
' Developer    : Oleh Bondarenko
' Created      : 2008-07-25
' Last Updated : 2025-01-05 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Event #1        : ltSentFlag_AfterUpdate
' Purpose         : Automatically updates the sending date based on
'                   the mail sending marker.
' Behavior        : If the mail sending marker is True, sets the current
'                   date as the sending date. Otherwise, resets it to Null.
' -------------------------------------------------------------------
Private Sub ltSentFlag_AfterUpdate()

    If Me!ltSentFlag = True Then
        Me!ltSentDate = Date                            ' - Set current date
    Else
        Me!ltSentDate = Null                            ' - Reset to Null if marker is False
    End If

End Sub
' -------------------------------------------------------------------
' Event #2        : Form_Current
' Purpose         : Executes when the form's current record changes.
'                   Filters the child form if it is open.
' Error Handling  : Displays an error message if an exception occurs.
' -------------------------------------------------------------------
Sub Form_Current()
On Error GoTo Form_Current_Err

    If ChildFormIsOpen() Then FilterChildForm           ' - Apply filter to child form

Form_Current_Exit:
    Exit Sub

Form_Current_Err:
    MsgBox Error$
    Resume Form_Current_Exit

End Sub
' -------------------------------------------------------------------
' Event #3        : ToggleLink_Click
' Purpose         : Toggles the visibility of the child form.
' Behavior        : Opens or closes the child form and applies filters if needed.
' Error Handling  : Displays an error message if an exception occurs.
' -------------------------------------------------------------------
Sub ToggleLink_Click()
On Error GoTo ToggleLink_Click_Err

    If ChildFormIsOpen() Then
        CloseChildForm                                  ' - Close if already open
    Else
        OpenChildForm                                   ' - Open if closed
        FilterChildForm                                 ' - Apply filter after opening
    End If

ToggleLink_Click_Exit:
    Exit Sub

ToggleLink_Click_Err:
    MsgBox Error$
    Resume ToggleLink_Click_Exit

End Sub
' -------------------------------------------------------------------
' Method #4       : FilterChildForm
' Purpose         : Applies a filter to the child form based on the current record.
' Behavior        : Enables data entry mode for new records or filters by TA and VIN.
' -------------------------------------------------------------------
Private Sub FilterChildForm()

    If Me.NewRecord Then
        'Enable data entry mode:
        Forms![frInfScan].DataEntry = True
    Else
        'Enable the filter:
        Forms![frInfScan].Filter = "[infTA] = " & """" & Me.[ltTA] & """" & " AND [infVIN] = " & """" & Me.[ltVIN] & """"
        Forms![frInfScan].FilterOn = True
    End If

End Sub
' -------------------------------------------------------------------
' Method #5       : OpenChildForm
' Purpose         : Opens the child form "frInfScan" and updates the toggle control.
' -------------------------------------------------------------------
Private Sub OpenChildForm()

    DoCmd.OpenForm "frInfScan"
    If Not Me.[ToggleLink] Then Me![ToggleLink] = True  ' - Update toggle state

End Sub
' -------------------------------------------------------------------
' Method #6       : CloseChildForm
' Purpose         : Closes the child form "frInfScan" and updates the toggle control.
' -------------------------------------------------------------------
Private Sub CloseChildForm()

    DoCmd.Close acForm, "frInfScan"
    If Me![ToggleLink] Then Me![ToggleLink] = False     ' Update toggle state

End Sub
' -------------------------------------------------------------------
' Method #7       : ChildFormIsOpen
' Purpose         : Checks if the child form "frInfScan" is currently open.
' Returns         : Boolean - True if the form is open, False otherwise.
' -------------------------------------------------------------------
Private Function ChildFormIsOpen()

    ChildFormIsOpen = (SysCmd(acSysCmdGetObjectState, acForm, "frInfScan") And acObjStateOpen) <> False

End Function

