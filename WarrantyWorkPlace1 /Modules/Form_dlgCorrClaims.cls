VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgCorrClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name  : dlgCorrClaims
' Description  : Handles the logic for the "Parts Selection" dialog form.
'                Includes functionality for generating reports, clearing selections,
'                and updating the interface dynamically.
'
' Parent Form  : frClaims
' Child Forms  : None
' Notes        : The form "dlgCorrClaims" is opened from the main application form
'                "frClaims" via the reference in the "refCorr" button.
'
' Key Features:
' 1. Dynamically updates the interface to reflect selected dealers in the form.
' 2. Generates custom reports based on user selections.
' 3. Implements error handling and supports both single and multi-selection modes.
'
' Events:
' 1. fldDlrSlct_AfterUpdate  : Updates the text box with the list of selected dealers.
' 2. cmdCancelReport_Click   : Closes the form without generating a report.
' 3. cmdClear_Click          : Clears the list of selected dealers and resets the display.
' 4. cmdMakeReport_Click     : Generates reports based on selected dealers and closes the form.
'
' Developer    : Oleh Bondarenko
' Created      : 2008-04-30
' Last Updated : 2025-01-09 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Explicit
Option Compare Database
' -------------------------------------------------------------------
' Subroutine Name: cmdCancelReport_Click
' Purpose         : Closes the form without generating a report.
' Notes           :
'   - This is triggered when the user clicks the "Cancel" button.
' -------------------------------------------------------------------
Private Sub cmdCancelReport_Click()

    DoCmd.Close                                                     ' - Close the current (dlgCorrClaims) form

End Sub
' -------------------------------------------------------------------
' Subroutine Name: cmdClear_Click
' Purpose         : Clears all selected items in the list box "fldDlrSlct"
'                   and resets the text box "txtSelected".
' Notes           :
'   - Iterates through the selected items in the list box to deselect them.
'   - Resets the text box to display no selected items.
' -------------------------------------------------------------------
Private Sub cmdClear_Click()

    Dim varValue As Variant
    
    For Each varValue In fldDlrSlct.ItemsSelected                   ' - Loop through all selected items in the list box and deselect them
        fldDlrSlct.Selected(varValue) = False
    Next varValue
    
    Me!txtSelected = ""                                             ' - Clear the text box showing selected items in form dlgCorrClaims

End Sub
' -------------------------------------------------------------------
' Subroutine Name: cmdMakeReport_Click
' Purpose         : Generates the report "rpCorrClaims" based on the selected
'                   dealers in the list box.
' Notes           :
'   - Calls the function "fnCorrClmRp" for each selected dealer.
'   - Closes the form after processing all selected dealers.
'   - The function "fnCorrClmRp" is defined in module "GlobalFn" and handles the report generation for each dealer.
' -------------------------------------------------------------------
Private Sub cmdMakeReport_Click()

    Dim varItem As Variant
    Dim btFnSt As Byte
    
    With fldDlrSlct
        For Each varItem In .ItemsSelected                          ' - Loop through all selected items in the list box
            btFnSt = fnCorrClmRp(.Column(0, varItem))               ' - Call the function (GlobalFn) to process each selected dealer
        Next varItem
    End With
    
    DoCmd.Close acForm, "dlgCorrClaims"                             ' - Close the form after generating the reports
    
End Sub
' -------------------------------------------------------------------
' Subroutine Name: fldDlrSlct_AfterUpdate
' Purpose         : Updates the text box "txtSelected" to display the names
'                   of all selected dealers from the list box.
' Notes           :
'   - Handles both single-selection and multi-selection modes of the list box.
'   - Displays the selected dealer names in a newline-separated format.
' -------------------------------------------------------------------
Private Sub fldDlrSlct_AfterUpdate()

    Dim varItem As Variant
    Dim strList As String
    
    With fldDlrSlct                                                 ' - Handle multi-select list box and update the text box with selected items
        If .MultiSelect = 0 Then                                    ' - For single-selection mode,
            Me!txtSelected = .Value                                 ' display the selected value
        Else
            For Each varItem In .ItemsSelected                      ' - For multi-selection mode, iterate through selected items
                strList = strList & .Column(1, varItem) & vbCrLf    ' - Retrieve the value from the second column of the selected row
            Next varItem
            Me!txtSelected = strList                                ' - Update the text box with the list of selected dealers
        End If
    End With
 
End Sub
