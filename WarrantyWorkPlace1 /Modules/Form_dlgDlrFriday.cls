VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgDlrFriday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name  : dlgDlrFriday
' Description  : Dialog form for selecting a dealer to create reports for claims
'                that are ready to be sent. The dialog is called from the
'                frClaims form via a hyperlink on the cmdSendRep button.
'
' Parent Form  : frClaims
' Child Forms  : None
'
' Key Features:
' 1. Allows multi-selection of dealers for generating reports.
' 2. Displays the list of selected dealers in a textbox.
' 3. Executes the report generation process for selected dealers and closes the form.
' 4. Includes an option to clear selections.
'
' External Functions:
' - fnFridayReport (Module: GlobalFn)
'   Description: Generates a report for a specific dealer and returns a status code.
'
' Events:
' 1. cmdCancelReport_Click   : Cancels report creation and closes the form.
' 2. cmdClear_Click          : Clears all selections from the list box.
' 3. cmdMakeReport_Click     : Generates reports for the selected dealers and closes the form.
' 4. fldDlrSlct_AfterUpdate  : Updates the textbox with the list of selected dealers.
'
' Developer    : Oleh Bondarenko
' Created      : 2007-02-15
' Last Updated : 2025-01-11 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Explicit
Option Compare Database
' -------------------------------------------------------------------
' Event #1       : cmdCancelReport_Click
' Purpose        : Cancels report creation and closes the dialog form.
' Behavior       : Closes the form without performing any actions.
' -------------------------------------------------------------------
Private Sub cmdCancelReport_Click()

    DoCmd.Close                                                         ' - Close the current form

End Sub
' -------------------------------------------------------------------
' Event #2       : cmdClear_Click
' Purpose        : Clears all selections from the list box.
' Behavior       : Iterates through selected items and deselects them.
'                  Resets the textbox displaying selected dealers.
' -------------------------------------------------------------------
Private Sub cmdClear_Click()

    Dim varValue As Variant
    For Each varValue In fldDlrSlct.ItemsSelected
        fldDlrSlct.Selected(varValue) = False                           ' - Deselect each selected item
    Next varValue
    txtSelected = ""                                                    ' - Clear the textbox content

End Sub
' -------------------------------------------------------------------
' Event #3       : cmdMakeReport_Click
' Purpose        : Generates reports for the selected dealers.
' Behavior       : Iterates through the selected dealers in the list box,
'                  calls the fnFridayReport function, and then closes the form.
' -------------------------------------------------------------------
Private Sub cmdMakeReport_Click()

    Dim varItem As Variant
    Dim btFnSt As Byte
    
    With fldDlrSlct
        For Each varItem In .ItemsSelected
            btFnSt = fnFridayReport(.Column(0, varItem))                ' - Generate report for each selected dealer
        Next varItem
    End With
    
    DoCmd.Close acForm, "dlgDlrFriday"                                  ' - Close the dialog form
    
End Sub
' -------------------------------------------------------------------
' Event #4       : fldDlrSlct_AfterUpdate
' Purpose        : Updates the textbox with the list of selected dealers.
' Behavior       : Retrieves dealer names from the list box based on selected items.
'                  Updates the textbox with a formatted list.
' -------------------------------------------------------------------
Private Sub fldDlrSlct_AfterUpdate()
' Print the list of selected items to the textbox txtSelected.

    Dim varItem As Variant
    Dim strList As String
    
    With fldDlrSlct
        If .MultiSelect = 0 Then
            txtSelected = .Value                                        ' - Single selection: update textbox directly
        Else
            For Each varItem In .ItemsSelected
                strList = strList & .Column(1, varItem) & vbCrLf        ' - Add dealer name to list
            Next varItem
            txtSelected = strList                                       ' - Update the textbox with the complete list
        End If
    End With
 
End Sub
