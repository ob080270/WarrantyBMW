VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgSendRpQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
' Module Name  : dlgSendRpQuery
' Description  : Handles interactions for creating a report based on the
'                claim submission date.
' Parent Form  : rpBMWsnd
' Child Form   : None
'
' Key Features:
' 1. Provides options to select a claim submission date (today or previous).
' 2. Populates the combo box with the latest submission date.
' 3. Controls form behavior based on user input.
'
' Events:
' 1. Form_Open               : Populates the combo box with the latest date.
' 2. cmdCancelBt_Click       : Closes the form.
' 3. cmdOKbt_Click           : Processes the report based on the selected date.
' 4. swgrDateChoose_AfterUpdate : Enables/disables the date field based on user choice.
'
' Developer    : Oleh Bondarenko
' Created      : 2004-11-25
' Last Updated : 2025-01-12 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Event #1        : Form_Open
' Purpose         : Populates the combo box with the most recent date of claim submission.
' Behavior        : Fetches the latest date from the "tblBMWNo" table and sets it as the default value.
' -------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    Me!Sended = DMax("[sndSendDate]", "tblBMWNo")       ' - Populate the combo box with the latest submission date
    
End Sub
' -------------------------------------------------------------------
' Event #2        : cmdCancelBt_Click
' Purpose         : Closes the form when the cancel button is clicked.
' Behavior        : Invokes the DoCmd.Close method to close the form.
' -------------------------------------------------------------------
Private Sub cmdCancelBt_Click()

    DoCmd.Close                                         ' - Close the current form (dlgSendRpQuery)

End Sub
' -------------------------------------------------------------------
' Event #3        : cmdOKbt_Click
' Purpose         : Generates a report based on the selected submission date.
' Behavior        : Executes a query to set the submission date if the "Today" option is selected.
'                   Then hides the form to proceed with further processing.
' Query Details   : Updates the "sndSendDate" field in the "tblBMWNo" table to the current date (Date())
'                   for all records where the "sndSendDate" is NULL and the "clRegICS" field is TRUE.
' Error Handling  : Displays an error message if an exception occurs.
' -------------------------------------------------------------------
Private Sub cmdOKbt_Click()
On Error GoTo Err_cmdOKbt_Click
    
    If swgrDateChoose = 1 Then                          ' - if "Today" is selected
        DoCmd.SetWarnings False
            DoCmd.OpenQuery "qSetBMWsendingDate"        ' - execute the action query to update data
        DoCmd.SetWarnings True
        Sended = Date                                   ' - Set the current date to form field (Sended)
    End If

    Me.Visible = False                                  ' - Hide the form for further processing

Exit_cmdOKbt_Click:
    Exit Sub

Err_cmdOKbt_Click:
    MsgBox Err.Description
    Resume Exit_cmdOKbt_Click
    
End Sub
' -------------------------------------------------------------------
' Event #4        : swgrDateChoose_AfterUpdate
' Purpose         : Adjusts the enabled state of the combo box based on the selected option.
' Behavior        : Enables the date field when "Previous Date" is selected.
'                   Disables it otherwise.
' -------------------------------------------------------------------
Private Sub swgrDateChoose_AfterUpdate()

    If swgrDateChoose = 2 Then
        Sended.Enabled = True                           ' - Enable the combo box for date selection
    Else
        Sended.Enabled = False                          ' - Disable the combo box
    End If
    
End Sub
