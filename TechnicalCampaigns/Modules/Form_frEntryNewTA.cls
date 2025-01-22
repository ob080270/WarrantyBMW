VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frEntryNewTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ================================================================================================================
' Module Name       : frEntryNewTA
' Purpose           : Manages the user interface for the entry and management of Technical Campaigns.
'                     Includes features for conditional formatting, navigation, and application exit.
'
' Parent Form       : frTA (CommandButton Name ="cmdEnterData" HyperlinkSubAddress ="Form frEntryNewTA")
' Child Forms       :
'   - tblAV subform - Allows entering the list of VINs for vehicles affected by the Technical Campaign.
'   - sqTA_Vcl      - Displays related information about the vehicle, as well as the status of campaign execution.
'
' Opening Trigger   :
'   - This form is opened via the cmdEnterData button in the frTA form.
'
' Events            :
'   1. Form_Current()       - Adjusts the interface to visually distinguish current campaigns from stopped ones.
'   2. cmdPartsOrd_Click()  - Opens the "frTA" form filtered by the currently selected Technical Campaign
'                             for further ordering of parts and writing letters to customers.
'   3. cmdExitApp_Click()   - Closes the entire application (Exits Application).
'
' Notes:
'   - The module is tightly coupled with the frEntryNewTA form.
'   - Ensure proper error handling in user-initiated events.
'
' Developer      : Oleh Bondarenko
' Created        : 2008-07-25
' Last Updated   : 2025-01-10 by Oleh Bondarenko - Added comments about Parent & Child Forms
' ===================================================================================================================
Option Compare Database
Option Explicit

Private Sub Form_Current()
' Adjusts the interface to visually distinguish current campaigns from stopped ones

    If Me.aStop Then                                        ' Formatting for stopped campaigns:
        Me.aNo.ForeColor = 8421504                              ' - Grey color for campaign number
        Forms!frEntryNewTA!‡Descr.ForeColor = 8421504           ' - Grey color for campaign description
        Me.lbRecall.ForeColor = 8421504                         ' - Grey color for "Recall" label
        Me.lbStop.ForeColor = 255                               ' - Red color for "Stop" label
        Me.lbStop.FontBold = True                               ' - Bold font for "Stop" label
    Else                                                    ' Formatting for active campaigns:
        Me.aNo.ForeColor = 16711680                             ' - Red color for campaign number
        Forms!frEntryNewTA!‡Descr.ForeColor = 16711680          ' - Red color for campaign description
        Me.lbRecall.ForeColor = 255                             ' - Red color for "Recall" label
        Me.lbStop.ForeColor = -2147483630                       ' - Default color for "Stop" label
        Me.lbStop.FontBold = False                              ' - Regular font for "Stop" label
    End If
    
End Sub

Private Sub cmdPartsOrd_Click()
' Opens the "frTA" form filtered by the selected Technical Campaign number

On Error GoTo Err_cmdPartsOrd_Click

    Dim stDocName As String                                     ' - Name of the form to open
    Dim stLinkCriteria As String                                ' - Filter criteria for the form

    stDocName = "frTA"                                          ' - Form name (main application form)
    
    stLinkCriteria = "[aNo]=" & "'" & Me![aNo] & "'"            ' - Build filter to open the form for the selected campaign
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    MsgBox "The form has been filtered by: " & vbCrLf & stLinkCriteria, _
           vbInformation, "Filter Applied"
    Form_frTA.insTabComtrol.Value = 1                           ' - Set the tab to "Spare parts"

Exit_cmdPartsOrd_Click:
    Exit Sub

Err_cmdPartsOrd_Click:
    MsgBox Err.Description                                      ' - Handle errors by displaying a message box
    Resume Exit_cmdPartsOrd_Click
    
End Sub

Private Sub cmdExitApp_Click()

On Error GoTo Err_cmdExitApp_Click

    DoCmd.Close                                                  ' - Close current form

Exit_cmdExitApp_Click:
    Exit Sub

Err_cmdExitApp_Click:
    MsgBox Err.Description
    Resume Exit_cmdExitApp_Click
    
End Sub
