VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frCompareQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================================================================================================
' Module Name  : frCompareQ
' Description  : Manages reconciliation of discrepancies in warranty claims for AW (After Work) and spare parts quantities.
'                Includes functionality for filtering, importing, and updating claims data for BMW dealers.
' Parent Form  : frClaims
'
' Key Features:
' 1. Automatically handles discrepancies in AW and spare parts quantities.
' 2. Dynamically updates the state of form controls based on data changes.
' 3. Integrates with external modules to manage claims and discrepancies.
'
' Events:
' 1. Form_Close : Finalizes the form state, enabling or disabling specific controls based on the presence of discrepancies.
'
' Dependencies:
' 1. BSI_Class  - Manages BSI-specific actions and states.
' 2. DivQ_Class - Handles discrepancies in quantities of AW and spare parts.
'
' Notes         : Ensure that the parent form "frClaims" is active and contains the required controls for seamless integration.
'
' Developer     : Oleh Bondarenko
' Created       : 2017-06-13
' Last Updated  : 2025-01-20 by Oleh Bondarenko - Added comments for GitHub upload
' =====================================================================================================================
Option Compare Database
Option Explicit
Private objBSIAct As BSI    ' - managing BSI actions and associated states
Private objDivQ As DivQ     ' - managing discrepancies in AW and parts quantities

' ---------------------------------------------------------------------------------------------------------------------
' Event #1        : Form_Close
' Purpose         : Ensures the appropriate state of form controls related to discrepancy handling and claim submission.
' Behavior        :
'                   - Initializes objects for handling discrepancies and BSI actions.
'                   - Updates the state of the "cmdDivQ" and "cmdDlrAct" buttons on the parent form based on discrepancies.
'                   - Refreshes the "cmdActBSI" button state through external methods.
' External Calls:
'                   - objDivQ.RecQ (Property of Class: DivQ)            : Checks for quantity discrepancies.
'                   - objBSIAct.mtBtnStatBSI (Method of Class: BSI)     : Updates the status of the BSI button.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub Form_Close()
'If there are records with a discrepancy in the quantity field - close the button for printing dealer reports,
'and open the button for printing dealer reports to correct the quantity

    Set objDivQ = New DivQ                          ' - Class for handling AW and spare parts quantity discrepancies
    Set objBSIAct = New BSI                         ' - Class for managing BSI actions
    
    Forms!frClaims!clComment.SetFocus
    
'   Check for discrepancies in quantities and update control states accordingly:
    If objDivQ.RecQ > 0 Then
        Forms!frClaims!cmdDivQ.Enabled = True       ' - Enable discrepancy resolution button
        Forms!frClaims!cmdDlrAct.Enabled = False    ' - Disable dealer action button
        Forms!frClaims!cmdActBSI.Enabled = False    ' - Disable BSI action button
    Else
        Forms!frClaims!cmdDivQ.Enabled = False      ' - Disable discrepancy resolution button
        Forms!frClaims!cmdDlrAct.Enabled = True     ' - Enable dealer action button
        objBSIAct.mtBtnStatBSI                      ' - Update BSI action button state
    End If
    
End Sub
