VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sqParts_Parts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name               : sqParts_Parts
' Description               : Subform for adding spare parts required to fulfill a Technical Campaign
'                             for a single vehicle. Automates the process of creating a parts order
'                             by cloning sets of spare parts, adjusting part numbers, and updating totals.
'
' Parent Form               : frTA ("Spare parts" tab)
' Child Forms               : None
'
' Key Features              :
' 1. Automates cloning of spare parts sets for all vehicles affected by a campaign.
' 2. Removes spaces from part numbers when added from clipboard.
' 3. Updates part totals dynamically after record insertion.
'
' Events                    :
' 1. cmdMultiply_Click      : Executes a query to clone spare parts sets.
' 2. fldPND_AfterUpdate     : Removes spaces from part numbers upon user input.
' 3. Form_AfterInsert       : Recalculates total parts in the subform sfPartSum.
'
' Developer                 : Oleh Bondarenko
' Created                   : 2008-07-25
' Last Updated              : 2025-01-07 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Event #1        : cmdMultiply_Click
' Purpose         : Executes a query to "clone" the set of spare parts required
'                   for the current vehicle's Technical Campaign for all vehicles.
' Behavior        : Automates the creation of a spare parts order by replicating
'                   the current set of parts.
' -------------------------------------------------------------------
Private Sub cmdMultiply_Click()

    DoCmd.OpenQuery ("qaPartsMulti")                    ' - Execute the cloning query

End Sub
' -------------------------------------------------------------------
' Event #2        : fldPND_AfterUpdate
' Purpose         : Removes spaces from part numbers copied from external sources.
'                   Converts part numbers into a consistent format without spaces.
' Behavior        : Removes all spaces in the fldPND field and updates fldPartNo.
' -------------------------------------------------------------------
Private Sub fldPND_AfterUpdate()

    Me!fldPartNo = Replace(Me!fldPND, " ", "")          ' Remove spaces and update fldPartNo

End Sub
' -------------------------------------------------------------------
' Event #3        : Form_AfterInsert
' Purpose         : Recalculates the total parts quantities displayed in the
'                   related subform sfPartSum.
' Behavior        : Refreshes the sfPartSum subform to reflect the inserted parts.
' -------------------------------------------------------------------
Private Sub Form_AfterInsert()

    Forms!frTA!sfPartSum.Form.Requery                   ' Requery the subform to update totals

End Sub
