VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfPartSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module Name       : sfPartSum
' Parent Form       : frTA ("Spare parts" tab)
' Child Forms       : None
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Subroutine Name   : Form_Click
' Purpose           : Stores the count of selected rows in the form's
'                     Tag property. This value is later used when the
'                     "On Order" button (cmdMoveToOrd on frTA form)
'                     is clicked.
' Behavior          : Sets the form's Tag property to the current
'                     selection height (number of selected rows).
' Preconditions     : A selection must be made in the subform.
' Notes             : The value stored in the Tag property is essential
'                     for transferring the selected rows to the Order
'                     subform (sf2OrdItem).
' -------------------------------------------------------------------
Private Sub Form_Click()

    Me.Tag = Me.SelHeight       ' - Save the number of selected rows in the current form's Tag property
    
End Sub
