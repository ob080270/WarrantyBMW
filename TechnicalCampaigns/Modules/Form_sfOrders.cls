VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module Name     : sfOrders
' Parent Form       : frTA ("Spare parts" tab)
' Child Forms       : None
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Subroutine Name : Form_Current
' Purpose         : Synchronizes parent form fields with subform fields
'                   during navigation.
'                   - Updates txtLinkOrd1 on the parent form with the
'                     current ordTA value from the subform.
'                   - Updates txtLinkOrd2 on the parent form with the
'                     current ordNo value from the subform.
' Error Handling  : Ignores errors using 'On Error Resume Next' to
'                   prevent interruption due to potential null values.
' -------------------------------------------------------------------
Private Sub Form_Current()
On Error Resume Next                        ' - Ignore runtime errors
    Me.Parent("txtLinkOrd1") = [ordTA]      ' - Synchronize parent form control "txtLinkOrd1" with subform's "ordTA"
    Me.Parent("txtLinkOrd2") = [ordNo]      ' - Synchronize parent form control "txtLinkOrd2" with subform's "ordNo"
End Sub
