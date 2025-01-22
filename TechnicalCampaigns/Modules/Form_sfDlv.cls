VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfDlv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module Name     : sfDlv
' Parent Form       : frTA ("Spare parts" tab)
' Child Forms       : None
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Subroutine Name : Form_Current
' Purpose         : Synchronizes parent form fields with subform fields
'                   during navigation.
'                   - Updates txtLinkDlv1 on the parent form with the
'                     current dlvTA value from the subform.
'                   - Updates txtLinkDlv2 on the parent form with the
'                     current dlvNo value from the subform.
' Error Handling  : Ignores errors using 'On Error Resume Next' to
'                   prevent interruption due to potential null values.
' -------------------------------------------------------------------
Private Sub Form_Current()
On Error Resume Next                            ' - Ignore runtime errors
    Me.Parent("txtLinkDlv1") = [dlvTA]          ' - Synchronize parent form control "txtLinkDlv1" with subform's "dlvTA"
    Me.Parent("txtLinkDlv2") = [dlvNo]          ' - Synchronize parent form control "txtLinkDlv2" with subform's "dlvNo"
End Sub
