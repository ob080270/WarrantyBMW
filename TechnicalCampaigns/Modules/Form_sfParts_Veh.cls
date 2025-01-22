VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfParts_Veh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module Name     : sfParts_Veh
' Parent Form       : frTA ("Spare parts" tab)
' Child Forms       : None
Option Compare Database
Option Explicit
' -------------------------------------------------------------------
' Subroutine Name : Form_Current
' Purpose         : Synchronizes parent form fields with subform fields
'                   during navigation.
'                   - Updates txtLink1 on the parent form with the
'                     current avActNo value from the subform.
'                   - Updates txtLink2 on the parent form with the
'                     current avVIN value from the subform.
' Error Handling  : Ignores errors using 'On Error Resume Next' to
'                   prevent interruption due to potential null values.
' -------------------------------------------------------------------
Private Sub Form_Current()
On Error Resume Next                        ' - Ignore runtime errors
    Me.Parent("txtLink1") = [avActNo]       ' - Synchronize parent form control "txtLink1" with subform's "avActNo"
    Me.Parent("txtLink2") = [avVIN]         ' - Synchronize parent form control "txtLink2" with subform's "avVIN"
End Sub

