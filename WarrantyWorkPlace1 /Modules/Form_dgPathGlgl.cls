VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dgPathGlgl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ========================================================================================
' Form Module   : dgPathGlgl
' Purpose       : Provides a convenient way to access the generated batch file "glglgl20"
'                 without navigating through file explorer.
' Trigger       : This form is displayed after the batch file creation process is completed.
'                 The batch file creation is initiated by the "cmdGLGL20" button on the
'                 "frNewClmSend" form.
'
' Parent Form   : frNewClmSend
' Child Forms   : None
'
' Features      :
'   - Displays the message confirming the successful creation of the batch file.
'   - Includes a hyperlink to directly open the file's location.
'
' Notes         : This form is designed to finalize the file generation process and
'               improve user experience by reducing navigation effort.
'
' Developer     : Oleh Bondarenko
' Created       : 2007-01-17
' Last Updated  : 2025-01-09 by Oleh Bondarenko - Added comments for GitHub upload
' ========================================================================================

Option Compare Database
Option Explicit

Private Sub Form_Load()

    If IsNull(Me.OpenArgs) Then Exit Sub
    
    Me.lbHl.Caption = Me.OpenArgs                       ' - Set the label caption to display the provided file path or folder
    Me.lbHl.HyperlinkAddress = Me.OpenArgs              ' - Set the label hyperlink to allow direct access to the file or folder
    
End Sub

Private Sub lbHl_Click()

    DoCmd.Close acDefault, , acSaveNo                   ' - Close the current form after hyperlink interaction
    
End Sub
