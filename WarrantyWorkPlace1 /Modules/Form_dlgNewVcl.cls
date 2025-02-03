VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgNewVcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name  : dlgNewVcl
' Description  : Dialog form for entering or updating vehicle data.
'                Manages data saving to the database and highlights necessary fields for input.
' Parent Form  : frClaims
' Child Form   : None
'
' Key Features:
' 1. Saves user and timestamp information into tb$UserLog.
' 2. Validates and highlights missing data fields.
' 3. Supports various operational modes, including vehicle creation and data update.
' 4. Handles integration with VIN clipboard operations.
'
' Events:
' 1. cmdAprove_Click    : Saves the data and closes the form.
' 2. Form_Load          : Initializes the form based on the passed arguments (OpenArgs).
' 3. lstSeller_NotInList: Handles cases where a seller is not in the list.
' 4. luBSI_AfterUpdate  : Toggles visibility and layout of BSI fields.
' 5. luEng_AfterUpdate  : Validates and formats the engine field.
' 6. luModel_AfterUpdate: Validates and formats the model field.
' 7. luSerial_AfterUpdate: Validates and formats the serial field.
' 8. luSerial_NotInList : Handles cases where a serial is not in the list.
'
' Parent Form Calls:
' 1. Triggered by frClaims.Current when VIN.Tag <> "NewVehicle".
' 2. Triggered by sf1Defects.Enter when VIN.Tag = "NewVehicle".
' 3. Triggered by warrImportTopFromExcell in GlobalFn when vehicle data is missing.
' 4. BSI claims handling also routes to dlgNewVcl for data adjustments.
'
' Developer    : Oleh Bondarenko
' Created      : 2004-11-25
' Last Updated : 2025-01-12 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Compare Database
Option Explicit
' Declare a clipboard object at the form level:
Private Clipboard As New CClipboard
' -------------------------------------------------------------------
' Event #1: cmdAprove_Click
' Purpose : Confirms the entered data and closes the form.
'           Saves user and VIN data into tb$UserLog based on the form's mode.
' -------------------------------------------------------------------
Private Sub cmdAprove_Click()
' This event handles pressing the button "V" (OK):
' - Save the entered data and close the form
' - Save who and when worked with the form in the table tb$UserLog, as well as with which VIN was edited
'
' - If the dialog form of saving car data is opened for entering a new car - then the VIN of this car is saved in the Me.Tag,
' which is further recorded in the ccMode field of the table “tb$UserLog”.
'
' If the dialog form is opened in the data replenishment mode - then “NF” is written to Me.Tag,
' which is further saved in the ccMode field of the table “tb$UserLog”.
'
' So, to determine who first entered the data of a particular car, in the table “tb$UserLog”
' we look for User in combination with VIN.
'
    Dim rst As Recordset
    
    ' Open tb$UserLog table to record user activity:
    Set rst = CurrentDb.OpenRecordset("tb$UserLog")
    With rst
    ' Save user
    ' OpenArgs = Me.Tag = "NF"
        If Me.Tag = "NF" Or Me.Tag = "BSI" Then         ' - Mode of supplementing with missing data
            .AddNew
            !ccWhen = Now
            !ccUser = CurrentUser
            !ccVIN = Me!VIN
            !ccMode = "NF"
            .Update
            .Close
        ElseIf Me.Tag <> "WS" Then                      ' - New data entry mode
            .AddNew
            !ccWhen = Now
            !ccUser = CurrentUser
            !ccVIN = Me!VIN
            !ccMode = Me.Tag
            .Update
            .Close
        End If
    End With

    DoCmd.Close acForm, "dlgNewVcl"
    
End Sub
' -------------------------------------------------------------------
' Event #2: Form_Load
' Purpose : Initializes the form's fields based on passed arguments.
'           Highlights missing fields and sets focus accordingly.
' -------------------------------------------------------------------
Private Sub Form_Load()
    
    Dim btFocuse As Byte
    btFocuse = 0                                        ' - Initial focus on VIN field
    
    With Me
        If IsNull(.OpenArgs) Then Exit Sub              ' - Exit if no arguments are provided
        
        ' If the form opens with the NF argument,
        ' it means that not all fields are filled,
        ' and the form is opened for editing,
        If .OpenArgs = "NF" Then                        ' - Form opens for completion
            .Caption = "Enter the missing data:"
            
            ' with highlighting of necessary fields:
            If IsNull(.luSerial) Then
                .luSerial.BorderColor = BRIGHT_RED
                btFocuse = 1
            End If
            If IsNull(.luModel) Then
                .luModel.BorderColor = BRIGHT_RED
                If btFocuse = 0 Then
                    btFocuse = 2
                End If
            End If
            
            If Not isMoto(.luVIN) Then                  ' - It is not necessary to illuminate the engine field for motorcycles
                If IsNull(.luEng) Then
                    .luEng.BorderColor = BRIGHT_RED
                    If btFocuse = 0 Then
                        btFocuse = 3
                    End If
                End If
            Else
                !luMoto = True                          ' - Mark as motorcycle
            End If
            
            Select Case btFocuse
                Case 1: !luSerial.SetFocus
                Case 2: !luModel.SetFocus
                Case 3: !luEng.SetFocus
            End Select
        
        ElseIf .OpenArgs = "WS" Then                    ' - If the form is opened to adjust the date of sale
        
            .Caption = "Edit warranty start date:"
            .StartWarranty.BorderColor = BRIGHT_RED
            .StartWarranty.SetFocus
                    
        ElseIf .OpenArgs = "BSI" Then                    ' - If the form is opened to adjust BSI data
            .Caption = "Enter BSI data:"
            !luBSI = True
            
            ' Display BSI fields:
            !luRem.Height = 593
            !luRem.Top = 1593
            !luBSIend.Visible = True
            !luBSIkm.Visible = True
            !luBSIlt.Visible = True
            
            ' Highlight fields with a red frame:
            .luBSIo.BorderColor = BRIGHT_RED
            .luBSIend.BorderColor = BRIGHT_RED
            .luBSIkm.BorderColor = BRIGHT_RED
            .luBSIlt.BorderColor = BRIGHT_RED
            .luBSIo.SetFocus
            
        Else
            !VIN = .OpenArgs
            !StartWarranty = Forms!frClaims!Wstart
            If isMoto(.luVIN) Then                      ' - Mark as motorcycle
                .luMoto = True
            End If
            !VIN.SetFocus
        End If
        
        Me.Tag = .OpenArgs                              ' - Save OpenArgs in the form tag for further use in buttons
        
        If Not IsNull(!VIN) Then
            Clipboard.SetText !VIN                      ' - Save VIN to clipboard
        End If
        
    End With

End Sub
' -------------------------------------------------------------------
' Event #3: lstSeller_NotInList
' Purpose : Handles cases where a seller is not in the list.
' Behavior: Prompts the user to create a new seller record and opens the corresponding form if confirmed.
' -------------------------------------------------------------------
Private Sub lstSeller_NotInList(NewData As String, Response As Integer)
    Dim intNew As Integer
    
    intNew = MsgBox("The entered name is not in the list. Create a new record?", vbYesNo)
    If intNew = vbYes Then
        Response = acDataErrAdded
        DoCmd.OpenForm "dlg2NewSeller", DataMode:=acFormAdd, WindowMode:=acDialog, _
            OpenArgs:=NewData
    Else
        Me!lstSeller.Undo
    End If

End Sub
' -------------------------------------------------------------------
' Event #4: luBSI_AfterUpdate
' Purpose : Adjusts the layout and visibility of BSI-related fields
'           based on the state of the BSI checkbox.
' Behavior:
'   - If BSI is checked:
'       - Make BSI fields visible.
'       - Adjust the size and position of the remarks field (luRem).
'   - If BSI is unchecked:
'       - Hide BSI fields.
'       - Restore the size and position of the remarks field (luRem).
' -------------------------------------------------------------------
Private Sub luBSI_AfterUpdate()

    With Me
        If !luBSI Then
        ' BSI is enabled: Adjust layout to show BSI fields:
            !luRem.Height = 593                                 ' - Reduce remarks field height
            !luRem.Top = 1593                                   ' - Move remarks field down
            !luBSIend.Visible = True                            ' - Show BSI end date field
            !luBSIkm.Visible = True                             ' - Show BSI kilometers limit field
            !luBSIlt.Visible = True                             ' - Show BSI oil limit field
        Else
        ' BSI is disabled: Restore original layout:
            !luRem.Height = 1613                                ' - Restore remarks field height
            !luRem.Top = 573                                    ' - Move remarks field up
            !luBSIend.Visible = False                           ' - Hide BSI end date field
            !luBSIkm.Visible = False                            ' - Hide BSI kilometers limit field
            !luBSIlt.Visible = False                            ' - Hide BSI oil limit field
        End If
    End With

End Sub
' -------------------------------------------------------------------
' Event #5: luEng_AfterUpdate
' Purpose : Validates and formats the Engine field after user input.
' Behavior:
'   - If the field is empty, retains the red border as an error indicator.
'   - If the field contains a value, converts it to uppercase and changes the border color to white.
' -------------------------------------------------------------------
Private Sub luEng_AfterUpdate()
        
    With Me
        If Len(Nz(.luEng)) > 0 Then
            .luEng.BorderColor = WHITE          ' - Change border to white for valid input
            .luEng = UCase$(.luEng)             ' - Convert value to uppercase
        End If
    End With

End Sub
' -------------------------------------------------------------------
' Event #6: luModel_AfterUpdate
' Purpose : Validates and formats the Model field after user input.
' Behavior:
'   - If the field is empty, retains the red border as an error indicator.
'   - If the field contains a value, converts it to uppercase and changes the border color to white.
' -------------------------------------------------------------------
Private Sub luModel_AfterUpdate()
        
    With Me
        If Len(Nz(.luModel)) > 0 Then
            .luModel.BorderColor = WHITE        ' - Change border to white for valid input
            .luModel = UCase$(.luModel)         ' - Convert value to uppercase
        End If
    End With

End Sub
' -------------------------------------------------------------------
' Event #7: luSerial_AfterUpdate
' Purpose : Validates and formats the Serial field after user input.
' Behavior:
'   - If the field is empty, retains the red border as an error indicator.
'   - If the field contains a value, converts it to uppercase and changes the border color to white.
' -------------------------------------------------------------------
Private Sub luSerial_AfterUpdate()
        
    With Me
        If Len(Nz(.luSerial)) > 0 Then
            .luSerial.BorderColor = WHITE       ' - Change border to white for valid input
            .luSerial = UCase$(.luSerial)       ' - Convert value to uppercase
        End If
    End With

End Sub
' -------------------------------------------------------------------
' Event #8: luSerial_NotInList
' Purpose : Handles the case when a serial number entered by the user
'           is not found in the list. Provides an option to add the
'           new serial number to the database.
' Behavior:
'   - Displays a prompt to the user asking if they want to add the new serial.
'   - If the user confirms:
'       - Opens the "tblSeries" table.
'       - Adds the new serial number in uppercase along with the current user and timestamp.
'       - Updates the recordset and sets the response to indicate that data has been added.
'   - If the user declines:
'       - Displays the standard error message and undoes the user input.
' -------------------------------------------------------------------
Private Sub luSerial_NotInList(NewData As String, Response As Integer)
    Dim strMsg As String
    Dim rst As Recordset
    
    ' Prepare the message for the user:
    strMsg = "'" & NewData & "' is not in the list." & Chr$(13) & Chr$(10)
    strMsg = strMsg & "Would you like to add it?"
    
    If vbYes = MsgBox(strMsg, vbYesNo + vbQuestion, "New Serial") Then
        ' User chooses to add the new serial:
        Set rst = CurrentDb.OpenRecordset("tblSeries")  ' - Open the "tblSeries" table
        With rst
            .AddNew
            !luSerial = UCase$(NewData)                 ' - Add the new serial in uppercase
            !luUser = CurrentUser                       ' - Log the current user
            !luDtEntr = Now                             ' - Log the current timestamp
            .Update
            .Close
            Response = acDataErrAdded                   ' - Indicate that data has been added
        End With
    Else
        ' User chooses not to add the new serial:
        Response = acDataErrDisplay                     ' - Display standard error message
    End If
    
End Sub

