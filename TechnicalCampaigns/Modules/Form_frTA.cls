VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ================================================================================
' Module Name    : frTA
' Purpose        : This is the main form for managing BMW Technical Campaigns.
'                  It supports adding new campaigns, creating spare part requests,
'                  generating client letters, and monitoring their sending.
'
' Parent Form    : None
' Child Forms    :
'   - sqTA_Vcl (Completion tab)   : Displays information about the vehicle, client, and campaign completion status.
'   - Spare Parts tab:
'       - sfParts_Veh            : List of vehicles affected by the Technical Campaign.
'       - sqParts_Parts          : List of spare parts required for the campaign on a single vehicle.
'       - sfPartSum              : Total quantity of parts required for the campaign on all vehicles.
'       - sfOrders               : Information about spare part orders.
'       - sf2OrdItem             : Contents of the spare part orders (ordered items).
'       - sfDlv                  : Information about spare part deliveries.
'       - sf2DlvItem             : Contents of the spare part deliveries (delivered items).
'   - sfStockCars (Stock tab)    : Displays vehicles in the importer’s stock.
'   - sfLetters (Letters tab)    : Manages client letters (generation, tracking dispatch, and delivery).
'
' Key Features   :
'   - Campaign tracking and management.
'   - Automated letter generation and mail merging.
'   - Spare parts orders.
'
' Notes          : This form is loaded automatically upon opening the database.
'
' Events         :
'   1. cmdLetterBlank_Click  - Handles the "Create Blank Letter" button click event.
'   2. cmdLetters_Click      - Handles the "Generate Letters" button click event.
'   3. cmdMoveToDlv_Click    - Handles the "Move to Delivery" button click event.
'   4. cmdMoveToOrd_Click    - Handles the "Move to Order" button click event.
'   5. Form_Current          - Formats the form fields and labels based on the campaign status.
'   6. cmdFrmCls_Click       - Handles the "Close" button click event to safely close the form.
'
' Methods        :
'   1. MailMerge         - Exports data to Word and generates personalized letters.
'   2. FixPath           - Ensures a file path ends with a backslash.
'
' Developer      : Oleh Bondarenko
' Created        : 2008-07-25
' Last Updated   : 2025-01-10 by Oleh Bondarenko - Added comments about Parents and Child Forms
' ================================================================================
Option Compare Database
Option Explicit
' ---------------------
' Variable Definitions:
' ---------------------
Private mstrMailTo As String                                                ' - Stores the email list for notifications.

' Constants for templates and queries used in the mail merge process:
Private Const conTemplate As String = "acbMailMerge.doc"
Private Const conQuery As String = "qsLetterSource"
' -------------------------------------------------------------------
' Event #1        : cmdLetterBlank_Click
' Purpose         : Creates a letter table for the current campaign and
'                   performs a mail merge using a predefined Word template.
' Notes           : The campaign number is retrieved from the main form.
' -------------------------------------------------------------------
Private Sub cmdLetterBlank_Click()

' Generate a table of clients for the current campaign and open the Word template:
    DoCmd.OpenQuery ("qaLtCreatDate")                                       ' - Add new records to the letters table
    Call MailMerge                                                          ' - Initiates the mail merge process.

End Sub
' -------------------------------------------------------------------
' Method #1         : MailMerge
' Purpose           : Performs a mail merge operation by exporting query data to a Word document.
'                     Generates personalized letters for clients based on campaign data.
' Error Handling    : Yes
' -------------------------------------------------------------------
Public Sub MailMerge()
    Dim strPath As String
    Dim strDataSource As String
    
    Dim doc As Word.Document
    Dim wrdApp As Word.Application
        
On Error GoTo HandleErrors
    ' Define the file path and remove any existing document:
    strPath = CurrentProject.Path & "\"
    strDataSource = strPath & conQuery & ".doc"
    Kill strDataSource
    
    ' Export the data to rtf format:
    DoCmd.OutputTo acOutputQuery, conQuery, _
     acFormatRTF, strDataSource, False

    ' Start Word and load the mail merge template:
    Set wrdApp = New Word.Application
    Set doc = wrdApp.Documents.Add(strPath & conTemplate)
        
    ' Perform the mail merge operation:
    With doc.MailMerge
        .OpenDataSource Name:=strDataSource
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        If .State = wdMainAndDataSource Then .Execute
    End With
    
    ' Make the Word application visible with the generated document:
    wrdApp.Visible = True
    
ExitHere:
    Set doc = Nothing
    Set wrdApp = Nothing
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case 53                                                 ' - File not found
            Resume Next
        Case Else
            MsgBox Err.Number & ": " & Err.Description
            Resume ExitHere
    End Select
    Resume

End Sub
' -------------------------------------------------------------------
' Method #2       : FixPath
' Purpose         : Ensures that a file path ends with a backslash ("\").
'                   This is useful for constructing consistent directory paths.
' Parameters      :
'   - strPath (String): The original file path to validate and adjust.
' Returns         :
'   - String: A corrected file path that always ends with a backslash.
' Notes           :
'   - If the input path already ends with a backslash, no changes are made.
' -------------------------------------------------------------------
Private Function FixPath(strPath As String) As String
    
    ' Check if the path ends with a backslash:
    If Right(strPath, 1) = "\" Then
        FixPath = strPath                                       ' - Path is already valid, return as is.
    Else
        FixPath = strPath & "\"                                 ' - Append a backslash if missing
    End If
    
End Function
' -------------------------------------------------------------------
' Event #2        : cmdLetters_Click
' Purpose         : Generates letter to dealers for a selected campaign using the required settings.
'                   Configures Microsoft Word mail merge and interacts with ActiveX components.
' Notes           :
'   - Requires the following references:
'       1. Microsoft ActiveX Data Objects 2.1 Library
'       2. Declare the rst variable in CreateTableFromRecordset as ADODB.Recordset.
'       3. Use ADODB.Recordset in the LetterCreate procedure after the "New" keyword
' -------------------------------------------------------------------
Private Sub cmdLetters_Click()
' Ensure proper library references are set up for ActiveX Data Objects.
' Instructions:
'   - Tools -> References -> Microsoft ActiveX Data Objects 2.1 Library
'   - in the CreateTableFromRecordset procedure declare the rst variable as ADODB.Recordset
'   - in the LetterCreate procedure, after the keyword New, specify ADODB.Recordset
    
    Call LetterCreate(Me)                                       ' - Initiate the letter creation process
    
End Sub
' -------------------------------------------------------------------
' Event #3        : cmdMoveToDlv_Click
' Purpose         : Transfers selected spare parts from the order to the delivery table.
'                   Updates the subforms for delivery items and part summaries accordingly.
' Notes           :
'   - Operates on the selected items in the "sf2OrdItem" subform of the main form "frTA".
'   - The "tblDlvItems" table is updated with the delivery data.
'   - Requires at least one item to be selected in the subform for processing.
' Error Handling  : Displays a message box if no items are selected.
' -------------------------------------------------------------------
Private Sub cmdMoveToDlv_Click()
' Transfer of allocated spare parts to the delivery
'
    Dim frm As Form                                     ' - Reference to the subform
    Dim rst As Recordset                                ' - Recordset for the selected items in the subform
    Dim rstDlv As Recordset                             ' - Recordset for the delivery table
    Dim i As Integer                                    ' - Counter for processing selected items
    
    ' Set references to the subform and its recordset:
    Set frm = Forms!frTA!sf2OrdItem.Form
    Set rst = frm.RecordsetClone
    Set rstDlv = CurrentDb.OpenRecordset("tblDlvItems", , dbAppendOnly)
    
    If frm.Tag > 0 Then                                 ' - Check if there are selected items in the subform
        rst.AbsolutePosition = frm.SelTop - 1
        For i = 1 To frm.Tag                            ' - Process each selected item and add it to the delivery table
            With rstDlv
                .AddNew                                 ' - Create a record in the deliveries table
                !diTA = Forms!frTA!aNo                  ' - Technical Campaign number
                !diNr = Forms!frTA!sfDlv.Form!lnDlvNo   ' - Delivery number
                !diPtNr = rst!oiPtNr                    ' - Part number
                !diPtQ = rst!oiPtQ                      ' - Quantity of parts
                .Update                                 ' - Save the new record
                rst.MoveNext                            ' - Move to the next selected item
            End With
        Next i
        rstDlv.Close
    Else
        ' Notify the user if no items are selected for delivery:
        MsgBox "Select items to be transferred to the order"
    End If
    
    ' Cleanup: release references and refresh subforms
    Set frm = Nothing
    Set rst = Nothing
    Set rstDlv = Nothing
    
    Me!sf2DlvItem.Form.Requery                          ' - Refresh the delivery items subform
    Me!sfPartSum.Form.Refresh                           ' - Refresh the part summary subform

End Sub
' -------------------------------------------------------------------
' Event #4          : cmdMoveToOrd_Click
' Purpose           : Transfers selected spare parts to an order and updates the order table.
'                     Generates a new order number and refreshes related subforms.
' Notes             :
'   - Operates on selected items in the "sfPartSum" subform of the main form "frTA".
'   - Updates the "tblOrdItems" table with new order items.
'   - Sends an email to the Parts Department to notify them of the new order.
' Error Handling    : Displays a message box if no items are selected for transfer.
' -------------------------------------------------------------------
Private Sub cmdMoveToOrd_Click()
' Transfer of required parts to the order
'
    Dim frm As Form                                         ' - Reference to the subform
    Dim rst As Recordset                                    ' - Recordset for the selected items in the subform
    Dim rstOrd As Recordset                                 ' - Recordset for the order table
    Dim i As Integer                                        ' - Counter for processing selected items
    
    ' Set references to the subform and its recordset:
    Set frm = Forms!frTA!sfPartSum.Form
    Set rst = frm.RecordsetClone
    
    If frm.Tag > 0 Then                                     ' - Check if there are selected items in the subform
        DoCmd.SetWarnings False
            DoCmd.OpenQuery "prgPartOrdAdd"                 ' - Create a record in the order table
        DoCmd.SetWarnings True
        '------------------------------ An order number has been created ------------------------------------
        
        Forms!frTA!sfOrders.Form.Requery                    ' - Refresh the order subform to display the new order
        
        ' Process each selected item and add it to the order table:
        Set rstOrd = CurrentDb.OpenRecordset("tblOrdItems", , dbAppendOnly)
        rst.AbsolutePosition = frm.SelTop - 1
        For i = 1 To frm.Tag
            With rstOrd
                .AddNew
                !oiTA = Forms!frTA!aNo                      ' - Technical campaign number
                !oiNr = Forms!frTA!sfOrders.Form!lnOrdNo    ' - Order number
                !oiPtNr = rst!ptNo                          ' - Part number
                !oiPtQ = rst!RqTot                          ' - Required quantity
                .Update                                     ' - Save the new record
                rst.MoveNext                                ' - Move to the next selected item
            End With
        Next i
        rstOrd.Close
    Else
        ' -Notify the user if no items are selected for transfer:
        MsgBox "Select items to be transferred to the order"
    End If
    
    ' Cleanup: release references and refresh subforms
    Set frm = Nothing
    Set rst = Nothing
    Set rstOrd = Nothing
    
    Me!sf2OrdItem.Form.Requery                              ' - Refresh the order items subform
    Me!sfPartSum.Form.Refresh                               ' - Refresh the part summary subform
    
    ' Trigger the procedure to notify the Parts Department via email:
    Call PartOrdMail

End Sub
' -------------------------------------------------------------------
' Event #5        : Form_Current
' Purpose         : Adjusts the visual formatting of the form's fields and labels
'                   based on the status of the current technical campaign.
' Notes           :
'   - Highlights the campaign as "stopped" or "active" using color and font formatting.
'   - Ensures that users can visually distinguish between active and stopped campaigns.
' Error Handling  : None
' -------------------------------------------------------------------
Private Sub Form_Current()

    ' Check the "stopped" status of the current technical campaign:
    If Me.aStop Then
        ' Format the form for a stopped campaign:
        Me.aNo.ForeColor = 8421504                                      ' - Gray color for campaign number
        Forms!frTA!àDescr.ForeColor = 8421504                           ' - Gray color for campaign description
        Me.lbRecall.ForeColor = 8421504                                 ' - Gray color for "Recall" label
        Me.lbStop.ForeColor = 255                                       ' - Red color for "Stop" label
        Me.lbStop.FontBold = True                                       ' - Bold font for the "Stop" label
        
    Else
        ' Format the form for an active campaign:
        Me.aNo.ForeColor = 16711680                                     ' - Bright red color for campaign number
        Forms!frTA!àDescr.ForeColor = 16711680                          ' - Bright red color for campaign description
        Me.lbRecall.ForeColor = 255                                     ' - Red color for "Recall" label
        Me.lbStop.ForeColor = -2147483630                               ' - Default label color for "Stop."
        Me.lbStop.FontBold = False                                      ' - Normal font for the "Stop" label
        
    End If

End Sub
' -------------------------------------------------------------------
' Event #6        : cmdFrmCls_Click
' Purpose         : Closes the current form (main form: frTA), handling any unexpected errors.
' Notes           :
'   - Designed for use as a "Close" button event handler.
'   - Ensures that the form closes without disrupting other processes.
' Error Handling  : Yes
' -------------------------------------------------------------------
Private Sub cmdFrmCls_Click()
On Error GoTo Err_cmdFrmCls_Click


    DoCmd.Close     ' Close current form (main form: frTA)

Exit_cmdFrmCls_Click:
    Exit Sub

Err_cmdFrmCls_Click:
    MsgBox Err.Description
    Resume Exit_cmdFrmCls_Click
    
End Sub
