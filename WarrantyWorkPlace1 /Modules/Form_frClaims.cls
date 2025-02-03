VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' Module Name  : frClaims
' Description  : Main form for managing warranty claims. The form is used by the
'                lead warranty specialist at the Importer and serves as the starting
'                point of the application after the database is loaded.
'
' Parent Form  : None
' Child Form   : sfDefects, sf2Parts, sf1BMWNo
'
' Key Features :
' 1. Manages and processes warranty claims, including filtering, auditing, and corrections.
' 2. Allows interaction with related data such as parts, defects, and audit information.
' 3. Provides integration with BMW systems for claim submissions and processing.
' 4. Includes functionality for generating and sending warranty reports and managing claim lifecycle.
'
' Events:
' 1. btNClosedFlt_Click         : Applies a filter for unclosed claims.
' 2. btRjPosRs_Click            : Copies rejected positions into a new claim.
' 3. clActive_AfterUpdate       : Marks the claim as ready-to-send to the manufacturer,
'                                 and updates the open date when the active flag is toggled.
' 4. clClosed_AfterUpdate       : Marks the claim as closed and also performs some actions upon claim closure.
' 5. clDealerID_AfterUpdate     : Adjusts dealer ID mappings based on new system data.
' 6. clWType_AfterUpdate         : Adjusts related fields when warranty type changes.
' 7. cmdActBSI_Click             : Generates and previews BSI reports.
' 8. cmdClcNotSendedCount_Click  : Calculates the count of ready-to-send claims.
' 9. cmdCreditImport_Click       : Imports credit data and updates related claims.
' 10. cmdFilter2_Click           : Applies a filter for claims under correction.
' 11. cmdIIW_Click               : Imports warranty data from Impulse dealer (.csv).
' 12. cmdNewSendOpen_Click       : Opens the form for viewing new claim sends.
' 13. cmdRjResendOpen_Click      : Opens the form for rejected claim resends.
' 14. dtRepair_AfterUpdate       : Warns if the repair date is too old.
' 15. Form_Current               : Handles record navigation, refreshing data views.
' 16. Form_Open                  : Initializes form controls and data on load.
' 17. lstSendNr_AfterUpdate      : Updates the subtable for BMW sends.
' 18. objCredit_CreditImported   : Handles the event when credit data is imported.
' 19. OrdNo_AfterUpdate          : Checks the uniqueness of the order number.
' 20. sfDefects_Enter            : Ensures proper vehicle data entry upon defect form entry.
' 21. swSubformEdit_AfterUpdate  : Toggles edit mode for subform sf2Parts fields.
' 22. txtRemBMW_AfterUpdate      : Removes invalid characters from the BMW comments field.
' 23. txtRemBMW_Change           : Counts the characters entered in the BMW comments field.
' 24. txtRemBMW_Undo             : Restores the BMW comments field to its previous value.
' 25. VIN_AfterUpdate            : Checks vehicle data and adjusts related fields.
' 26. cbDlrFilter_AfterUpdate    : Filters records based on the selected dealer.
' 27. swView_AfterUpdate         : Toggles filters for work and parts subforms (sf2Parts).
' 28. swfltrDA_AfterUpdate       : Toggles the filter for deactivated claims.
' 29. cmdCredRep_Click           : Generates the BMW warranty report (last Credit Import).
' 30. cmdDlrAct_Click            : Generates the Dealer Warranty Reimbursement Invoice.
' 31. cmdWarBlankImport_Click    : Imports data from the warranty blank Excel file.
'
' Properties:
' 1. ptDA                       : Tracks the number of deactivated claims.
' 2. isExistVIN                 : Checks if the vehicle is already in the database.
'
' Methods:
' 1. getEUR                     : Downloads and saves current exchange rates.
' 2. clcCR                      : Calculates the currency rate from XML data.
'
' Dependencies:
' 1. BSI_Class          - Used for managing actions and states related to BSI (BMW Service Inclusive).
' 2. DivQ_Class         - Resolves discrepancies in quantities of AW (After Work) and spare parts.
' 3. NewCredit_Class    - Processes credit-related claims data, including importing
'                         and distributing credit note information.
' 4. CClipboard_Class   - Provides clipboard functionality for managing copied claim data.
' 5. GlobalFn.bas       - Supplies global functions for claim validation, exchange rate
'                         management, and Outlook task integration.
'
' Developer    : Oleh Bondarenko
' Created      : 2014-12-18
' Last Updated : 2025-01-20 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================
Option Compare Database
Option Explicit

' MODULE-LEVEL VARIABLE DECLARATUIONS:
' (These variables are used throughout the module for managing state
' and implementing the business logic of the frClaims form)
Private intptDA As Integer                                              ' - Tracks the number of claims removed from active status and not yet closed
Private strClaimNrDefault As String                                     ' - Stores the default format for generating new claim numbers

' Event sinks for managing file operations and synchronization
' Objects below provide functionality for processing warranty-related files
Private objDivQ As DivQ                                                 ' - Handles operations related to discrepancy queries for claim items
Private objWarrAct As WarrAct                                           ' - Handles warranty-related actions, including creating reports and managing status
Private objBSIAct As BSI                                                ' - Handles actions and calculations related to BSI (BMW Service Inclusive) processes
Private objSend As NewSend                                              ' - Manages the sending of warranty claims and associated operations

' Handles credit file imports, including event handling for data processing and validation:
Private WithEvents objCredit As NewCredit
Attribute objCredit.VB_VarHelpID = -1

' -------------------------------------------------------------------
' Event #1        : btNClosedFlt_Click
' Purpose         : Applies a filter to show only claims that are not closed.
' Behavior        : Executes when the corresponding button is clicked. Filters
'                   the records in the form to display only claims where the
'                   "clClosed" field is False.
' -------------------------------------------------------------------
Private Sub btNClosedFlt_Click()

    DoCmd.ApplyFilter , "clClosed = false"                              ' - Apply filter to show only unclosed claims.
    
End Sub

' -------------------------------------------------------------------
' Event #2        : btRjPosRs_Click
' Purpose         : Copies rejected warranty claim positions into a new claim.
' Behavior        : Triggered when the corresponding button is clicked. It calculates
'                   the rejected positions, prompts the user to confirm creating a new
'                   warranty claim, and then copies relevant data from the current
'                   claim into the new one.
' -------------------------------------------------------------------
Private Sub btRjPosRs_Click()
    ' Define variables for copying data into a new record:
    Dim CrLf As String * 2      ' - Line break
    Dim lnClNo As Long          ' - Claim number
    Dim btDlrID As Byte         ' - Dealer ID
    Dim lnOrdNr As Long         ' - Order number
    Dim lnZayav As Long         ' - Request number
    Dim strVIN As String        ' - Vehicle Identification Number
    Dim dtStartWar As Date      ' - Warranty start date
    Dim dtRpr As Date           ' - Repair date
    Dim lnMileage As Long       ' - Mileage
    Dim btWarTp As Byte         ' - Warranty type (1-vehicle / 2-parts)
    Dim blCustTp As Boolean     ' - Transit customer flag
    Dim strAutor As String      ' - Authorization code
    Dim strCmnt As String       ' - Comments
    Dim strCmtBMW As String     ' - BMW comments
    Dim strCustCompl As String  ' - Customer complaint
    Dim blAud As Boolean        ' - Audit flag
    Dim clcCount As Byte        ' - Count of rejected positions
    Dim strCriteria As String   ' - Criteria for selecting rejected records
    Dim strMsg As String        ' - User prompt message
    Dim intUsrChoose As Integer ' - User's choice from the prompt
    Dim lnOldClm As Long        ' - Old claim number
    Dim rst As Recordset        ' - Recordset for data manipulation
    
    CrLf = Chr(13) & Chr(10)    ' - Line break
    
    ' Count rejected positions for the current claim:
    strCriteria = "[ptClaimeNo] =" & Forms!frClaims!clNo & " AND [ptCompensBMW] = 0"
    clcCount = DCount("[ptNo]", "tblParts", strCriteria)
    
    ' Prompt the user to confirm creating a new warranty claim:
    strMsg = "The current claim has " & str(clcCount) & " rejected positions. " & _
            CrLf & "Do you want to create a new Warranty Claim?"
    intUsrChoose = MsgBox(strMsg, vbYesNo + vbInformation + vbDefaultButton2, "Reject pos. resubmission")
    
    If intUsrChoose = vbYes Then
        ' Prepare data for the new claim:
        lnClNo = DMax("[clNo]", "tblClaimes", "not fnIsParity([clNo])") + DLookup("[opClaimNr]", "tblOptions")
        lnOldClm = Me!clNo
        btDlrID = Me!clDealerID
        lnOrdNr = Me!clOrdNo
        lnZayav = Me!clZayav
        strVIN = Me!clVIN
        dtStartWar = Me!Wstart
        dtRpr = Me!dtRepair
        lnMileage = Me!clMileage
        btWarTp = Me!clWType
        blCustTp = Me!clCustType
        strAutor = IIf(IsNull(Me!clAutoriz), "", Me!clAutoriz)
        strCmnt = IIf(IsNull(Me!clComment), "см. " & lnOldClm, _
                "см. " & lnOldClm & CrLf & Me!clComment)
        strCmtBMW = IIf(IsNull(Me!clCommentBMW), "", Me!clCommentBMW)
        strCustCompl = IIf(IsNull(Me!clCustCompl), "", Me!clCustCompl)
        Me!clComment = " Rj перенесено в новый Claim " & _
            CStr(DMax("[clNo]", "tblClaimes", "not fnIsParity([clNo])") + DLookup("[opClaimNr]", "tblOptions")) _
            & CrLf & Me!clComment
        blAud = Me!clAudit
        Me!clClosed = True
        
        ' Copy data to the new claim:
        Set rst = Me.RecordsetClone
        With rst
            .AddNew
            !clNo = lnClNo
            !clActive = True
            !clOpenDate = Date
            !clManager = CurrentUser()
            !clDealerID = btDlrID
            !clOrdNo = lnOrdNr
            !clZayav = lnZayav
            !clVIN = strVIN
            !clWType = btWarTp
            !clWstart = dtStartWar
            !clRepair = dtRpr
            !clMileage = lnMileage
            !clCustType = blCustTp
            !clAutoriz = IIf(strAutor = "", Null, strAutor)
            !clComment = IIf(strCmnt = "", Null, strCmnt)
            !clCommentBMW = strCmtBMW
            !clCustCompl = strCustCompl
            !clAudit = blAud
            .Update
        End With
        ' Close the old claim.
        Me!clNo.SetFocus
        DoCmd.FindRecord lnOldClm, , , , , acCurrent
        DoCmd.SetWarnings False
            DoCmd.OpenQuery "qaRPDdefectAdd"    ' <- copying defect codes to a new claim
            DoCmd.OpenQuery "qaRPDpartAdd"      ' <- copying of work, spare parts and materials into a new claim
            DoCmd.OpenQuery "qaRPDRjCostCh"     ' <- Set prices of Rj-items to 0
        DoCmd.SetWarnings True
        rst.Close
        
        ' Mark the new claim for sending to BMW:
        Set rst = CurrentDb.OpenRecordset("tblBMWNo", , dbAppendOnly)
        With rst
            .AddNew
            !sndClaimNo = lnClNo
            !sndOldBMWNo = lnOldClm
            !sndBMWNo = !sndClaimNo
            .Update
        End With
        rst.Close
        
        ' Notify the user of the actions performed:
        MsgBox "- Verify the price reset in the previous claim." & CrLf _
             & "- Check the comments."
        
        ' Apply filter to show related claims:
        DoCmd.ApplyFilter , "clOrdNo =" & Me!clOrdNo
        Me.Requery
    End If
    
End Sub

' -------------------------------------------------------------------
' Event #3        : clActive_AfterUpdate
' Purpose         : Updates the "Open Date" field when the "Active" flag
'                   is toggled to True.
' Behavior        : Ensures that the "clOpenDate" field reflects the current
'                   date whenever the "Active" flag is marked as True.
'                   This is necessary to correctly generate 'Ready-To-Send' mailings.
' -------------------------------------------------------------------
Private Sub clActive_AfterUpdate()

    If Me.clActive Then                                                 ' - Check if the "Active" flag is True
        Me.clOpenDate = Now()                                           ' - Set the "Open Date" to the current date
    End If

End Sub

' -------------------------------------------------------------------
' Event #4        : clClosed_AfterUpdate
' Purpose         : Handles the actions needed when the "Closed" flag
'                   for a warranty claim is toggled.
' Behavior        : When the flag is set to True, the event ensures that
'                   all associated flags and data are updated accordingly.
'                   The user is prompted to confirm closing the claim.
' -------------------------------------------------------------------
Private Sub clClosed_AfterUpdate()
    ' Variables for user prompt and decision:
    Dim strMsg As String
    Dim intUserChoose As String
    
    If Me!clClosed Then                                                                                 ' - Check if the "Closed" flag is set to True
        strMsg = "Do you want to close this warranty claim and stop further processing?"                ' - Prompt the user to confirm closing the claim
        intUserChoose = MsgBox(strMsg, vbYesNo + vbExclamation + vbDefaultButton2, "Warranty closing")
    
        If intUserChoose = vbYes Then
        
            'Proceed with closing the claim:
            DoCmd.SetWarnings False
                DoCmd.OpenQuery "qaWrClExec"        ' - Execute query to handle claim closure
            DoCmd.SetWarnings True
            
            'Update fields related to claim closure:
            Me!clCorr = False                       ' - Reset the correction flag
            Me!dtCls = Date                         ' - Set the closure date
            Me!Closer = CurrentUser()               ' - Record the user who closed the claim
        Else
            'Revert the "Closed" flag if the user cancels:
            Me!clClosed = False
            'Me!dtCls = Date
            'Me!Closer = CurrentUser()
        End If
        
        Forms!frClaims.Refresh                      ' - Refresh the form to reflect changes
                
    End If

End Sub

' -------------------------------------------------------------------
' Event #5        : clDealerID_AfterUpdate
' Purpose         : Automatically adjusts dealer IDs based on predefined mappings.
' Behavior        : If the selected dealer ID is 3, it is replaced with 20 to
'                   reflect updated dealer information in the system.
'                   This is due to the new dealer number in the system.
' -------------------------------------------------------------------
Private Sub clDealerID_AfterUpdate()

    If Me!clDealerID = 3 Then
        Me!clDealerID = 20
    End If

End Sub

' -------------------------------------------------------------------
' Event #6        : clWType_AfterUpdate
' Purpose         : Adjusts related fields when the warranty type changes.
' Behavior        : If the warranty type is set to 2, the "Transit Customer"
'                   flag is automatically reset to False.
' -------------------------------------------------------------------
Private Sub clWType_AfterUpdate()

    If Me.clWType = 2 Then
        Me.clCustType = False
    End If

End Sub

' -------------------------------------------------------------------
' Event #7        : cmdActBSI_Click
' Purpose         : Generates and opens the BSI (BMW Service Inclusive) claims report.
' Behavior        : Ensures that the current date is updated in the
'                   auxiliary table for proper DTPicker functionality,
'                   then opens the BSI report in preview mode.
' External Calls  :
'                   - fnCmdSndNewSw (Module: GlobalFn)
'                     Counts the number of claims ready for sending.
' -------------------------------------------------------------------
Private Sub cmdActBSI_Click()
    Dim stDocName As String                         ' - Name of the report to open
    Dim stLinkCriteria As String
    
On Error GoTo Err_cmdDlrBSI

'   Update the auxiliary table with the current date
'   for DTPicker functionality:
    If DLookup("[auxDt]", "tb$lAux") <> Date Then
        DoCmd.SetWarnings False
            DoCmd.OpenQuery "qa$DTPicker"           ' - Updates the date to current in the auxiliary table (tb$lAux)
        DoCmd.SetWarnings True
    End If
        
'   Open the BSI report in preview mode:
    stDocName = "rpDlrActBSI"
    DoCmd.OpenReport stDocName, acPreview
    
Exit_cmdDlrBSI:
    'DoCmd.Hourglass False
    Exit Sub

Err_cmdDlrBSI:
'   Handle the case where the report opening is canceled (error 2501):
    If Err.Number = 2501 Then
        Resume Next
    Else
'   Display other errors:
        MsgBox Err.Description
        Resume Exit_cmdDlrBSI
    End If
    
End Sub

' -------------------------------------------------------------------
' Event #8        : cmdClcNotSendedCount_Click
' Purpose         : Calculates and displays the number of claims ready for sending
'                   and prompts the user to confirm sending the claims.
' Behavior        :
'                   1. Calls the external function fnCmdSndNewSw to calculate the
'                      number of claims ready for sending.
'                   2. Displays a message box with the count and asks the user
'                      to confirm sending.
'                   3. If the user agrees, generates the sending list and opens
'                      the form displaying it.
' External Calls  :
'                   - fnCmdSndNewSw (Module: GlobalFn)
'                     Counts the number of claims ready for sending.
' Error Handling  : None.
' -------------------------------------------------------------------
Private Sub cmdClcNotSendedCount_Click()
' This procedure is activated by pressing the cmdClcNotSendedCount button
' and counts the number of claims registered and ready to be sent.

    Dim clcCount As Integer         ' - Number of claims ready for sending
    Dim strMsg As String            ' - Message box text
    Dim intUserChoose As Integer    ' - User's response to the message box
        
    clcCount = fnCmdSndNewSw()      ' - Calculate the number of claims ready for sending

'   Prepare the message box text:
    strMsg = " оличество претензий, готовых к отправке: " & str$(clcCount) & _
            Chr(10) & Chr(13) & "ќтправить ?"

'   Display the message box and get the user's choice:
    intUserChoose = MsgBox(strMsg, vbYesNo + vbInformation + vbDefaultButton2, "ќтправка претензий")
    
'   If the user selects "Yes":
    If intUserChoose = vbYes Then
        DoCmd.SetWarnings False             ' - Suppress warnings during query execution
            DoCmd.OpenQuery "qNewClmSend"   ' - Execute the query to generate the sending list
        DoCmd.SetWarnings True              ' - Restore warnings
        DoCmd.OpenForm "frNewClmSend"       ' - Open the form displaying the sending list
    End If

End Sub

' ------------------------------------------------------------------------------------------------------------------------
' Event #9        : cmdCreditImport_Click
' Purpose         : Imports credit data from a predefined CSV file
'                   ("L:\Warranty\CreditImport\CFI.csv" by default)
'                   into the "tblImport" table and processes it.
' Behavior        :
'                   1. Imports data from the default file path or
'                      prompts the user if the file has already been processed.
'                   2. Updates records in related tables and handles discrepancies.
'                   3. Displays the status of the import and enables relevant
'                      UI controls based on the result.
' External Calls  :
'                   - mtCreditImport (Method of Class: NewCredit)       : Imports credit data from the CSV file.
'                   - ExistInHistory (Property of Class: NewCredit)     : Checks if the current file has already been imported.
'                   - mtAddNewCred (Method of Class: NewCredit)         : Distributes Credit Note information across the database tables.
'                   - objDivQ.RecQ (Property of Class: DivQ)            : Checks for quantity discrepancies.
'                   - objBSIAct.mtBtnStatBSI (Method of Class: BSI)     : Updates the status of the BSI button.
'                   - ExchRateSave (Global Function in GlobalFn.bas)    : Saves the current exchange rate to the CI table.
'                   - fnBtnNewSndSw (Global Function in GlobalFn.bas)   : Manages button visibility based on certain conditions.
'                   - ApprPartsTask (Global Function in GlobalFn.bas)   : Creates an Outlook task for reporting approved parts.
'
' Error Handling  :
'                   - Displays a message if the import fails.
'                   - Handles cases where the file has already been processed.
' ------------------------------------------------------------------------------------------------------------------------
Private Sub cmdCreditImport_Click()
' This procedure imports data from the L:\Warranty\CreditImport\CFI.csv file (default) into the table УtblImportФ

    Dim X As Integer                                ' - MsgBox and fnBtnNewSndSw() response
    
'   Messages to display the import status:
    Dim strOkImportMsg As String
    Dim strOkImportMsg1 As String
    Dim strOkImportMsg2 As String
    Dim strOkImportMsg3 As String
    Dim strOkImportMsg4 As String
    Dim strOkImportMsg5 As String
    Dim strOkImportMsg6 As String
'   Carriage return and line feed:
    Dim CrLf As String
    CrLf = Chr(10) & Chr(13)
    
    Dim lnUpdtRecCount As Long                      ' - Number of updated records (sum of work and parts tables)
                                                    
    Set objCredit = New NewCredit                   ' - Initialize the credit object

'   Import the credit data:
    With objCredit
        If Not .mtCreditImport(.pptPath & "CFI.csv") Then           ' - If mtCreditImport() fails
            MsgBox "File not imported."
            Exit Sub
        End If
    
        If .ExistInHistory(.pptCredNo) Then                         ' - the file is already in the history
            MsgBox "This CFI file is already in the history."
            Exit Sub
        Else
            DoCmd.SetWarnings False
                DoCmd.OpenQuery "qAddToHistory"                     ' - add imported data to the history (tblImportHistory)
            DoCmd.SetWarnings True
        End If
        
        lnUpdtRecCount = .mtAddNewCred(.pptCredNo, .pptCredDate)    ' - distribute Credit Note information across tables
        
'       Enable "Reject Resend" button if there are rejected claims:
        If .pptCountRjCl > 0 Then
            Forms!frClaims!cmdRjResend.Enabled = True
        End If
        
'       Handle quantity discrepancies
'       and update the BSI button status:
        Set objDivQ = New DivQ                  ' - discrepancy in quantity (sent/received)
        Set objBSIAct = New BSI                 ' - BSI-related claims
        
        If objDivQ.RecQ > 0 Then                ' - If the discrepancy in quantity is not resolved
            Me!cmdDivQ.Enabled = True           ' - Enable the УResolve quantity discrepancyФ button
            Me!cmdDlrAct.Enabled = False        ' - Disable the button for Warranty invoices
            Me!cmdActBSI.Enabled = False        ' - Disable the button for BSI invoices
        Else                                    ' ----------------- else (if settled) -----------------
            Me!cmdDivQ.Enabled = False          ' - Disable the УResolve quantity discrepancyФ button
            Me!cmdDlrAct.Enabled = True         ' - Enable the button for Warranty invoices
            objBSIAct.mtBtnStatBSI              ' - Enable the button for BSI invoices
        End If
       
'       Prepare the status message:
        strOkImportMsg1 = "Credit Number:" & String$(17, " ") & .pptCredNo & CrLf
        strOkImportMsg2 = "Approval Date:" & String$(15, " ") & .pptCredDate & CrLf
        strOkImportMsg3 = "Imported Records: " & String$(10 - Len(str$(.pptImpRecCount)), " ") & str$(.pptImpRecCount) & CrLf
        strOkImportMsg4 = "Updated Positions: " & String$(20 - Len(lnUpdtRecCount), " ") & str$(lnUpdtRecCount) & CrLf
        strOkImportMsg5 = "Rejected Claims: " & String$(20 - Len(str$(.pptCountRjCl)), " ") & .pptCountRjCl & CrLf
        strOkImportMsg6 = "Rejected Items: " & String$(22 - Len(str$(.pptCountRjPos)), " ") & .pptCountRjPos & CrLf
        
    End With
        
'   Display the status message:
    strOkImportMsg = strOkImportMsg1 & strOkImportMsg2 & strOkImportMsg3 & strOkImportMsg4 & _
                     strOkImportMsg5 & strOkImportMsg6
    X = MsgBox(strOkImportMsg, vbOKOnly + vbInformation, "Last Credit Status")
    
    Forms!frClaims.Requery
    
'   Cleanup:
    Set objCredit = Nothing
'   Execute cleanup queries:
    DoCmd.SetWarnings False
        DoCmd.OpenQuery "qaReimCls1"            ' - Close an approved warranty (tblClaimes!clDateCls = current date)
        DoCmd.OpenQuery "qaReimCls2"            ' - delete the closing date in rejected or partially rejected claims
        DoCmd.OpenQuery "MarinaImport"          ' - copy the import table to Workplace #2 (Kiev)
        DoCmd.OpenQuery "AbsentCI_Add"          ' - Add motorcycle claims invoices (imported separately) to tblCredNo
    DoCmd.SetWarnings True

'   Create a task in Outlook
'   to control the sending of the approved parts report
'   to the Parts Department:
    Call ApprPartsTask
    
'   Calculate the Exchange Rate on the current Credit Invoice,
'   and save it in the table tblCredNo (!cnExcRate):
    Call ExchRateSave
    
'   Check if the current Credit Invoice has any unallocated (hanging) records,
'   and set access to the corresponding button:
    X = fnBtnNewSndSw("ck3ForgetMeNot", Forms!frClaims!cmdFMN)
    If X Then
        X = fnBtnNewSndSw("ck1CI", Forms!frClaims!cmdFMN)
    End If

End Sub

' -------------------------------------------------------------------
' Event #10       : cmdFilter2_Click
' Purpose         : Applies a filter to display only claims marked
'                   as being in correction.
' Behavior        : Filters the current formТs records to show only
'                   claims where "clCorr = True".
' Error Handling  : None.
' -------------------------------------------------------------------
Private Sub cmdFilter2_Click()

    DoCmd.ApplyFilter , "clCorr = True" ' - Apply a filter to show claims marked for correction
    
End Sub

' -------------------------------------------------------------------
' Event #11       : cmdIIW_Click
' Purpose         : Handles the import of warranty claims sent by the dealer "Impulse" in a single CSV file format.
'
' Behavior        :
'                   1. Prompts the user to specify the file path for the CSV file.
'                   2. Clears temporary tables and imports data from the CSV file.
'                   3. Assigns unique claim numbers for each record.
'                   4. Adds claims, defects, and parts data to the main database tables.
'                   5. Updates customer transit status for imported claims.
'                   6. Displays the number of imported claims in a message box.
'                   7. Refreshes the form to reflect the updated data.
'
' External Calls  : ImpulseImportComments (Global Function in GlobalFn.bas): Handles the import of comments for the claims.
'
' Error Handling  : Displays a message if the file path is not specified.
' -------------------------------------------------------------------
Private Sub cmdIIW_Click()

    Dim strPath As String           ' - Path to the import file
    Dim lnNextClm As Long           ' - Last registered claim number
    Dim rst As Recordset            ' - Recordset for assigning claim numbers
    Dim btClmStep As Byte           ' - Increment step for claim numbers
    Dim X As Integer                ' - Used to display the number of imported claims
    
    strPath = DLookup("[dlrInboxPath]", "tblDealers", "[dlrNo] = 1")        ' - Default path for the import file
    btClmStep = DLookup("[opClaimNr]", "tblOptions")                        ' - Read the increment step for claim numbers
'   Prompt the user to specify the file path:
    strPath = InputBox("¬ведите маршрут файла импорта:", "»мпорт гарантии дилера", strPath)
    
'   Exit the procedure if no path is specified:
    If strPath = "" Then
        MsgBox "No file path specified. Operation canceled."
        Exit Sub
    End If
    
    DoCmd.SetWarnings False                                                     ' - Disable warnings before executing queries
'       Clear temporary tables for importing data:
        DoCmd.OpenQuery "qClearDlrWarrImport"
        DoCmd.OpenQuery "qClearDCC"

        DoCmd.TransferText acImportDelim, "DlrWarr", "tblDlrWarrImp", strPath   ' - Import data from the specified file
        DoCmd.OpenQuery "qaDlrImpClm"                                           ' - Generate unique numbers of imported claims
    
        lnNextClm = DMax("[clNo]", "tblClaimes")                                ' - Retrieve the next claim number
    
'       Assign unique claim numbers:
        Set rst = CurrentDb.OpenRecordset("tblDlrImpClm")
            With rst
                .MoveFirst
                While Not .EOF
                    .Edit
                    !Clm = lnNextClm + btClmStep
                    lnNextClm = !Clm
                    .Update
                    .MoveNext
                Wend
            End With
        
        DoCmd.OpenQuery "qaImpClmAdd"               ' - Claims import
        DoCmd.OpenQuery "qaImpDfAdd"                ' - Defect-codes import
        DoCmd.OpenQuery "qaImpPartAdd"              ' - Items import
        Call ImpulseImportComments                  ' - Comments import
        DoCmd.OpenQuery "qaTransitImp"              ' - Control of Transit/Non-transit status.
    DoCmd.SetWarnings True                          ' - Enable warnings after executing queries
    
'   Refresh the form to reflect the updated data:
    Me.Requery
'   Display the number of imported claims:
    X = MsgBox("Number of imported claims: " & rst.RecordCount & " клеймов.", vbOKOnly + vbInformation, "Import Complete")
    rst.Close
    
End Sub

' -------------------------------------------------------------------
' Event #12       : cmdNewSendOpen_Click
' Purpose         : Opens the form "frNewClmSend" to display the list of new warranty claims ready for submission.
' -------------------------------------------------------------------
Private Sub cmdNewSendOpen_Click()

    DoCmd.OpenForm "frNewClmSend"
    
End Sub

' -------------------------------------------------------------------
' Event #13       : cmdRjResendOpen_Click
' Purpose         : Opens the form "frRjResend" to display the list of
'                   rejected and corrected warranty claims ready for re-submission.
' -------------------------------------------------------------------
Private Sub cmdRjResendOpen_Click()

    DoCmd.OpenForm "frRjResend"
    
End Sub

' -------------------------------------------------------------------
' Event #14       : dtRepair_AfterUpdate
' Purpose         : Validates the repair date entered in the "dtRepair" field to ensure it is not excessively old.
'
' Behavior        :
'                   1. Checks if the entered repair date is more than
'                      20 days earlier than the current date.
'                   2. If the repair date is excessively old, displays
'                      a warning message to the user.
' -------------------------------------------------------------------
Private Sub dtRepair_AfterUpdate()

    If Me!dtRepair < Date - 20 Then
        MsgBox "Warning! The repair date entered is excessively old."
    End If

End Sub

' -------------------------------------------------------------------
' Event #15       : Form_Current
' Purpose         : Executes tasks and updates form elements when
'                   the user navigates to a different record in the form.
'
' Behavior        :
'                   1. Updates the BMW comments field caption to reflect the maximum allowed characters (210).
'                   2. Adjusts the visibility and accessibility of form controls based on the type of claim and its stage.
'                   3. Checks if the vehicle's VIN exists in the database and prompts for additional data if required.
'                   4. Updates the default dealer ID for new claims.
'                   5. Requeries the BMW send list and updates the displayed value.
'                   6. Handles filter states and updates form flags and counters.
'
' Error Handling  : None.
' -------------------------------------------------------------------
Private Sub Form_Current()

    txtRemBMW.Controls(0).Caption = " оментарии дл€ Ѕћ¬ (max 210):"     ' - Update the caption for the BMW comments field
    lbCommentsBMW.ForeColor = -2147483630                               ' - Set text color to black
    
'   Adjust form controls based on the current record:
    With Forms!frClaims!sf2Parts.Form                                   ' - work by default with the work-parts subform
        If Forms!frClaims!sfDefects.Form!dfWStage = 7 Then
'       If the current BSI (BMW Service Inclisive) claim:
            !DlrBuy.Enabled = False                                     ' - Hide warranty price
            !DlrActBSI.Visible = True                                   ' - Show BSI price
            !ptAct.Visible = False                                      ' - Hide warranty invoice number
            !ptActBSI.Visible = True                                    ' - Show BSI invoice number
        Else
'       else (if Warranty Claim):
            !DlrBuy.Enabled = True                                      ' - Show warranty price
            !DlrActBSI.Visible = False                                  ' - Hide BSI price
            !ptAct.Visible = True                                       ' - Show warranty act number
            !ptActBSI.Visible = False                                   ' - Hide BSI act number
        End If
    End With
        
'   Check and prompt for additional data for new vehicles:
    With Me
        If (!VIN.Tag <> "NewVehicle") And (.clWType = 1) And (Not IsNull(!VIN)) And CurrentUser() = !clManager Then
            If .isExistVIN Then
                !VIN.Tag = "Exist"                                      ' - Mark VIN as existing if all necessary fields are filled
            Else
'               If any field from (Serie / Model / Engine) is not filled in -
'               Open a form with an NF (Not Filled) argument, which is then processed when opened:
                DoCmd.OpenForm "dlgNewVcl", , , "luVIN = '" & Me!VIN & "'", DataMode:=acFormEdit, WindowMode:=acDialog, _
                    OpenArgs:="NF"
            End If
        End If
    End With

    Me!clDealerID.DefaultValue = Me!clDealerID.Value                    ' - Update the default dealer ID for new claims

'   Requery the BMW send list and update the displayed value:
    lstSendNr.Requery
    lstSendNr = DMax("[sndBMWNo]", "tblBMWNo", "[sndClaimNo] = " & Forms!frClaims!clNo)
    
'   Update the counter for active claims:
    Me.ptDA = CInt(DCount("*", "tblClaimes", "[clActive] = False AND [clClosed] = False"))
    
'   Corrects the situation when removing the filter using the standard button (on the toolbar):
    If Me.FilterOn = False Then     ' - If the filter is off
        swfltrDA = False            ' - unpress the button
    End If
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #16       : Form_Open
' Purpose         : Executes initial setup tasks and configures the state of form controls when the form is opened.
'
' Behavior        :
'                   1. Resets filters and initializes form elements.
'                   2. Checks for quantity discrepancies and enables or disables relevant controls.
'                   3. Verifies the availability of current currency exchange rates and downloads them if missing.
'                   4. Updates the EUR rate display on the main form.
'                   5. Adjusts the visibility of buttons for creating warranty and BSI acts based on the current state.
'
' External Calls  :
'                   - objDivQ.RecQ (Method of Class: DivQ)          : Checks for unresolved quantity discrepancies.
'                   - objWarrAct.WarQ (Method of Class: WarrAct)    : Returns the count of claims ready for warranty acts.
'                   - objBSIAct.bsiQ (Method of Class: BSI)         : Returns the count of claims ready for BSI acts.
'                   - fnCurrentRt (Global Function in GlobalFn.bas) : Verifies if the current currency rate is available.
'                   - getEUR (Global Function in GlobalFn.bas)      : Downloads the current exchange rate if not available.
'
' Error Handling  :
'                   - Handles missing currency rates by prompting for manual input.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
        
    Dim blDivOK As Boolean                          ' - Discrepancies in quantities have been resolved (marker)
    
    Me.FilterOn = False                             ' - Reset filters
    swfltrDA = False                                ' - Set the button to the unpressed state
    
    Set objDivQ = New DivQ                          ' - quantity discrepancies
    Set objWarrAct = New WarrAct                    ' - warranty invoices
    Set objBSIAct = New BSI                         ' - BSI invoices
    
    If Cancel = False Then
        DoCmd.Maximize                              ' - Maximize the form
    
'       Check for quantity discrepancies:
        If objDivQ.RecQ > 0 Then
            Me!cmdDivQ.Enabled = True               ' - enable the quantity adjustment button
            blDivOK = False
        Else
            Me!cmdDivQ.Enabled = False              ' - disable the quantity adjustment button
            blDivOK = True
        End If
        
'       Enable or disable warranty invoices button based on conditions:
        If blDivOK And objWarrAct.WarQ > 0 Then     ' If quantity discrepancies are resolved and there are warranty payments:
            Me!cmdDlrAct.Enabled = True             '   - enable the button for creating warranty invoices
        Else                                        ' else:
            Me!cmdDlrAct.Enabled = False            '   - disable the button for creating warranty invoices
        End If
        
'       Enable or disable BSI act button based on conditions:
        If blDivOK And objBSIAct.bsiQ > 0 Then      ' If quantity discrepancies are resolved and there are BSI payments:
            Me!cmdActBSI.Enabled = True             '   - enable the button for creating BSI invoices
        Else                                        ' else:
            Me!cmdActBSI.Enabled = False            ' - disable the button for creating BSI invoices
        End If
        
    End If
    
    ' Check if the current currency exchange rate is available:
    If fnCurrentRt() = 0 Then                                                       ' - If the current rate is not available - download
        If Not Me.getEUR Then                                                       ' If the download failed:
            MsgBox "Failed to download currency rates. Enter manually."             '   - inform the user
            DoCmd.OpenForm "frHistEUR", DataMode:=acFormAdd, WindowMode:=acDialog   '   - enter rate manually
        End If
    End If
    
    Me!clcCurRate = DLookup("[eurRt]", "$$ActualCur")   ' - display EUR rate in the corresponding field of the main form
    
End Sub

' --------------------------------------------------------------------------------------------------------------------------
' Event #17       : lstSendNr_AfterUpdate
' Purpose         : Updates the subform displaying BMW send details when a new record is selected in the list box "lstSendNr".
'
' Behavior        :
'                   1. Requeries the "tblBMWNo" subform to reflect the details of the selected send record.
'                   2. Ensures the subform displays the most current data related to the selected BMW send number.
' --------------------------------------------------------------------------------------------------------------------------
Private Sub lstSendNr_AfterUpdate()
    
    tblBMWNo.Requery    ' - Refresh the subform to display details for the selected send number
    
End Sub

' --------------------------------------------------------------------------------------------------------------------------
' Event #18       : objCredit_CreditImported
' Purpose         : Handles the "CreditImported" event triggered by the "NewCredit" object when credit data is successfully imported.
'                   Updates the credit object properties with imported data and calculates additional statistics.
'
' Behavior        :
'                   1. Updates the credit object properties:
'                      - Credit number (`pptCredNo`)
'                      - Credit approval date (`pptCredDate`)
'                      - Number of imported records (`pptImpRecCount`)
'                   2. Calculates and updates:
'                      - Number of rejected claims (`pptCountRjCl`)
'                      - Number of rejected individual items (`pptCountRjPos`)
'                   3. Uses queries to gather necessary data and updates the credit object properties.
'
' External Calls  :
'                   - Queries `qRjClaimesCount` and `qRjPosCount`
'                     (used to calculate rejected claims and items).
' --------------------------------------------------------------------------------------------------------------------------
Private Sub objCredit_CreditImported(CredNo As String, CredDate As Date, iRecCount As Integer)

    Dim rst1 As Recordset                   ' - Recordsets for rejected claims
    Dim rst2 As Recordset                   ' - Recordsets for rejected items
        
'   Open recordsets to calculate the number of rejected claims and items:
    Set rst1 = CurrentDb.OpenRecordset("qRjClaimesCount", dbOpenDynaset, dbReadOnly)
    Set rst2 = CurrentDb.OpenRecordset("qRjPosCount", dbOpenDynaset, dbReadOnly)
    
'   Update the credit object properties with imported data and statistics:
    With objCredit
        .pptCredNo = CredNo               ' - Set credit number
        .pptCredDate = CredDate           ' - Set credit approval date
        .pptImpRecCount = iRecCount       ' - Set number of imported records
        .pptCountRjCl = rst1.RecordCount  ' - Set number of rejected claims
        .pptCountRjPos = rst2.RecordCount ' - Set number of rejected items
    End With
    
    rst1.Close
    rst2.Close

End Sub

' -------------------------------------------------------------------
' Event #19       : OrdNo_AfterUpdate
' Purpose         : Validates the uniqueness of the order number entered in the "OrdNo" field
'                   to prevent duplicate processing of the same claim.
' Behavior        :
'                   1. Checks if the entered order number already exists
'                      in the database for the same dealer.
'                   2. If a duplicate order number is found, displays a
'                      warning message to the user.
'
' Error Handling  : Displays a warning message if a duplicate order number is found.
' -------------------------------------------------------------------
Private Sub OrdNo_AfterUpdate()
    
    Dim vrDoubleOrdNo As Variant        ' - Stores the first matching record (if any)
    Dim strCriteria As String           ' - Criteria for filtering records
    Dim strWarning As String            ' - Warning message text
    
'   Each dealer has its own unique order numbering.
'   Therefore, the criterion consists of dealer number and order number:
    strCriteria = "[clOrdNo] =" & Me!OrdNo & " AND [clDealerID] =" & Me!clDealerID

'   If a record corresponding to strCriteria is already found in the table tblClaimes -
'   then we are dealing with a duplicate record:
    vrDoubleOrdNo = DFirst("[clOrdNo]", "tblClaimes", strCriteria)
    
'   If a duplicate is found, display a warning message:
    If Not IsNull(vrDoubleOrdNo) Then
        strWarning = "A claim with this order number already exists."
        MsgBox strWarning, vbExclamation, "Duplicate Order Number"
    End If

End Sub

' -------------------------------------------------------------------
' Event #20        : sfDefects_Enter
' Purpose          : Handles the event when the user enters the "sfDefects" subform.
'                    Verifies mileage input accuracy and checks vehicle data.
' Behavior         :
'                     1. If the vehicle is new and linked to warranty type 1, prompts the user to input vehicle data.
'                     2. Compares the entered mileage with existing data and warns if a higher mileage is found.
'                     3. If vehicle data is missing, opens a form for data entry.
' Notes            :
'                     Tags used for process control:
'                       - sf1Defects.Tag:
'                         - "Auto" - automatic claim processing (e.g., dealer warranty import).
'                         - "" - manual claim processing (set by Form_Current event).
'                       - VIN.Tag:
'                         - "NewVehicle" - vehicle not found in the reference table.
'                         - "Exist" - vehicle found or added to the reference table.
'                       - clMileage.Tag:
'                         - "New" - mileage is new (no higher mileage found in history).
'                         - "Old" - mileage is old (higher mileage found in history).
' -------------------------------------------------------------------
Private Sub sfDefects_Enter()
' Handles entry into the "sfDefects" subform.
' Validates mileage and checks vehicle data.
    
'   If the process is automated (sf1Defects.Tag = "Auto"),
'   check only if vehicle data needs to be added:
    If Form_sf1Defects.Tag = "Auto" Then
'       If the vehicle is new and linked to warranty type 1,
        If Me!VIN.Tag = "NewVehicle" And Forms!frClaims!clWType = 1 Then
'           open a form to input vehicle data:
            DoCmd.OpenForm "dlgNewVcl", DataMode:=acFormAdd, WindowMode:=acDialog, _
                OpenArgs:=Me!VIN
            Me!VIN.Tag = "Exist"                            ' - Mark the vehicle as existing
        End If
        Exit Sub
    End If

'   If manual processing (sf1Defects.Tag = ""),
'   validate mileage for warranty type 1:
    Dim lnHigherMileage As Variant                          ' - Stores the highest mileage found in the database
    Dim lnCurrentMileage As Long                            ' - Current mileage input by the user
    Dim strCriteria As String                               ' - Filter criteria for mileage comparison
    Dim strWarning As String                                ' - Warning message for higher mileage
  
    If Me!clWType = 1 And Me!clMileage.Tag = "New" Then
        lnCurrentMileage = Nz(Me!clMileage)                 ' - Get the current mileage or default to 0

'       Define the criteria to search for higher mileage entries in the database:
        strCriteria = "[clVIN] ='" & Me!VIN & "' AND [clMileage]>" & str(lnCurrentMileage)
        strCriteria = strCriteria & " AND [clWType] = 1"
        lnHigherMileage = DFirst("[clMileage]", "tblClaimes", strCriteria)
    
'       Warn the user if a higher mileage is found and update the mileage tag:
        If Not IsNull(lnHigherMileage) Then
            strWarning = "A higher mileage value was found: " & lnHigherMileage
            MsgBox strWarning, vbExclamation
            Me!clMileage.Tag = "Old"                        ' - Mark the mileage as outdated.
        End If
        
        If Me!VIN.Tag = "NewVehicle" Then
'       If the vehicle is new (VIN.Tag = "NewVehicle"), prompt the user to add it:
            DoCmd.OpenForm "dlgNewVcl", DataMode:=acFormAdd, WindowMode:=acDialog, _
                OpenArgs:=Me!VIN
            Me!VIN.Tag = "Exist"                            ' - Update the vehicle tag to indicate existence
        End If
    
    End If
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #21        : swSubformEdit_AfterUpdate
' Purpose          : Toggles the editability of specific fields in the "sf2Parts" subform based on the toggle button state.
' Behavior         :
'                     1. When the toggle button is ON (True), enables editing of restricted fields in the subform.
'                     2. When the toggle button is OFF (False), disables editing of restricted fields and resets focus to "WorkID."
' Notes            :
'                     - Affects fields related to compensation, rejection codes, credit numbers, and dealer information.
'                     - Field `ckAWTwrnt` (Importer Warranty Marker):
'                       - Indicates whether a position will be included in the reimbursement invoice to the dealer at the
'                         importer purchase price, regardless of manufacturer approval status.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub swSubformEdit_AfterUpdate()
'   Get the reference to the "sf2Parts" subform:
    Dim frm As Form
    Set frm = Me![sf2Parts].Form
    
    With frm
'   Enable or disable fields based on the toggle button state:
        If Me!swSubformEdit = True Then
            !CompensBMW.Enabled = True          ' - compensation field
            !RjCode.Enabled = True              ' - rejection code field
            !CredNo.Enabled = True              ' - credit number field
            !ckAWTwrnt.Enabled = True           ' - Importer warranty marker field
            !DlrInvoice.Enabled = True          ' - dealer invoice field
        Else
            !WorkID.SetFocus                    ' - Reset focus to "WorkID" field
            !CompensBMW.Enabled = False         ' - Disable compensation field
            !RjCode.Enabled = False             ' - Disable rejection code field
            !CredNo.Enabled = False             ' - Disable credit number field
            !ckAWTwrnt.Enabled = False          ' - Disable warranty marker field
            !DlrInvoice.Enabled = False         ' - Disable dealer invoice field
        End If
    End With
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #22        : txtRemBMW_AfterUpdate
' Purpose          : Ensures that the "txtRemBMW" field (BMW comments) does not contain forbidden characters,
'                    such as carriage return (CR) or line feed (LF).
' Behavior         :
'                     1. Checks if the field contains CR or LF characters and removes them.
'                     2. Replaces CR or LF with spaces using a custom function `fnCrLfRemove` (GlobalFn module).
' Notes            :
'                     - CR and LF characters can block the transmission of claim files (glglgl20) to the manufacturer.
'                     - Applies validation only if the field is not empty.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub txtRemBMW_AfterUpdate()

'   Check if the field contains a value:
    If Not IsNull(Me!txtRemBMW) Then
'       The fnCrLfRemove function returns a string in which CrLf has been replaced by a SPACE character:
        Me!txtRemBMW = fnCrLfRemove(Me!txtRemBMW)
    End If

End Sub

' -------------------------------------------------------------------
' Event #23        : txtRemBMW_Change
' Purpose          : Tracks the number of characters entered in the "txtRemBMW" field (BMW comments)
'                    and provides a visual warning when the character limit is exceeded.
' Behavior         :
'                     1. Counts the current number of characters in the field.
'                     2. Updates the associated label to display the character count.
'                     3. Changes the label text color to red if the count exceeds
'                        the maximum allowed (210 characters).
'                     4. Resets the label text color to black if within the limit.
' Notes            :
'                     - Maximum allowed characters: 210.
'                     - The label's default color is black; it turns red as a warning.
' -------------------------------------------------------------------
Private Sub txtRemBMW_Change()
' ѕодсчет количества введенных символов в поле комментариев Ѕћ¬ (допускаетс€ максимум 210)
'
    Dim strLen As Integer                   ' - Stores the length of the text in "txtRemBMW"
    
    strLen = Len(Me.txtRemBMW.Text)         ' - Get the current character count in the text box
    
'   Change label color based on character count:
    If strLen > 209 Then
        lbCommentsBMW.ForeColor = BRIGHT_RED    ' - Set label color to red if limit exceeded
    Else
        lbCommentsBMW.ForeColor = -2147483630   ' - Set label color to black if within limit
    End If
    
'   Update the label to show the character count:
    txtRemBMW.Controls(0).Caption = "Entered: " & CStr(strLen)

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #24        : txtRemBMW_Undo
' Purpose          : Restores the previously saved value in the "txtRemBMW" field (BMW comments) when the undo action is triggered.
' Behavior         :
'                     1. Updates the associated label with the restored character count.
'                     2. Resets the label's text color to black, in case it was previously red due to exceeding the character limit.
'
' Notes            : The event is triggered when the user cancels changes in the "txtRemBMW" field using the undo operation.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub txtRemBMW_Undo(Cancel As Integer)

    txtRemBMW.Controls(0).Caption = "¬ведено: " & Len(Me!clCommentBMW)      ' - Update the label with the character count of the restored value
    lbCommentsBMW.ForeColor = -2147483630                                   ' - Reset the label's text color to black
    
End Sub

' -------------------------------------------------------------------------------------------------------------
' Event #25        : VIN_AfterUpdate
' Purpose          : Validates the VIN entered by the user and checks for its presence in the reference table.
'                    Updates related warranty fields and provides feedback on specific conditions.
' Behavior         :
'                     1. Searches for the entered VIN in the reference table.
'                     2. If found:
'                        - Updates the "Wstart" field with the warranty start date.
'                        - Displays any associated remarks (if available).
'                     3. If not found:
'                        - Marks the vehicle as new ("NewVehicle").
'                     4. Checks if the VIN is listed in sold vehicles or motorcycles tables:
'                        - Updates the "clCustType" field accordingly.
'                     5. Applies a filter to display previous claims for the same VIN.
'                     6. Alerts the user if multiple previous claims are found.
' Notes            :
'                     - VIN.Tag values:
'                       "NewVehicle" - VIN is not found in the reference table.
'                       "Exist" - VIN exists in the reference table.
'                     - The event is triggered after the user updates the VIN field.
' -------------------------------------------------------------------------------------------------------------
Private Sub VIN_AfterUpdate()

    Dim rst As Recordset                                ' - Recordset for querying the reference table
    Dim strSQL As String                                ' - SQL query string
    Dim qryDef As QueryDef                              ' - Query definition for dynamic parameterized queries
    
'   Search for the VIN in the reference table
'   using a parameterized query:
    Set qryDef = CurrentDb.QueryDefs("prgVclLkUp")
        qryDef("prmVIN") = Forms!frClaims!VIN
    Set rst = qryDef.OpenRecordset()
        
    With rst
        If Not (rst.BOF And rst.EOF) Then               ' - VIN found in the reference table
            Me!VIN.Tag = "Exist"
            If Me!clWType = 1 Then                      ' - If warranty type is 1, update the "Wstart" field with the start date
                If (IsNull(!luWstart) Or !luWstart = #12:00:00 AM#) And Not IsNull(Me!Wstart) Then
'                   Update the warranty start date in the reference table if missing:
                    .Edit
                        !luWstart = Me!Wstart           ' - Take the warranty start date from the claim and store it in the reference table
                    .Update
                    Me.Requery
                Else
                    Me!Wstart = !luWstart               ' - Populate the warranty start date field from the reference table
                End If
                
            End If
            If Not IsNull(!luRem) Then                  ' - Display any remarks associated with the vehicle
                MsgBox !luRem, vbExclamation + vbOKOnly, "Attention!"
            End If
        Else
            Me!VIN.Tag = "NewVehicle"                   ' - VIN not found in the reference table
        End If
        
        .Close
    End With
    
'   Check if the VIN exists in the sold vehicles table:
    If Not IsNull(DLookup("[sVIN]", "lkSold", "[sVIN]='" & Forms!frClaims!VIN & "'")) Then
        Me!clCustType = False                           ' - Not a transit customer
    Else
        If Me.clWType = 1 Then                          ' - Vehicle Warranty
            Me!clCustType = True                        ' - Transit customer
        End If
    End If
    
'   Check if the VIN exists in the sold motorcycles table (for Waranty Type = 1):
    If Not IsNull(DLookup("[sVIN]", "lkSoldMoto", "[sVIN]='" & Forms!frClaims!VIN & "'")) And Me.clWType = 1 Then
        Me!clCustType = False
    End If
    
'   Filter the form to show previous claims for the entered VIN:
    Me.Filter = "clVIN = '" & Me!VIN & "'"
    Me.FilterOn = True

'   Notify the user if multiple previous claims are found for the same VIN:
    If Me.RecordsetClone.RecordCount > 1 Then
        MsgBox "Previous claims found: " & str(Me.RecordsetClone.RecordCount - 1)
    End If
    
    SendKeys "{Enter}"                                  ' - Simulate an Enter key press to finalize the update
        
End Sub

' -----------------------------------------------------------------------------------------------------------
' Event #26        : cbDlrFilter_AfterUpdate
' Purpose          : Filters the claims displayed in the form based on the dealer selected in the combo box.
' Behavior         :
'                     1. Applies a filter to display claims for the selected dealer ID.
'                     2. If dealer ID 3 is selected, displays a message indicating it corresponds to an outdated dealer ID.
'                     3. Activates the filter for the form.
' Notes            :
'                     - The combo box "cbDlrFilter" allows the user to select a dealer ID from a predefined list.
'                     - Dealer ID 3 represents the old ID for "Bavaria Motors."
' -----------------------------------------------------------------------------------------------------------
Private Sub cbDlrFilter_AfterUpdate()
    
    Dim stLinkCriteria As String                                ' - Filter criteria for the form
    
'   Construct the filter condition based on the selected dealer ID:
    stLinkCriteria = "[clDealerID]=" & str(Me!cbDlrFilter)

'   Apply the filter to the form:
    Form_frClaims.Filter = stLinkCriteria
    Form_frClaims.FilterOn = True

    If Me!cbDlrFilter = 3 Then
'   If the selected dealer ID is 3, display a warning message:
        MsgBox "Dealer ID 3 corresponds to an outdated record for Bavaria Motors.", vbExclamation, "Dealer ID Warning"
    End If
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #27        : swView_AfterUpdate
' Purpose          : Filters the records displayed in the "sf2Parts" subform based on the selected toggle button state.
' Behavior         :
'                     1. When the toggle button is set to:
'                        - 1: Displays all records.
'                        - 2: Filters records to show only work items (ptType = 'W').
'                        - 3: Filters records to show only parts (ptType = 'P').
'                     2. Applies the filter condition to the "sf2Parts" subform.
' Notes            :
'                     - Toggle button "swView" options:
'                       1 = All records
'                       2 = Work items only
'                       3 = Parts only
'                     - Ensures that the filter is cleared when option 1 is selected.
' Error Handling   : None
' ---------------------------------------------------------------------------------------------------------------------
Private Sub swView_AfterUpdate()

    Dim frm As Form                         ' - Reference to the "sf2Parts" subform
    Set frm = sf2Parts.Form                 ' - Get the reference to the subform
    
'   Apply filter based on the toggle button state:
    Select Case swView
        Case 1                              ' - Option 1: Show all records
            frm.FilterOn = False
        Case 2                              ' - Option 2: Show only work items (ptType = 'W')
            frm.Filter = "ptType = 'W'"
            frm.FilterOn = True
        Case 3                              ' - Option 3: Show only parts (ptType = 'P')
            frm.Filter = "ptType = 'P'"
            frm.FilterOn = True
    End Select

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #28        : swfltrDA_AfterUpdate
' Purpose          : Toggles the application of a filter to display only
'                    inactive claims (clActive = False) that are not closed (clClosed = False).
' Behavior         :
'                     1. When the toggle button "swfltrDA" is ON (True), applies the filter to display only inactive and open claims.
'                     2. When the toggle button is OFF (False), removes the filter and displays all records.
' Notes            :
'                     - The filter criteria:
'                       "clActive = False AND clClosed = False"
'                     - Ensures that the filter state matches the toggle button state.
' Error Handling   :
'                     - Displays an error message if the filter operation fails.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub swfltrDA_AfterUpdate()
    
    Dim frm As Form                                         ' - Reference to the main form

On Error GoTo Err_fltrDA_Click

    Set frm = Forms!frClaims                                ' - Get the reference to the main form
    
    frm.Filter = "clActive = False AND clClosed = False"
    
'   Apply or remove the filter based on the toggle button state:
    If swfltrDA = False Then
        frm.FilterOn = False                                ' - Remove the filter
    Else
        frm.FilterOn = True                                 ' - Apply the filter
    End If

Exit_fltrDA_Click:
    Exit Sub

Err_fltrDA_Click:
'   Display an error message if an error occurs during the filtering process:
    MsgBox "An error occurred while applying the filter: " & Err.Description, _
           vbCritical, "Filter Error"
    Resume Exit_fltrDA_Click
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #29        : cmdCredRep_Click
' Purpose          : Executes a macro to generate and display the warranty credit report for the current claims data.
' Behavior         :
'                     1. Runs the macro "mbkWarrBMW" to process and display the warranty credit report.
'                     2. Ensures the user can view the report in the required format.
' Notes            :
'                     - The macro "mbkWarrBMW" is responsible for creating and opening the warranty credit report.
'                     - This event is triggered when the "cmdCredRep" button is clicked.
' Error Handling   :
'                     - Displays an error message if the macro fails to execute.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub cmdCredRep_Click()
On Error GoTo Err_cmdCredRep_Click

    Dim stDocName As String         ' - name of the macro

    stDocName = "mbkWarrBMW"
    DoCmd.RunMacro stDocName        ' - Run the specified macro
    
Exit_cmdCredRep_Click:
    Exit Sub

Err_cmdCredRep_Click:
'   Display an error message if the macro fails:
    MsgBox "An error occurred while generating the credit report: " & Err.Description, _
           vbCritical, "Report Error"
    Resume Exit_cmdCredRep_Click
    
End Sub

' -------------------------------------------------------------------------------------------------
' Event #30        : cmdDlrAct_Click
' Purpose          : Generates and opens the dealer warranty claims report.
' Behavior         :
'                     1. Updates the auxiliary table "tb$lAux" with the current date
'                        to ensure correct functionality of the DTPicker control.
'                     2. Executes the report "rpDlrAct" in preview mode.
' Notes            :
'                     - The auxiliary table "tb$lAux" is used to store the report date
'                       for proper filtering and data representation.
'                     - The report "rpDlrAct" provides a summary of dealer warranty claims.
' Error Handling   :
'                     - Handles the "2501" error, which occurs if the report preview is canceled.
'                     - Displays an error message for other unexpected errors.
' -------------------------------------------------------------------------------------------------
Private Sub cmdDlrAct_Click()
    Dim stDocName As String                             ' - name of the report
    Dim stLinkCriteria As String

On Error GoTo Err_cmdDlrAct

'   Ensure the auxiliary table has the current date
'   for DTPicker functionality:
    If DLookup("[auxDt]", "tb$lAux") <> Date Then
        DoCmd.SetWarnings False
            DoCmd.OpenQuery "qa$DTPicker"               ' - Update the date in the auxiliary table
        DoCmd.SetWarnings True
    End If
    
'   Define the report name and open it in preview mode:
    stDocName = "rpDlrAct"
    DoCmd.OpenReport stDocName, acPreview
    
Exit_cmdDlrAct:
    Exit Sub

Err_cmdDlrAct:
'   Handle the case where the report preview is canceled:
    If Err.Number = 2501 Then
        Resume Next
    Else
'   Display other errors:
        MsgBox "An error occurred while generating the dealer report: " & Err.Description, _
               vbCritical, "Report Error"
        Resume Exit_cmdDlrAct
    End If
    
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Event #31       : cmdWarBlankImport_Click
' Purpose         : Imports warranty claim data from an Excel warranty blank provided by the dealer (excluding Emerald Motors).
'
' Behavior        :
'                   1. Checks if the warranty claim is already closed (indicated by a non-empty "clOrdNo" field).
'                      - If closed, prevents import and notifies the user.
'                   2. Calls "warrImportTopFromExcell" to import header data from the Excel file.
'                   3. Based on the global flag "blBSI", calls one of:
'                      - "bsiImportDataFromExcell" for BSI claims.
'                      - "warrImportDataFromExcell" for standard claims.
'                   4. Refreshes the form to unlock the "tblVclLkUp" table.
'                   5. Executes VIN validation ("VIN_AfterUpdate") and
'                      comprehensive checks ("sfDefects_Enter").
'
' External Calls  :
'                   - warrImportTopFromExcell (Global Procedure)    : Imports the header data from the Excel file.
'                   - bsiImportDataFromExcell (Global Procedure)    : Imports data specific to BSI claims.
'                   - warrImportDataFromExcell (Global Procedure)   : Imports data specific to standard warranty claims.
'                   - VIN_AfterUpdate (current Form Event)          : Performs VIN validation after import.
'                   - sfDefects_Enter (current Form Event)          : Executes comprehensive defect checks.
'
' Error Handling  :
'                   - Displays a message if the warranty claim is already closed.
' ---------------------------------------------------------------------------------------------------------------------
Private Sub cmdWarBlankImport_Click()

'   Prevent import if the warranty claim is already exist:
    If Not IsNull(Me!clOrdNo) Then
        MsgBox "The warranty claim is already exist. Data import is not allowed.", _
               vbExclamation, "Import Not Allowed"
        Exit Sub
    End If
    
    Call warrImportTopFromExcell            ' - Import the header data from the Excel file
    
    If blBSI Then                           ' - Import the claim details based on the BSI flag
        Call bsiImportDataFromExcell        ' - Import BSI claim data
    Else
        Call warrImportDataFromExcell       ' - Import standard warranty claim data
    End If
    
    Me.Refresh                              ' - Refresh the form to unlock the tblVclLkUp table
    
'   Perform additional checks:
    Call VIN_AfterUpdate                    ' Validate VIN
    Call sfDefects_Enter                    ' Perform comprehensive checks

End Sub

' -------------------------------------------------------------------------------------------------------------------------------
' Property #1     : ptDA
' Purpose         : Controls and displays the count of deactivated warranty claims (removed from sending but not closed).
'
' Behavior        :
'                   - "Get": Returns the current count of deactivated claims from the private variable "intptDA".
'                   - "Let": Updates the count of deactivated claims and:
'                       1. Adjusts the font style and color of the label "swfltrDA" based on whether there are deactivated claims.
'                       2. Updates the label caption with the new count in the format "—ќ: <count>".
'                       3. Stores the new value in the private variable "intptDA".
'
' Key Features    :
'                   - Uses conditional formatting to visually highlight the presence of deactivated claims.
'                   - Ensures the display is dynamically updated whenever the count changes.
' -------------------------------------------------------------------------------------------------------------------------------
' Getter for ptDA:
Public Property Get ptDA() As Integer

    ptDA = intptDA

End Property
' Setter for ptDA:
Public Property Let ptDA(ByVal NewVal As Integer)

'   Adjust label appearance based on the count of deactivated claims:
    If NewVal > 0 Then
        swfltrDA.FontBold = True                ' - Bold font for active deactivated claims
        swfltrDA.ForeColor = BRIGHT_RED         ' - Highlight with red color
    Else
        swfltrDA.FontBold = False               ' - Regular font for no deactivated claims
        swfltrDA.ForeColor = GRAY               ' - Dim the label
    End If

    swfltrDA.Caption = "—ќ: " & CStr(NewVal)    ' - Update the label caption to display the new count
    
    intptDA = NewVal                            ' - Store the new value in the private variable
    
End Property

' -------------------------------------------------------------------------------------------------------------
' Property #2     : isExistVIN
' Purpose         : Determines whether the vehicle identified by the VIN (field `clVIN` in the form context)
'                   exists in the database table `tblVclLkUp` and whether all necessary vehicle data is complete.
'
' Behavior        :
'                   1. Searches for the VIN in the `tblVclLkUp` table.
'                   2. Returns `False` if the VIN is not found.
'                   3. If the VIN is found:
'                       - Verifies the completeness of vehicle data:
'                         - `luSerial` (Serial number)
'                         - `luModel` (Model)
'                         - `luEng` (Engine, mandatory for cars only).
'                       - Returns `True` if all required fields are filled,
'                         otherwise `False`.
'
' External Calls  :
'                   - isMoto (Global Function in GlobalFn.bas)
'                     Determines if the VIN belongs to a motorcycle.
' -------------------------------------------------------------------------------------------------------------
Public Property Get isExistVIN() As Boolean

    Dim rst As Recordset        ' - Recordset to search VIN in the table
    Dim strCriteria As String   ' - Filter criteria for finding the VIN
    Dim blChk As Boolean        ' - Result of data completeness check
    
    Set rst = CurrentDb.OpenRecordset("tblVclLkUp")     ' - Initialize the recordset
    strCriteria = "[luVIN] = '" & Me!VIN & "'"          ' - and set the search criteria
    blChk = True                                        ' - Assume all fields are complete initially
    
    With rst
        
        .FindFirst strCriteria                          ' - Search for the VIN in the table
        If .NoMatch Then                                ' - If VIN is not found
            isExistVIN = False                          ' - return False
        Else                                            ' - If the VIN is found:
            ' Check completeness of vehicle data:
            If IsNull(!luSerial) Then
                blChk = False                           ' - Serial number is missing
            End If
            If IsNull(!luModel) Then
                blChk = False                           ' - Model is missing
            End If
            
            If Not isMoto(Me!VIN) Then                  ' - For non-motorcycles
                If IsNull(!luEng) Then                  ' - engine data must be complete
                    blChk = False                       ' - Engine data is missing
                End If
            End If

'           Set the property result based on data completeness:
            isExistVIN = blChk
        End If
    End With
    
    rst.Close
    Set rst = Nothing
    
End Property

' ---------------------------------------------------------------------------------------------------------------------
' Method #1       : getEUR
' Purpose         : Fetches currency exchange rates for USD and EUR from an online XML source and saves them to the `tb$HistEUR` table.
'
' Behavior        :
'                   1. Connects to the URL `http://bank-ua.com/export/currrate.xml` (this link is no longer valid).
'                   2. Parses the XML file to extract exchange rates for EUR (code 978) and USD (code 840).
'                   3. Calculates exchange rates using the `clcCR` function and stores them in the `tb$HistEUR` table.
'                   4. Records the current date and user details for auditing purposes.
'
' Calls           :
'                   - clcCR (Me: Method #2)   : Calculates the exchange rate based on nominal and value fields from the XML.
'
' Error Handling  :
'                   - Displays a message if the currency rates cannot be loaded or if an error occurs during execution.
'                   - Returns `False` if any error is encountered.
'                   - Returns `True` if execution is successful.
'
' Requirements    :
'                   - Microsoft XML v5.0 library must be enabled in
'                     Tools > References.
' ---------------------------------------------------------------------------------------------------------------------
Public Function getEUR() As Boolean

Dim url_request As String       ' - URL to fetch the currency rates
Dim nodeList As Object          ' - Collection of XML nodes
Dim xmldoc As Object            ' - XML document object
Dim xmlNode As Object           ' - Individual XML node

Dim i As Integer                ' - Loop variable for child nodes
Dim iIndex As Integer           ' - Loop variable for node collection
Dim rst As Recordset            ' - Recordset for tb$HistEUR table

On Error GoTo Err_getEUR

DoCmd.Hourglass True                                            ' - Show hourglass cursor during the operation
Set xmldoc = CreateObject("Msxml.DOMDocument")                  ' - Initialize the XML document:
xmldoc.async = False                                            ' - Disable asynchronous loading

' Define the URL for fetching currency rates:
url_request = "http://bank-ua.com/export/currrate.xml"  '?date_req=" + Format(Date, "dd\/mm\/yyyy")

' Load the XML document from the URL:
If Not xmldoc.Load(url_request) = True Then
   MsgBox "Currency rates not loaded.", vbExclamation, "Error"
   GoTo Err_getEUR
End If

Set rst = CurrentDb.OpenRecordset("tb$HistEUR", , dbAppendOnly) ' - Open a recordset for the tb$HistEUR table

With rst
    .AddNew                     ' - Add a new record to the table
    !eurDt = Date               ' - Save the current date

' Get the collection of nodes for processing:
Set nodeList = xmldoc.selectNodes("*/item")
' Iterate through each node in the collection:
For iIndex = 0 To nodeList.Length - 1
'     get a copy of the node with index iIndex from the collection of nodes with all child nodes:
      Set xmlNode = nodeList.Item(iIndex).CloneNode(True)
        For i = 0 To xmlNode.childNodes.Length - 1              ' - move through all child nodes of the selected node
            If xmlNode.childNodes(i).baseName = "code" Then     ' - If the node name is УcodeФ
                If xmlNode.childNodes(i).Text = "978" Then      ' - and the value = У978Ф (Euro)
'                   Record the value of the Euro exchange rate in the table:
                    !eurRt = clcCR(xmlNode.childNodes(5).Text, xmlNode.childNodes(3).Text)
                End If
                If xmlNode.childNodes(i).Text = "840" Then      ' - if the node value = У840Ф (USD)
'                   Save the value of USD exchange rate in the table:
                    !usdRt = clcCR(xmlNode.childNodes(5).Text, xmlNode.childNodes(3).Text)
                End If
            End If
         Next i
Next iIndex
!curUser = "System"             ' - Record the user
.Update                         ' - Save the record
.Close                          ' - close the recordset
End With

' Clean up objects:
Set xmldoc = Nothing
Set nodeList = Nothing
Set xmlNode = Nothing

    getEUR = True               ' - Indicate success
    
Exit_getEUR:
    DoCmd.Hourglass False       ' - Restore the default cursor after completion
    Exit Function

Err_getEUR:
'   Handle errors and return False:
    MsgBox Err.Description, vbCritical, "Error"
    getEUR = False
    Resume Exit_getEUR

End Function

' ------------------------------------------------------------------------------------------------------------------
' Method #2       : clcCR
' Purpose         : Calculates the currency exchange rate based on the provided value and nominal using the formula:
'                   Round(strVal / strNominal, 4).
'
' Behavior        :
'                   1. Converts the `strNominal` argument to an integer.
'                   2. Converts the `strVal` argument to a single-precision floating-point number, ensuring compatibility
'                      with the decimal format.
'                   3. Divides `strVal` by `strNominal` and rounds the result to four decimal places.
'                   4. Returns the calculated exchange rate as a single.
'
' Parameters      :
'                   - strVal (String)       : The currency value extracted from an XML node
'                   - strNominal (String)   : The nominal value of the currency extracted from an XML node.
'
' External Calls  :
'                   - This method is called by `getEUR` to calculate exchange rates for USD and EUR.
' ------------------------------------------------------------------------------------------------------------------
Public Function clcCR(strVal As String, strNominal As String) As Single
'   ‘ункци€ рассчитывает курс валюты по формуле Round(strVal / strNom,4), предварительно преобразовав аргументы к числам
'       strVal - значение курса, вз€тое из XML-узла (см. getEUR)
'       strNominal - значение номинала валюты, вз€тое из XML-узла (см. getEUR)
'   ¬ызываетс€ из getEUR
    
    Dim intNom As Integer
    
    intNom = CInt(strNominal)                                       ' - Convert the nominal value to an integer
'   Calculate the exchange rate:
'   - Replace "." with "," for locale compatibility.
'   - Convert the currency value to a single.
'   - Divide by the nominal and round to four decimal places.
    clcCR = Round(CSng(Replace(strVal, ".", ",")) / intNom, 4)

End Function
