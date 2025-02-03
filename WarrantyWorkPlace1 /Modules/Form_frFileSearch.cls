VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================================================================================================
' Module Name  : frFileSearch
' Description  : Handles the search, processing, and management of ICS-related files and their data for integration into claims.
'                Includes functionality for saving paths, executing queries, and managing search progress.
'
' Key Features:
' 1. Performs advanced file searches based on specific criteria (e.g., path, file creation date, and predefined patterns).
' 2. Updates and validates search results in the database.
' 3. Provides user feedback through progress indicators.
'
' Dependencies:
' 1. basStatusMeter (Module in libOB.mda)   - Provides progress indicator functionality
' 2. AutoPlus (External DMS Database)       - Sources data via queries (e.g., qaAAFAP -> qryAutoPlus) connected to
'                                           tables (FF_020, FF_010) in the repair workshop DMS.
'
' Events:
' 1. btAutoPlus_Click     : Runs a query to fetch new acts for Auto+.
' 2. btFileSearch_Click   : Performs a file search and saves the results to
'                           a temporary table.
' 3. btSavePath_Click     : Saves the specified search path to the options table.
' 4. btAddNewRec_Click    : Executes a query to add new data to the claims table.
' 5. btFrmCls_Click       : Closes the form.
'
' External Calls:
'                   - acbInitMeter (Function in basStatusMeter, libOB.mda)    : Initializes the progress indicator.
'                   - acbUpdateMeter (Function in basStatusMeter, libOB.mda)  : Updates the progress indicator.
'                   - acbCloseMeter (Function in basStatusMeter, libOB.mda)   : Closes the progress indicator.
'
' Notes        : Ensure that the required queries (qaAAFAP, qaAddNewICS) and
'                tables (tblOptions, tmpICS) are properly configured.
'
' Developer    : Oleh Bondarenko
' Created      : 2006-03-30
' Last Updated : 2025-01-20 by Oleh Bondarenko - Added comments for GitHub upload
' =====================================================================================================================
Option Explicit
Option Compare Database

' -------------------------------------------------------------------
' Event #1        : btAutoPlus_Click
' Purpose         : Registration of new acts closed in Auto+ (adding to the register).
' Behavior        :
'                   - Opens and runs the query "qaAAFAP" to fetch new data.
'                   - The query relies on the external DMS tables "FF_020" and "FF_010".
' Notes           : Ensure the DMS system is accessible; otherwise, the query will fail.
' -------------------------------------------------------------------
Private Sub btAutoPlus_Click()
      
'   Registration of new acts closed in Auto+ (adding to the register):
    DoCmd.OpenQuery "qaAAFAP"                                       ' - qaAAFAP -> qryAutoPlus -> AutoPlus (tables FF_020 & FF_010)
    
End Sub

' -----------------------------------------------------------------------------------------------------------
' Event #2        : btFileSearch_Click
' Purpose         : Searches for ICS files in the specified path and saves the results to a temporary table.
' Behavior        :
'                   - Checks if the search path is specified and valid.
'                   - Initializes the progress indicator using acbInitMeter.
'                   - Uses Application.FileSearch to locate files in the specified path.
'                   - Filters files by creation date, extracting claim and act numbers.
'                   - Saves file details to the tmpICS table and updates the opLastICS field.
'                   - Updates the progress indicator using acbUpdateMeter and closes it upon completion.
'                   - Displays appropriate messages in case of errors or successful completion.
'
' External Calls  : acbInitMeter, acbUpdateMeter, acbCloseMeter (basStatusMeter).
'
' Notes           : Ensure that Application.FileSearch is supported and the tblOptions and tmpICS tables are properly configured.
' -----------------------------------------------------------------------------------------------------------
Private Sub btFileSearch_Click()
' Perform complex search using the FileSearch object and its PropertyTests collection.

    Dim varFile As Variant          ' - current file path
    Dim varFileDate As Variant      ' - last file modification date
    Dim strActNr As String          ' - act #
    Dim strClaimNr As String        ' - claim #
    Dim rst As Recordset            ' - Recordset for adding the search results to the tmpICS table
    Dim fOK As Boolean              ' - managing progress indicator updates
    Dim intCounter As Integer       ' - counting processed files during the search
    Dim rst2 As Recordset           ' - Recordset for updating the options table with the latest file date
    
    Call acbInitMeter("Search Status", True)                        ' - Initialize progress indicator
    
    intCounter = 0
    varFileDate = DLookup("[opLastICS]", "tblOptions")
    Set rst = CurrentDb.OpenRecordset("tmpICS", , dbAppendOnly)
    
'   Configure and execute file search:
    With Application.FileSearch
        .NewSearch
        
        If Not IsNull(txtPath) Then
            .LookIn = txtPath
        Else
            MsgBox "Specify the search path"
            Exit Sub
        End If
        
        .SearchSubFolders = True
        .FileType = msoFileTypeAllFiles
        .Execute
        
        If .FoundFiles.Count > 0 Then
            For Each varFile In .FoundFiles
                intCounter = intCounter + 1
                If CDate(varFileDate) < FileDateTime(varFile) Then
                    strClaimNr = Mid$(varFile, InStr(1, varFile, "awt") + 3, 6)
                    strActNr = Mid$(varFile, InStr(1, varFile, "awt") + 10, 6)
            
                    With rst                                    ' Save file details to tmpICS:
                        .AddNew                                 ' - Add a new record
                        !tmpAct = strActNr                      ' - Save an Invoice Nr
                        !tmpClaim = strClaimNr                  ' - Save a Claim Nr
                        !tmpFilePath = "#" & varFile & "#"      ' - Storing the file path as a hyperlink
                        .Update
                    End With
                    
                End If
                
                If intCounter = .FoundFiles.Count Then
'                   Update the last file's creation date:
                    Set rst2 = CurrentDb.OpenRecordset("tblOptions")
                        rst2.Edit
                        rst2!opLastICS = FileDateTime(varFile)
                        rst2.Update
                        rst2.Close
                End If
                
                fOK = acbUpdateMeter(intCounter / .FoundFiles.Count * 100)      ' - Update progress indicator
                
'               Handle user cancel:
                If Not fOK Then
                    MsgBox "Search interrupted!", vbOKOnly + vbCritical, _
                    "Claims table partially populated"
                    Exit For
                End If
                
            Next varFile
            rst.Close
            
            If fOK Then
                Call acbCloseMeter                                          ' - Finalize progress indicator
                MsgBox "Search completed!", vbOKOnly + vbInformation, _
                "Claims table populated"
            End If
        End If
    End With
    
'   Refresh the form to display updated claims:
    Me.Refresh

End Sub
Private Sub btSavePath_Click()
' Сохранить путь поиска

    Dim rst As Recordset
    
    Set rst = CurrentDb.OpenRecordset("tblOptions")
    rst.Edit
    rst!opPathICS = Me!txtPath
    rst.Update
    rst.Close
    
    MsgBox "Маршрут поиска " & Me!txtPath & " сохранен."

End Sub
Private Sub btAddNewRec_Click()

On Error GoTo HandleErr
    
    DoCmd.OpenQuery "qaAddNewICS"
    MsgBox "Новые данные добавлены в таблицу клеймов."
    
ExitHere:
    Exit Sub
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
         , "acbCloseMeter"
    End Select
    Resume ExitHere
    
End Sub
Private Sub btFrmCls_Click()
On Error GoTo Err_btFrmCls_Click


    DoCmd.Close

Exit_btFrmCls_Click:
    Exit Sub

Err_btFrmCls_Click:
    MsgBox Err.Description
    Resume Exit_btFrmCls_Click
    
End Sub
