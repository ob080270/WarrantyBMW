Attribute VB_Name = "MyFunktions"
' ================================================================================================
' Module Name : MyFunctions
' Purpose     : Contains service procedures and functions (run manually, prefixed “ut”)
'               and automation functions (involved in programs)
'
' Functions Overview:
'   1. utCheckVIN(tableName As String, fieldName As String) As Void
'      - Validates VINs for length and invalid characters. Outputs results to the debug console.
'
'   2. utReplVIN(tableName As String, fieldName As String) As Void
'      - Replaces invalid characters in VINs with valid ones based on predefined rules.
'
'   3. utReplVIN2(tableName As String, fieldName As String) As Void
'      - An optimized version of utReplVIN using a dictionary for character replacements.
'
'   4. fDealerID(strDlrName As String) As Byte
'      - Maps dealer names to numeric codes for use in database queries.
'
'   5. PartOrdMail() As Void
'      - Generates a Word document for a parts order and sends an email with it attached.
'
'   6. fnDateInv(dtArg As Variant) As String
'      - Formats a date in "yymmdd" format for use as part order numbers.
'
'   7. fnUser(x As Byte) As String
'      - Returns user-specific information (e.g., name, email) based on context.
'
'   8. fnFinfWoman(strArg As String) As Boolean
'      - Determines if a client name likely belongs to a female client.
'
'   9. fnRef(strArg As String, sx As Boolean) As String
'      - Generates a personalized salutation for a client in a letter.
'
'  10. fnMailTo() As String
'      - Generates a list of email addresses for the "To" field of a letter.
'
'  11. fnMailCopy() As String
'      - Generates a list of email addresses for the "CC" field of a letter.
' ===================================================================================================
Option Compare Database
Option Explicit
' -------------------------------------------------------------------------------
' Function Name : CheckVIN
' Purpose       : Checks a specified field in a given table for VINs
'                 containing non-English characters, invalid symbols,
'                 or incorrect length.
' Arguments     :
'   - tableName (String): The name of the table containing the VIN field.
'   - fieldName (String): The name of the field to be checked for invalid VINs.
' Output        : Outputs invalid VINs to the debug console with a reason.
' Notes         : The function expects a 7-character VIN field and validates:
'                 - Each character for non-English letters or invalid symbols.
'                 - The length of the VIN (must be exactly 7 characters).
' -------------------------------------------------------------------------------
Public Function utCheckVIN(tableName As String, fieldName As String)
    Dim rstVIN As Recordset
    Dim strChar As String
    Dim i As Byte

On Error GoTo ErrorHandler
    Set rstVIN = CurrentDb.OpenRecordset(tableName)
    
    Do Until rstVIN.EOF
        If Len(rstVIN.Fields(fieldName).Value) <> 7 Then                                                            ' - Check if the VIN length is not 7
            Debug.Print "Invalid VIN (incorrect length): " & rstVIN.Fields(fieldName).Value
        Else
            For i = 1 To 7
                strChar = Mid$(rstVIN.Fields(fieldName).Value, i, 1)                                                ' - Extract each character from the VIN field
                If Asc(strChar) > 122 Or Asc(strChar) = 79 Then                                                     ' - Check for invalid characters (non-English or specific symbols)
                    Debug.Print "Invalid VIN (invalid character): " & rstVIN.Fields(fieldName).Value
                    Exit For                                                                                        ' - Exit the loop to avoid duplicate debug prints for the same VIN
                End If
            Next i
        End If
        rstVIN.MoveNext
    Loop
    
    Debug.Print "VIN validation process is complete."
    
Cleanup:
    ' Close the recordset and release resources
    If Not rstVIN Is Nothing Then
        rstVIN.Close
        Set rstVIN = Nothing
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    Resume Cleanup
End Function
' ------------------------------------------------------------------------------------------------
' Function Name : utReplVIN
' Purpose       : Corrects 7-character VIN values in the specified field of a table
'                 by replacing Cyrillic letters or invalid symbols with English letters.
' Arguments     :
'   - tableName (String): The name of the table containing the VIN field.
'   - fieldName (String): The name of the field to be corrected.
' Output        : Updates the VIN field directly in the table.
' -------------------------------------------------------------------------------------------------
Public Function utReplVIN(tableName As String, fieldName As String)
    Dim rstVIN As Recordset
    Dim strChar As String
    Dim i As Byte

    Set rstVIN = CurrentDb.OpenRecordset(tableName, , dbSeeChanges)
    
    Do Until rstVIN.EOF
        If Len(rstVIN.Fields(fieldName).Value) <> 7 Then                                                            ' - Check if the VIN length is not 7
            Debug.Print "Invalid VIN (incorrect length): " & rstVIN.Fields(fieldName).Value
        Else
            For i = 1 To 7
                strChar = Mid$(rstVIN.Fields(fieldName).Value, i, 1)
                If Asc(strChar) > 122 Or Asc(strChar) = 79 Then
                    rstVIN.Edit
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "À", "A")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Â", "B")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Ñ", "C")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Å", "E")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Í", "H")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Ê", "K")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Ì", "M")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Î", "0")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "O", "0")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Ð", "P")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Ò", "T")
                    rstVIN.Fields(fieldName).Value = Replace(rstVIN.Fields(fieldName).Value, "Õ", "X")
                    rstVIN.Update
                End If
            Next i
        End If
        rstVIN.MoveNext
    Loop
    
    rstVIN.Close
    Set rstVIN = Nothing

    Debug.Print "VIN validation process is complete."

End Function

' -------------------------------------------------------------------
' Function Name : utReplVIN2 (2-nd realization of function utReplVIN)
' Purpose       : Corrects 7-character VIN values in the specified field of a table
'                 by replacing Cyrillic letters or invalid symbols with English letters.
' Arguments     :
'   - tableName (String): The name of the table containing the VIN field.
'   - fieldName (String): The name of the field to be corrected.
' Output        : Updates the VIN field directly in the table and logs invalid or corrected VINs.
' Notes         : Ensures that the VIN length is exactly 7 and corrects characters only if needed.
' -------------------------------------------------------------------
Public Function utReplVIN2(tableName As String, fieldName As String)
    Dim rstVIN As Recordset                                                                     ' - Recordset (tableName)
    Dim strVIN As String                                                                        ' - Current VIN value
    Dim originalVIN As String                                                                   ' - To store the original VIN value before processing it
    Dim charToReplace As String                                                                 ' - processed character
    Dim corrections As Object                                                                   ' - Scripting.Dictionary
    Dim i As Byte                                                                               ' - cycle counter

    Set corrections = CreateObject("Scripting.Dictionary")                                      ' - Initialize the mapping of Cyrillic to English characters
        corrections.Add "À", "A"
        corrections.Add "Â", "B"
        corrections.Add "Ñ", "C"
        corrections.Add "Å", "E"
        corrections.Add "Í", "H"
        corrections.Add "Ê", "K"
        corrections.Add "Ì", "M"
        corrections.Add "Î", "0"
        corrections.Add "O", "0"
        corrections.Add "Ð", "P"
        corrections.Add "Ò", "T"
        corrections.Add "Õ", "X"

On Error GoTo ErrorHandler

    Set rstVIN = CurrentDb.OpenRecordset(tableName, , dbSeeChanges)

    Do Until rstVIN.EOF                                                                         ' - Iterate through each record in the table
        strVIN = rstVIN.Fields(fieldName).Value
        
        If Len(strVIN) <> 7 Then                                                                ' - Check VIN length
            Debug.Print "Invalid VIN (incorrect length): " & strVIN
        Else
            originalVIN = strVIN

            For i = 1 To Len(strVIN)                                                            ' - Correct invalid characters in VIN:
                charToReplace = Mid$(strVIN, i, 1)
                If corrections.Exists(charToReplace) Then
                    strVIN = Replace(strVIN, charToReplace, corrections(charToReplace))
                End If
            Next i

            If strVIN <> originalVIN Then                                                       ' - Update the VIN field if changes were made
                rstVIN.Edit
                rstVIN.Fields(fieldName).Value = strVIN
                rstVIN.Update
                Debug.Print "Corrected VIN: " & originalVIN & " -> " & strVIN
            End If
        End If

        rstVIN.MoveNext
    Loop

    Debug.Print "VIN correction process is complete."

Cleanup:
' Release resources:
    If Not rstVIN Is Nothing Then
        rstVIN.Close
        Set rstVIN = Nothing
    End If
    If Not corrections Is Nothing Then
        Set corrections = Nothing
    End If
    Exit Function

ErrorHandler:
    Debug.Print "Error: " & Err.Description
    Resume Cleanup
End Function
' ---------------------------------------------------------------------------------------------------------
' Function Name : fDealerID
' Purpose       : Returns the dealer code based on the dealer name (strDlrName).
'                 The function is used in queries (NewVhclAdd) for adding data to the "lkSold" table.
'                 Ensures that dealer names are mapped to correct IDs.
' Arguments     :
'   - strDlrName (String): Input string representing the dealer name.
' Returns       : Byte - The dealer code. If the dealer name is invalid, returns 99.
' Notes         :
'   - This function was originally used in the "NewVhclAdd" query.
'   - The "NewVhclAdd" query is no longer functional in the isolated version of the project,
'     as it requires dependencies on Sales Department tables, which are not included.
' Error Handling: If the dealer name does not match any predefined cases,
'                 an error message is displayed, and the dealer is marked with code 99.
' -------------------------------------------------------------------------------------------------------
Public Function fDealerID(strDlrName As String) As Byte
' Map the dealer name to the corresponding dealer ID:
    Select Case strDlrName
        Case "AWT Bavaria", "BMW Diplomatic", "AWT Bavaria_Kiev", "BMW Diplomatic_Kiev"
            fDealerID = 0
        Case "Impuls", "Impuls_Donetsk"
            fDealerID = 1
        Case "Bavaria Motors", "Bavaria Motors_Kharkov"
            fDealerID = 3
        Case "AWT Bavaria Dnepropetrovsk"
            fDealerID = 4
        Case "Emerald Motors", "Emerald Motors_Odessa"
            fDealerID = 2
        Case "Avtodel", "Avtodel_Simpferopol"
            fDealerID = 6
        Case "Khrystyna", "Khrystyna_Lviv"
            fDealerID = 5
        Case "Autoservice Aljans", "Autoservice Aljans_Kremenchung"
            fDealerID = 7
        Case "Konstructor", "Konstructor_Nikolaev"
            fDealerID = 10
        Case "Bavaria Jug", "Bavaria Jug_Kherson"
            fDealerID = 9
        Case "AWT Zaporozh'je"
            fDealerID = 8
        Case Else                                                                                   ' - Handle invalid dealer names
            MsgBox "Unknown dealer: " & strDlrName                                                  ' - Notify the administrator of the invalid dealer name
            fDealerID = 99                                                                          ' - Mark the record for manual review
    End Select

End Function

Public Sub PartOrdMail()
' Purpose: Create an email to the Parts Department with an order for the Technical Campaign.
'          This procedure automates the generation of an order document, email creation, and task creation in Outlook.
'
' Trigger: Called from the procedure cmdMoveToOrd_Click in the module of the form frTA.
'
' Workflow:
' 1. Exports query data to an RTF file.
' 2. Creates a Word document with the order details.
' 3. Generates an email in Outlook with the order attached.
' 4. Creates a task in Outlook to follow up on the order.

    Dim doc As Word.Document
    Dim wrdApp As Word.Application
    Dim strDataSource, CrLf As String
    Dim strTA As String                                                                 '- Campaign number
    Dim strMailBody As String                                                           '- e-mail-body, containing the order
    Dim strSFP As String                                                                '- SFP - SavedFilePath - order file saving path
    
    ' Delete the rtf file, if it already exists
    strDataSource = "L:\TA\BMW\OrdPart\ltPartOrd.doc"
    Kill strDataSource
    
    ' Export the data to rtf format
    DoCmd.OutputTo acOutputQuery, "qsLtPartOrd", acFormatRTF, strDataSource, False
    
    ' Create a letter in Word:
    Set wrdApp = New Word.Application
    Set doc = wrdApp.Documents.Add(strDataSource)
    CrLf = chr(13) & chr(10)
    strTA = Forms!frTA!aNo                                                              '- Campaign number
    strSFP = "L:\TA\BMW\OrdPart\op" & strTA & ".doc"                                    '- File storage path
    
    With wrdApp.Selection
        .WholeStory
        .Cut
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .PasteAndFormat (wdPasteDefault)
        .HomeKey Unit:=wdStory
        .Text = "Spare parts for the Campaign " & strTA
        .WholeStory
        .Font.Name = "Arial"
    End With
    
    ' Construct email body:
    strMailBody = "Dear colleagues," & CrLf & CrLf
    strMailBody = strMailBody & "Please order/reserve the spare parts sent as an attachment " & strTA & CrLf
    strMailBody = strMailBody & "for the Technical Campaign." & CrLf & CrLf
    strMailBody = strMailBody & "-----------------------" & CrLf
    strMailBody = strMailBody & "Best regards," & CrLf & fnUser(1)
                
    ' Display document for debugging
    'wrdApp.Visible = True
    doc.SaveAs strSFP                                                                   '- Save the Word document
    Forms!frTA!sfOrders.Form.hypOrdRef = "#" & strSFP & "#"                             '- Save the hyperlink to the order file
    doc.Close wdDoNotSaveChanges                                                        '- Close file
    
    ' Transfer data to Outlook:
    Dim objOutlook As Outlook.Application
    Dim objMailItem As MailItem
    Dim objAtt As Attachment
    
    Set objOutlook = New Outlook.Application
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    
    With objMailItem
    ' e-mail recipient creation:
        .To = "SparePartsRecipient1@bmw.ua"
        .CC = "SparePartsRecipient2@bmw.ua"
        .Subject = "Parts Order: Technical campaign " & strTA & " Car"
        .Body = strMailBody
        Set objAtt = .Attachments.Add(strSFP)                                           '- Attach the parts order file
        .Display
    End With
    Set objMailItem = Nothing
    
    ' Create an Outlook task to control the ordering of spare parts:
    Dim objTaskItem As TaskItem
    Set objTaskItem = objOutlook.CreateItem(olTaskItem)
    
    With objTaskItem
        .Subject = "Check the order of spare parts for Campaign " & strTA & " Car"
        .Body = strMailBody
        .Categories = "Order control"
        .DueDate = Date + 2
        .ReminderSet = True
        .ReminderTime = DateAdd("d", 2, Date)
        .Save
    End With
        
    ' Clean up object:
    Set objTaskItem = Nothing
    Set objOutlook = Nothing
    Set doc = Nothing
    Set wrdApp = Nothing

End Sub

Public Function fnDateInv(dtArg As Variant) As String
' Purpose   : Returns a date in the format "yymmdd".
' Trigger   : Called from the query prgPartOrdAdd to generate the order number based on the date.
' Arguments :
'   - dtArg : A Variant that represents the input date.
' Returns   :
'   - A string in the format "yymmdd", or a blank string if the input date is null.

    Dim strArg As String
    
    If IsNull(dtArg) Then
        fnDateInv = String(6, " ")
    Else
        strArg = CStr(dtArg)
        fnDateInv = Right(strArg, 2) & Mid(strArg, 4, 2) & Left(strArg, 2)
    End If

End Function

Public Function fnUser(x As Byte) As String
' Purpose: Returns user-specific information based on the input argument.
' Trigger: Called by the PartOrdMail function to create the sender's signature for the order email.
' Arguments:
'   - x (Byte): Determines the type of information to return:
'       1 - User signature
'       2 - User email
' Returns:
'   - A string containing the requested information.

    Select Case x
        Case 1
            If CurrentUser() = "Admin" Then
                fnUser = "Name1 Surname1"
            Else
                fnUser = "Name2 Surname2"
            End If
        Case 2
            If CurrentUser() = "Admin" Then
                fnUser = "Admin@bmw.ua"
            Else
                fnUser = "User2@bmw.ua"
            End If
        End Select
End Function
' -----------------------------------------------------------------------------------------------------
' Function Name : fnFinfWoman
' Purpose       : Determines if the client's name field (sCustName) likely belongs to a female client.
'                 The function analyzes the second space in the name field and checks if the character
'                 immediately before it matches specific criteria for female names.
'
' Arguments     :
'   - strArg (String): The client name field (expected format: "LastName FirstName MiddleName").
'
' Returns       :
'   - Boolean:
'       TRUE  - If the client's name likely belongs to a female client.
'       FALSE - Otherwise, or if the name format does not match expectations.
'
' Notes         :
'   - This function is designed to set the sWoman field in the lkSold table, which is used for customizing
'     communication (e.g., letters) based on the client's gender.
'   - If the name field contains fewer than two spaces, the function defaults to FALSE.
' -----------------------------------------------------------------------------------------------------
Public Function fnFinfWoman(strArg As String) As Boolean

    Dim tmp As Long
    Dim chr As String
    
    tmp = InStr(strArg, " ")                            ' Looking for the first space (presumably between last name and first name)
    
    If tmp > 0 Then
        tmp = InStr(tmp + 1, strArg, " ") - 1           ' Looking for the 2nd space (presumably between first name and middle name) and indent 1 character to the left
        If tmp = -1 Then                                ' If the 2nd space is not found - return to the 1st found space
            tmp = InStr(strArg, " ") - 1
            fnFinfWoman = False
            Exit Function
        End If
        chr = Mid$(strArg, tmp, 1)
    
        If chr = "à" Or chr = "ÿ" Then
            fnFinfWoman = True
        Else
            fnFinfWoman = False
        End If
        
        Else
        fnFinfWoman = False                             ' If the client field contains 1 word, then leave FALSE
    End If

End Function
' ---------------------------------------------------------------------------------------
' Function Name : fnRef
' Purpose       : Generates a personalized salutation for a client in a letter
'                 based on their name format and gender.
' Description   :
'   - If the name field contains the full name (Last Name, First Name, Middle Name),
'     the last name is truncated, and a salutation like "Dear Ivan Vasilyevich /
'     Dear Natalia Ivanovna" is generated.
'   - If the name field contains initials instead of full first name and patronymic,
'     a salutation like "Dear Mr. Ivanov / Mrs. Ivanova" is generated.
'
' Arguments     :
'   - strArg (String): The client's name field (e.g., "Surname FirstName MiddleName").
'   - sx (Boolean): Indicates the gender of the client.
'       - TRUE: Female client.
'       - FALSE: Male client.
'
' Returns       :
'   - String: A personalized salutation based on the client's name and gender.
' --------------------------------------------------------------------------------------
Public Function fnRef(strArg As String, sx As Boolean) As String

    Dim tmp As String   ' - Stores the extracted part of the name (First Name and Middle Name)
    Dim tmp2 As String  ' - Stores the extracted last name
    
    tmp = Mid$(LTrim(strArg), InStr(strArg, " ") + 1)                       ' - Extract the part of the name after the first space
    
    'Check if the extracted part is shorter than 6 characters (likely initials):
    If Len(tmp) < 6 Then
        tmp2 = Mid$(LTrim(strArg), 1, InStr(strArg, " ") - 1)               ' - Extract the last name (before the first space)
        
        'Generate a salutation based on the client's gender:
        If sx Then
            fnRef = "Dear Mrs. " & tmp2                                     ' - Salutation for a female client
        Else
            fnRef = "Dear Mr. " & tmp2                                      ' - Salutation for a male client
        End If
    Else
        fnRef = "Dear " & tmp                                               ' - If the extracted part is full, use it directly in the salutation
    End If

End Function
' -------------------------------------------------------------------------------------------------
' Function Name : fnMailTo
' Purpose       : Generates a list of email addresses to be placed in the "To" field of an email.
' Trigger       : Called from the LetterCreate function in the basWord module.
'
' Workflow      :
'   1. Retrieves dealers' email addresses from the query "qDlrMailTo".
'   2. Checks if there are cars in stock from the query "qMailStock".
'   3. Combines the email addresses into a single string separated by "; "
'
' Returns       :
'   - String: A concatenated list of email addresses.
' --------------------------------------------------------------------------------------------------
Public Function fnMailTo() As String

    Dim rst As Recordset        ' - Recordset to hold query results
    Dim strMailTo As String     ' - String to store the list of email addresses
    Dim qryDef As QueryDef      ' - Query definition object
    
    ' Step 1: Retrieve email addresses of dealers from the query "qDlrMailTo":
    Set qryDef = CurrentDb.QueryDefs("qDlrMailTo")                                  ' - Pass the Technical Campaign number as a parameter
        qryDef("prmaNo") = Forms!frTA!aNo
    Set rst = qryDef.OpenRecordset()
    
    strMailTo = ""                                                                  ' - Initialize the email string
    
    ' Loop through the recordset and collect email addresses:
    rst.MoveFirst
    While Not rst.EOF
        strMailTo = strMailTo & rst!dlrTo & "; "                                    ' - Append each email address with a separator
        rst.MoveNext
    Wend
    
    ' Clean up objects after processing the query
    Set rst = Nothing
    Set qryDef = Nothing
    
    ' Step 2: Check if there are cars in stock for the Technical Campaign
    Set qryDef = CurrentDb.QueryDefs("qMailStock")
        qryDef("prmaNo") = Forms!frTA!aNo                                           ' - Pass the Technical Campaign number as a parameter
    Set rst = qryDef.OpenRecordset()
    
    If rst.RecordCount > 0 Then
        rst.MoveLast                                                                ' - Move to the last record to ensure all records are loaded
        strMailTo = strMailTo & "StockResponsiblePerson@bmw.ua"                     ' - Add a stock-related email address
    End If
    
    ' Clean up objects
    Set rst = Nothing
    Set qryDef = Nothing
    
    ' Assign the concatenated email list to the function's return value
    fnMailTo = strMailTo

End Function

Public Function fnMailCopy()
' Get a string that contains a list of e-mails to be inserted into the letter in the "CC" field

    fnMailCopy = "DealerDepartmentManager@bmw.ua; ServiceDepartmentManager@bmw.ua; technologist@bmw.ua; WarrantyAdmin@bmw.ua"

End Function

