Attribute VB_Name = "basWord"
' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All Rights Reserved.
    
Option Compare Database
Option Explicit

' Function: CreateTableFromRecordset
' Purpose: Generates a Word table from an ADODB.Recordset and inserts it into a specified range within a Word document.
'          Optionally, it includes field names as the table header.
'
' Parameters:
'   rngAny (Word.Range): The Word document range where the table will be created.
'   rstAny (ADODB.Recordset): The Recordset containing data to populate the table.
'   fIncludeFieldNames (Boolean, Optional): If True, adds field names as the table's first row header. Default is False.
'
' Returns:
'   Word.Table: The created Word table object containing the Recordset data.
'
' Usage:
'   Set myTable = CreateTableFromRecordset(myRange, myRecordset, True)
'
' Example:
'   Dim rst As ADODB.Recordset
'   Dim rng As Word.Range
'   Dim tbl As Word.Table
'
'   ' Assuming rst and rng are properly initialized
'   Set tbl = CreateTableFromRecordset(rng, rst, True)
'
'   ' Result: A Word table with Recordset data, including field names as a header.
'
' Detailed Steps:
'   1. Extracts Recordset data as a tab-delimited string using GetString.
'   2. Inserts the data string into the specified Word range.
'   3. Converts the inserted text into a table using ConvertToTable.
'   4. If fIncludeFieldNames is True, adds an additional row with field names at the top as a header.
'
' Assumptions:
'   - The Recordset should be open and contain data.
'   - Word and ADODB references are included in the project.
'
Function CreateTableFromRecordset( _
    rngAny As Word.Range, _
    rstAny As ADODB.Recordset, _
    Optional fIncludeFieldNames As Boolean = False) _
    As Word.Table

    Dim objTable As Word.Table                  ' Word table object to hold the Recordset data
    Dim fldAny As ADODB.Field                   ' Field object to access field names
    Dim varData As Variant                      ' Variable to store Recordset data as tab-delimited string
    Dim cField As Long                          ' Counter for field names in header row

    ' Extract data from the Recordset as a tab-delimited string
    varData = rstAny.GetString()

    ' Insert data as text and convert it to a Word table
    With rngAny
        .InsertAfter varData
        Set objTable = .ConvertToTable()

        ' Optionally add field names as the table header
        If fIncludeFieldNames Then
            With objTable
                ' Add a new row at the top and format it as a header
                .Rows.Add(.Rows(1)).HeadingFormat = True

                ' Populate header row with field names
                For Each fldAny In rstAny.Fields
                    cField = cField + 1
                    .Cell(1, cField).Range.Text = fldAny.Name
                Next fldAny
            End With
        End If
    End With

    ' Return the created table object
    Set CreateTableFromRecordset = objTable
End Function

' Subroutine: LetterCreate
' Purpose: Generates a technical campaign letter in Microsoft Word using data from an Access form and database.
'          Saves the letter to a specified location. This function also includes conditional formatting for
'          "Recall" campaigns and populates email recipients in "To" and "Copy" fields.
'
' Parameters:
'   frm (Form_frTA): Reference to the form containing data fields for the technical campaign, such as campaign
'                    number, description, model series, and email recipient details.
'
' Procedure:
'   1. Initializes a new Word application and loads a pre-defined Word template for the campaign letter.
'   2. Populates the Word template with data from the Access form using bookmarks, including conditional formatting
'      for Recall campaigns.
'   3. Executes a SQL query to retrieve a list of vehicles in stock related to the campaign.
'   4. Uses CreateTableFromRecordset to create a table in the Word document for vehicle details, with custom
'      formatting applied to the header row.
'   5. Saves the generated letter to a specified network directory.
'   6. Clears the Tag property in the form to reset the Director's email address in the Copy field.
'
' Returns:
'   None. The procedure performs actions directly in the Word document and saves the generated file to the network.
'
' Assumptions:
'   - The form (frm) is open and contains valid data.
'   - Microsoft Word and ADODB references are included in the project.
'   - The Word template ("tmpNewTA.dot") is located in the current project path.
'   - The network path specified in strWFS is accessible and writable.
'
' Usage:
'   Call LetterCreate(Form_frTA)
'
' Example:
'   Sub GenerateCampaignLetter()
'       Dim frm As Form_frTA
'       ' Initialize frm with required campaign data
'       Call LetterCreate(frm)
'   End Sub
'
' Notes:
'   - Ensure the form fields are correctly populated before calling this subroutine to avoid errors.
'   - This function is designed for campaign letters specific to technical actions and is not generalized.
'   - The file path "L:\TA\BMW\Projects\Letters\" should be accessible, and the user should have permissions to save files there.
'
Public Sub LetterCreate(frm As Form_frTA)

    ' Initialize a new Word application and template
    Dim objWord As Word.Application                       ' Word application object for document creation
    Dim strTAF As String                                  ' Campaign file name suffix
    Dim strTmpNm As String                                ' Template file name
    Dim strSQL As String                                  ' SQL query for vehicle selection
    Dim strWFS As String                                  ' File path and name for saving the Word document
    Dim rst As ADODB.Recordset                            ' Recordset to hold vehicle list data
    
    ' Generate campaign file name from form campaign number
    strTAF = "tad" & Mid(frm!aNo, 3, 6)                  ' Extract a unique part of campaign number for file naming
    strTmpNm = "\tmpNewTA.dot"                           ' Specify template file name for Word

    ' Launch Word application and add template document
    Set objWord = New Word.Application
    objWord.Documents.Add Application.CurrentProject.Path & strTmpNm
    objWord.Visible = True

    ' Populate template bookmarks with form data
    With objWord.ActiveDocument.Bookmarks
        ' Apply conditional formatting if campaign is a Recall
        If frm!aRecall Then
            .Item("bkmTAN").Range.Font.Color = wdColorDarkRed
            .Item("bkmTAR").Range.Font.Color = wdColorDarkRed
            .Item("bkmTAR").Range.Font.Bold = True
            .Item("bkmTAR").Range.Text = "Yes"
        Else
            .Item("bkmTAR").Range.Text = "No"
        End If
        ' Fill remaining bookmarks with form data
        .Item("bkmTAN").Range.Text = frm!aNo                ' Campaign number
        If frm!aRecall Then
            .Item("bkmTAD").Range.Font.Color = wdColorDarkRed
        End If
        .Item("bkmTAD").Range.Text = frm!‡Descr            ' Campaign description
        .Item("bkmTAS").Range.Text = Nz(frm!aSituation)    ' Situation description
        .Item("bkmTAE").Range.Text = Nz(frm!aEffect)       ' Consequences description
        .Item("bkmTAM").Range.Text = frm!aSer              ' Series / Model
        .Item("bkmTAF").Range.Text = strTAF                ' Campaign file name
        .Item("bkmTo").Range.Text = fnMailTo()             ' TO email addresses
        .Item("bkmCopy").Range.Text = fnMailCopy()         ' COPY email addresses
        .Item("bkmTAN2").Range.Text = frm!aNo & " "        ' Campaign number in the footer
    End With

    ' SQL query to retrieve vehicle details in stock related to the campaign
    strSQL = "SELECT sqTA_Vcl.avActNo AS Campaign, sqTA_Vcl.avVIN AS VIN, sqTA_Vcl.sVehicle AS Models, " & _
             "sqTA_Vcl.sCustName AS Customer, sqTA_Vcl.sAddress AS Address " & _
             "FROM sqTA_Vcl " & _
             "WHERE (((sqTA_Vcl.avActNo)='" & [Forms]![frTA]![aNo] & "') AND ((sqTA_Vcl.sCustName)='Stock'));"
    
    ' Create a table in the Word document for the vehicle details
    Set rst = New ADODB.Recordset
    rst.Open strSQL, CurrentProject.Connection
    If Not (rst.BOF And rst.EOF) Then
        ' Call CreateTableFromRecordset to generate the vehicle details table with field headers:
        With CreateTableFromRecordset( _
            objWord.ActiveDocument.Bookmarks("bkmTAT").Range, rst, True)
            .AutoFormat wdTableFormatProfessional          ' Apply professional auto-formatting
            .AutoFitBehavior wdAutoFitContent              ' Auto-fit content for better readability
            .Select
        End With
        ' Format column headers
        objWord.Selection.Paragraphs.Indent
        objWord.Selection.MoveUp Unit:=wdLine, Count:=1
        objWord.Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
        objWord.Selection.Font.Color = wdColorWhite
    End If

   ' Save the document to the current directory where the .mdb file is located
strWFS = CurrentProject.Path & "\lt" & Mid(frm!aNo, 3, 6) & ".doc"
    objWord.ActiveDocument.SaveAs FileName:=strWFS, FileFormat:=wdFormatDocument

    ' Release objects and clear the form's Tag property
    Set objWord = Nothing
    Set rst = Nothing
    frm.Tag = ""                                         ' Clear the Tag property for reuse

End Sub

