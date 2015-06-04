Attribute VB_Name = "ModuleGenerateIADocument"
'The MIT License (MIT)
'
'Copyright (c) 2014 <Cyrus Wong>
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

Option Explicit

Dim wd As New Word.Application
Dim PersonCell As Range
'create copy of Word in memory
Dim doc As Word.Document
Dim docCopySource As Word.Document


Dim PersonRange As Range
Dim FilePathOpen As String
Dim FilePathSave As String
Dim FilePathSaveAllStudents As String

Sub InitPathVariable()
    FilePathOpen = Application.ActiveWorkbook.Path & "\"
    FilePathSave = Application.ActiveWorkbook.Path & "\Save\"
    FilePathSaveAllStudents = FilePathSave & "\AllStudents\"
End Sub

Sub CreateWordDocuments()

    InitPathVariable
    
    makeSaveDir FilePathSave
    makeSaveDir FilePathSaveAllStudents
    wd.Visible = True
   ' wdCopySource.Visible = True
    
    'create a reference to all the people
    SelectStudentsRange
        
    ProcessCreate "Statement of Understanding (Organization).docx"
    ProcessCreate "Final Report (CompanyOrganization Mentor).docx"
    ProcessCreate "Evaluation Report (Student).docx"
    ProcessCreate "Industrial Attachment Certificate (Template).docx"
    ProcessCreate "Industrial Attachment Form (Organization).docx"
    ProcessCreate "Insurance Coverage for Industrial Attachment Students.docx"
    ProcessCreate "Monthly Report (Student).docx"
    ProcessCreate "Student Information.docx"
    ProcessCreate "Visiting Report (IA Supervisor) (Optional).docx", True
           
    ClearClipBoard
    wd.Quit
    Set wd = Nothing
    
    MsgBox "Created files in " & FilePathSave
End Sub

Public Sub ClearClipBoard()
   'object to use the clipboard
  Dim oData   As New DataObject
  'Clear
  oData.SetText Text:=Empty
  'take in the clipboard to empty it
  oData.PutInClipboard
End Sub

Sub CopyCell(BookMarkName As String, RowOffset As Integer)
    'copy each cell to relevant Word bookmark
    If doc.Bookmarks.Exists(BookMarkName) = True Then
        wd.Selection.GoTo What:=wdGoToBookmark, Name:=BookMarkName
        wd.Selection.InsertAfter PersonCell.Offset(RowOffset, 0).value
        wd.Selection.GoTo What:=wdGoToBookmark, Name:=BookMarkName
        wd.Selection.Delete
    End If
End Sub

Sub CopyImageFromFile(BookMarkName As String, Optional imageFileName As String)
    Dim GraphImage As String
    If imageFileName = "" Then
        GraphImage = FilePathOpen & "SignAndChop/" & BookMarkName & ".jpg"
    Else
        GraphImage = FilePathOpen & "SignAndChop/" & imageFileName & ".jpg"
    End If
    
    Dim hasImage As Boolean
    hasImage = Dir(GraphImage) <> ""
    
    If hasImage And doc.Bookmarks.Exists(BookMarkName) = True Then
        wd.Activate
        wd.Selection.GoTo What:=wdGoToBookmark, Name:=BookMarkName
        wd.Selection.InlineShapes.AddPicture filename:=GraphImage, LinkToFile:=False, SaveWithDocument:=True
    End If
  
End Sub

Sub CopyImageFromWord(BookMarkName As String)
              
   ' Set docSource = wdSource.Documents.Open(PersonCell.Offset(-1, 0).value)
      ' Get the collection of all content controls with this tag.
    Dim sourceWordPathName As String
    'The first Row
    sourceWordPathName = PersonCell.Offset(-1, 0).value
    Set docCopySource = wd.Documents.Open(sourceWordPathName)
    
    docCopySource.ActiveWindow.View.ReadingLayout = False
    
    Dim cc As ContentControl
    Dim docCCs As ContentControls
    ' Get the collection of all content controls with this tag.
    Set docCCs = docCopySource.SelectContentControlsByTag(BookMarkName)
    
    Dim hasImage As Boolean
    hasImage = False
    
    If docCCs.Count <> 0 Then
        For Each cc In docCCs
            If cc.Tag = BookMarkName And cc.Range.InlineShapes.Item(1).Height <> 85 Then
            '85 size is the trick of no pic
                wd.Activate
                cc.Range.Select
                wd.Selection.Copy
                hasImage = True
            End If
        Next
    End If
    docCopySource.Close
    Set docCopySource = Nothing
    'SaveClipboardBMP
    If hasImage And doc.Bookmarks.Exists(BookMarkName) = True Then
        wd.Activate
        wd.Selection.GoTo What:=wdGoToBookmark, Name:=BookMarkName
        wd.Selection.Paste
    End If
End Sub


Sub DoFieldCopy()
    CopyCell "studentID", 0
    CopyCell "studentNameEng", 1
    CopyCell "studentNameChi", 2
    CopyCell "department", 3
    CopyCell "hkidNo", 4
    CopyCell "hkidNoLast", 5
    CopyCell "campus", 6
    CopyCell "courseCode", 7
    CopyCell "year", 8
    CopyCell "courseTitle", 9
    CopyCell "contactPhoneNo", 10
    CopyCell "email", 11
    CopyCell "organizationNameEng", 12
    CopyCell "organizationNameChi", 13
    CopyCell "organizationAddEng", 14
    CopyCell "organizationAddChi", 15
    CopyCell "orgMentorNameChi", 16
    CopyCell "orgMentorNameEng", 17
    CopyCell "orgMentorJobTitle", 18
    CopyCell "orgMentorDepartment", 19
    CopyCell "orgMentorPhoneNo", 20
    CopyCell "orgMentorFaxNo", 21
    CopyCell "orgMentorEmail", 22
    CopyCell "contactPersonName", 23
    CopyCell "contactPersonPhone", 24
    CopyCell "jobDurationWeek", 25
    CopyCell "jobStartDateDMY", 26
    CopyCell "jobFinishDateDMY", 27
    CopyCell "jobTitle", 28
    CopyCell "jobDeptment", 29
    CopyCell "workingDaysPerWeek", 30
    CopyCell "workingHoursForm", 31
    CopyCell "workingHoursFormAMPM", 32
    CopyCell "workingHoursTo", 33
    CopyCell "workingHoursToAMPM", 34
    CopyCell "basis", 35
    CopyCell "attachmentOutside", 36
    CopyCell "shiftDuty", 37
    CopyCell "iveCooNameEng", 38
    CopyCell "iveCooNameChi", 39
    CopyCell "iveCooDepCampus", 40
    CopyCell "iveCooRank", 41
    CopyCell "iveCooPhone", 42
    CopyCell "iveCooEmail", 43
    CopyCell "iveCooCampusNameNAdd", 44
    CopyCell "iveMentorNameEng", 45
    CopyCell "iveMentorNameChi", 46
    CopyCell "iveMentorDepCampus", 47
    CopyCell "iveMentorRank", 48
    CopyCell "iveMentorPhone", 49
    CopyCell "iveMentorEmail", 50
    CopyCell "iveMtCampusNameNAdd", 51
    CopyCell "emergencyNameEng", 52
    CopyCell "emergencyNameChi", 53
    CopyCell "emergencyRelation", 54
    CopyCell "emergencyPhone", 55
    CopyCell "HeadOfDeptName", 56
    CopyCell "workingHoursTotal", 58
    'CopyCell "docDate", 57
    CopyCell "iveMentorFax", 57
    'CopyCell "organMentorName", 58
    
    
    

    Dim i As Integer
    For i = 1 To 5
        CopyCell "studentID" & i, 0
        CopyCell "studentNameEng" & i, 1
        CopyCell "department" & i, 3
        CopyCell "campus" & i, 6
        CopyCell "organizationNameEng" & i, 12
        CopyCell "iveMentorNameEng" & i, 45
        CopyCell "jobStartDateDMY" & i, 26
        CopyCell "workingHoursTotal" & i, 58
        CopyCell "emergencyPhone" & i, 55
    Next i
End Sub

Sub DoCopy()
 Dim mentorCNA As String
    mentorCNA = PersonCell.Offset(50, 0).value
    Dim mentorCNAArray() As String
    mentorCNAArray = Split(mentorCNA, "@", 2)
    'Merge Image
    CopyImageFromFile "hodSign"
    CopyImageFromFile "DeptChop"
    CopyImageFromFile "mentorSign", mentorCNAArray(0)
    
    If PersonCell.Offset(-1, 0).value <> "" Then
        CopyImageFromWord "StudentSignature"
        CopyImageFromWord "StudentPhoto"
        CopyImageFromWord "CompanyChop"
        CopyImageFromWord "CompanyMentorSign"
    End If
    'AllPictSize
    'go to each bookmark and type in details
    DoFieldCopy
End Sub


Sub ProcessCreate(filename As String, Optional copyToAllStudents As Boolean = False)
    'for each person in list
    For Each PersonCell In PersonRange
    
        'open a document in Word
        Set doc = wd.Documents.Open(FilePathOpen & "/WordTemplate/" & filename)
    
        'go to each bookmark and type in details
        DoCopy
        
        Dim orgName As String
        orgName = Trim(PersonCell.Offset(12, 0).value)
        
        Dim studentId As String
        studentId = Trim(PersonCell.value)
              
        'create folder
        makeSaveDir FilePathSave & studentId & "(" & orgName & ")\"
        'save and close this document
        doc.SaveAs2 FilePathSave & studentId & "(" & orgName & ")\" & studentId & " " & filename
        
        'Save one more copy for batch printing
        If copyToAllStudents Then
            doc.SaveAs2 FilePathSaveAllStudents & "\" & studentId & " " & filename
        End If
        doc.Close
        Set doc = Nothing
    Next PersonCell
End Sub

Sub ZipAndEmail()
    InitPathVariable
    
    SelectStudentsRange
    'for each person in list
    For Each PersonCell In PersonRange
            
        Dim orgName As String
        orgName = Trim(PersonCell.Offset(12, 0).value)
        'create folder
        Dim studentId As String
        studentId = Trim(PersonCell.value)
        
        Dim documentFolder As String
        documentFolder = FilePathSave & studentId & "(" & orgName & ")"
       
        SendEmail studentId & "@stu.vtc.edu.hk", _
        "IA Documment Set", "Please check the content, and upload it to MyPortal, if there is no error!", _
        Zip_All_Files_in_Folder(documentFolder)
    Next PersonCell
End Sub



Sub makeSaveDir(Path As String)
    If Len(Dir(Path, vbDirectory)) = 0 Then
        MkDir (Path)
    End If
End Sub

Sub SelectStudentsRange()
    Worksheets("Student Data").Activate
    Range("B2").Select
    If IsEmpty(Cells(2, 3)) Then
        Set PersonRange = Range(ActiveCell, ActiveCell)
    Else
        Set PersonRange = Range(Cells(2, 2), Cells(2, 2).End(xlToRight))
    End If
End Sub

 
Sub WordExtract()
     '==
     'Add Word object reference library.
     'Tools->References - Check the Microsoft Word Object Libary box
    Dim wbWorkBook As Workbook
    Dim wsWorkSheet As Worksheet
    Dim oWord As Word.Application
    Dim WordWasNotRunning As Boolean
    Dim oDoc As Word.Document
    Dim varFileName As Variant
    Dim intHeaderRow As Integer, intNumberOfField As Integer, i As Integer
    Dim intFileProcessed As Integer
    Dim strPath As String, strDocFiles As String, strDisplayText As String
    Dim strFullName As String, strFieldName As String, strFieldValue As String
    Dim strFieldType As String
    Dim strCaption As String, strTempFieldValue As String
    Dim wsMessage As Object
    Dim xRow As Long
     
    Set wsMessage = CreateObject("WScript.Shell")
    Set wbWorkBook = ActiveWorkbook
    DeleteSheet "Imported Data"
    CreateSheet "Imported Data"
    Set wsWorkSheet = wbWorkBook.Worksheets("Imported Data")
    Range("A1").Select
     
     'For FYI Info....
    wsMessage.Popup " This Utility Only Works with *.Docx Files in the Import Folder, Press OK To Continue.... ", 5, _
    "..... Information .....", 4096
     
     'Get existing instance of Word if it's open; otherwise create a new one
    On Error Resume Next
    Set oWord = GetObject(, "Word.Application")
    If Not Err Then
         'Close the word instance if open
        oWord.Quit
    End If
     
    Set oWord = New Word.Application
    On Error GoTo Err_Handler
     
     'Prompt user for the directory where all the word document are located.
   ' With Application.FileDialog(msoFileDialogFolderPicker)
   '     .Show
   '     If .SelectedItems.Count = 1 Then
   '         strPath = .SelectedItems(1)
   '     End If
   ' End With
   ' If strPath = Empty Then
   '     wsMessage.Popup "No folder Selected", 5, "..... Error .....", 4096
   '     Exit Sub
   ' End If
    
    strPath = ActiveWorkbook.Path & "/Import"
     
     'Get the last row
    xRow = wsWorkSheet.Range("A" & Rows.Count).End(xlUp).Row
     'Append new Row.....
    xRow = xRow + 2
     
     'Keep track number of word document processed
    intFileProcessed = 0
     
     'Retreive list of all the word doc files in the given directory
     'For now this only works with *.Doc files only.  Not the *.Docx, as we only have word office 2003 installed.
     'It can be change easily to accomodiate the new docx format.
    strDocFiles = Dir(strPath & "\*.Doc*")
     
     ' Loop through all the word document in this directory, retrieve the info and insert it into the excel sheet.
    Do While strDocFiles <> ""
        intFileProcessed = intFileProcessed + 1
         
         'Prompt to select single file
         'varFileName = Application.GetOpenFilename("Word Files (*.doc; *.docx), *.doc; *.docx")
         'varFileName = Application.GetOpenFilename("Word Files (*.doc; *.docx), *.doc; *.docx", , , 1)
         'Prompt to select a directory (More than one word file)
        Set oDoc = oWord.Documents.Open(strPath & "\" & strDocFiles, Visible:=False)
         
        With wsWorkSheet
             'Get the Total Number of user fillable TextBox field in this Document
            intNumberOfField = oDoc.FormFields.Count
             
             'Get the Full Name of the current word document
            strFullName = oDoc.FullName
             
             'Display processing info in the popup window.....
            strDisplayText = "Processing..... " & strFullName & " - Total Field Count: " & Trim(str(intNumberOfField))
            'wsMessage.Popup strDisplayText, 1, "Processing", 4096
             
             'If this is the first file being processed, retrieve the header too..
             'At this point, haven't figure out how to retrieve the textbox title/label, so just retrieve the actual object name
            If intFileProcessed = 1 Then
                 'Loop through all the fillable fields
                For i = 1 To intNumberOfField
                    strCaption = ""
                     'Retrieve field object Name and insert/update into Excel cell
                    strFieldName = oDoc.FormFields(i).Name
                     'The following commented out line: Trying to get the field caption, didn't work....
                     'strCaption = oDoc.Bookmarks(strFieldName).Range.Text
                     '.Cells(xRow, i + 1) = strFieldName & " - " & strCaption
                     
                    wsWorkSheet.Activate
                    .Cells(xRow, i + 1) = strFieldName
                Next i
                 
                 'Save the header row # for setting them to Bold after all the files is run.
                intHeaderRow = xRow
                 'Add date and time stamp to first column in the header row
                .Cells(xRow, 1).Select
                Selection.value = Now()
                Selection.HorizontalAlignment = xlLeft
            End If
             
             'Append new Row.....
            xRow = xRow + 1
             
             'Update the full name of the word doc in the first column of the current row
            wsWorkSheet.Activate
            .Cells(xRow, 1) = strFullName
             
             'Retrieve the fillable field result for the current document.
            For i = 1 To intNumberOfField
                 'Retrieve Full Filename, field type, field name and fieldvalue
                strFieldType = oDoc.FormFields(i).Type
                strFieldName = oDoc.FormFields(i).Name
                strFieldValue = oDoc.FormFields(i).Result
            
                 
                 'Display processing info in the popup window.....
                 'strDisplayText = "Processing..... " & strFullName & " - Total Field Count: " & Trim(Str(intNumberOfField)) _
                 '& ", _Current Field: " & Trim(Str(i)) & " - " & strFieldName & ",  Field Value: " & strFieldValue
                 'wsMessage.Popup strDisplayText, 1, "Processing", 4096
                 
                 ' Type "wdFieldFormCheckBox" = 71, if it's a check box, the value store is either "1" or "0" for true or false
                 ' the following converts "1" to "True" and "0" to "False" for easier understanding by the users.
                If strFieldType = 71 Then
                    Select Case strFieldValue
                    Case "0"
                        strTempFieldValue = "No"
                    Case "1"
                        strTempFieldValue = "Yes"
                    End Select
                    wsWorkSheet.Activate
                    .Cells(xRow, i + 1) = strTempFieldValue
                Else
                    wsWorkSheet.Activate
                    .Cells(xRow, i + 1) = "" & strFieldValue
                End If
                 'Debug.Print strDisplayText
                 
            Next i
        End With
         
        oDoc.Close savechanges:=wdDoNotSaveChanges
        Set oDoc = Nothing
         
         'Get the next doc
        strDocFiles = Dir
    Loop
     
    oWord.Quit
     'Make sure you release object references.
    Set oWord = Nothing
    Set oDoc = Nothing
     
     'Set the header row to Bold font only
    wsWorkSheet.Activate
    wsWorkSheet.Cells.Select
    Selection.Font.Bold = False
    Rows(intHeaderRow & ":" & intHeaderRow).Select
    Selection.Font.Bold = True
     
     'Select the whole Excel sheet and expand all the columns
    wsWorkSheet.Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    ActiveWorkbook.Save
    
    Sheets("Imported Data").Select
    Range("A4:BM" & intFileProcessed + 3).Select
    'Range(Selection, Selection.End(xlDown)).Select
   'Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Student Data").Select
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    ActiveSheet.CircleInvalid
    
    Exit Sub
     
Err_Handler:
    Select Case Err
    Case -2147022986, 429
        Set oWord = CreateObject("Word.Application")
        Resume Next
    Case 5121, 5174
        MsgBox "You must select a valid Word document. " & "No data imported.", vbOKOnly, "Document Not Found"
    Case 5941
        MsgBox Err.Description
         '    'MsgBox "The document you selected does not " _
         '        & "contain the required form fields. " _
         '        & "No data imported.", vbOKOnly, _
         '        "Fields Not Found"
    Case Else
        MsgBox Err & ": " & Err.Description
        oWord.Quit
    End Select
End Sub





