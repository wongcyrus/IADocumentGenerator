Attribute VB_Name = "ModuleIARecord"
Public sourceName As String
Public distinationName As String

Public dataSheetName As String
Sub GenerateIARecord()
 
    dataSheetName = "Student Data Row"
    DeleteSheet dataSheetName
    CreateSheet dataSheetName
   
    ThisWorkbook.Worksheets("Student Data").Activate
    
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy

    ThisWorkbook.Worksheets(dataSheetName).Activate
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    
    sourceName = "IA.xlsm"
    distinationName = "Student IA Record(Filled).xls"
    'Create a copy and open it.
    CopyExcel "/ExcelTemplate/Student IA Record.xls", distinationName
    Application.Workbooks.Open Application.ActiveWorkbook.Path & "\" & distinationName
    
    Dim sourceSheet As String
    Dim distinationSheet As String

    sourceSheet = "Student Data Row"
    distinationSheet = "Student Attachment Records"

    Workbooks(sourceName).Worksheets(sourceSheet).Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Dim numberOfStudent As Integer
    numberOfStudent = Selection.Count
    'MsgBox numberOfStudent
     
    Dim i As Integer
    Dim distinationCol As String
    i = 2
    Workbooks(sourceName).Worksheets("ColumnMapping").Activate
    Do While Cells(i, 1).value <> ""
        Dim sourceCol As String
        
        Workbooks(sourceName).Worksheets("ColumnMapping").Activate
        sourceCol = Cells(i, 1).value
        distinationCol = Cells(i, 2).value
        
        CopyColumn sourceName, sourceCol, 2, distinationName, distinationCol, 2, _
        sourceSheet, distinationSheet, numberOfStudent, False
                
        i = i + 1
        Workbooks(sourceName).Worksheets("ColumnMapping").Activate
    Loop
    
    i = 1
    'Workbooks(sourceName).Worksheets("Data").Activate
    Do While Workbooks(sourceName).Worksheets("Data").Cells(i, 2).value <> ""
        Dim targetCol As String
            
        targetCol = Workbooks(sourceName).Worksheets("Data").Cells(i, 3).value
                
        If targetCol <> "" Then
            Workbooks(distinationName).Worksheets(distinationSheet).Activate
            Range(targetCol & "2:" & targetCol & (1 + numberOfStudent)).value = Workbooks(sourceName).Worksheets("Data").Cells(i, 2).value
        End If
                
        i = i + 1
        Workbooks(sourceName).Worksheets("Data").Activate
    Loop
    
    Dim orginalValue As String
    orginalValue = Workbooks(sourceName).Worksheets("Data").Range("B7").value
    Workbooks(sourceName).Worksheets("Data").Range("B7").value = distinationName
    Workbooks(distinationName).Close savechanges:=True
    
   
    GenerateStudentPortfolioIndustrialAttachmentBatch
        
    
    Workbooks(sourceName).Worksheets("Data").Range("B7").value = orginalValue
End Sub
