Attribute VB_Name = "ModuleBatchUpload"
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

Public sourceName As String
Public distinationName As String
Dim lastRow As Integer

Sub CopyCourseField()
'
' CopyCourseField Macro
'

'
    Windows("IA.xlsm").Activate
    Range("B1").Select
    Selection.Copy
    Windows(distinationName).Activate
    Range("B3").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("IA.xlsm").Activate
    Range("B2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(distinationName).Activate
    Range("B4").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("IA.xlsm").Activate
    Range("B3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(distinationName).Activate
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("IA.xlsm").Activate
    Range("B4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(distinationName).Activate
    Range("E4").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Save
End Sub




Sub GetLastRowForDisintaion()
    Windows(sourceName).Activate
    Range("E2").Select
    lastRow = Selection.End(xlDown).Row + 5
End Sub


Sub CopyBlock()
'
' CopyBlock Macro
'
    'Handle First part
    Windows(sourceName).Activate
    Range("I2:AE" & lastRow).Select
    Selection.Copy
    Windows(distinationName).Activate
    Range("c7").Select
   ' ActiveSheet.Paste
     
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
  'Handle Second part
    Windows(sourceName).Activate
    Range("AF2:AL" & lastRow).Select
    Selection.Copy
    Windows(distinationName).Activate
    Range("AA7").Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub FixAttachmentCategory()
'
' FixAttachmentCategory Macro
'

'
    FillColWith "I", "A"
End Sub

Sub FillColWith(distinationCol As String, value As String)

    Windows(distinationName).Activate
    Range(distinationCol & "7").Select
    ActiveCell.FormulaR1C1 = value
    Range(distinationCol & "7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
End Sub


Sub TruncateValue(distinationCol As String, codeLenght As Integer)
'
' FixAttachmentCategory Macro
'

'
    Windows(distinationName).Activate
     Range(distinationCol & "7:" & distinationCol & lastRow).Select
    
    For Each rCell In Selection
        rCell.value = Mid(rCell.value, 1, codeLenght)
    Next rCell
    
End Sub

Sub ReFormatCode(distinationCol As String)
'
' FixAttachmentCategory Macro
'

'
    Windows(distinationName).Activate
    Range(distinationCol & "7:" & distinationCol & lastRow).Select
    
    For Each rCell In Selection
    
        If rCell.value = "-" Then
            rCell.value = ""
        End If
    
        If Not IsFormatHyphenCorrect(rCell.value) Then
            rCell.value = Replace(rCell.value, "-", " - ")
        End If
        
    Next rCell
End Sub

Function IsFormatHyphenCorrect(sentence) As Boolean

    Dim regEx As New VBScript_RegExp_55.RegExp
    Dim matches, s
    regEx.Pattern = "\w[ ]-[ ]\w" 'Match abc - def
    regEx.IgnoreCase = True 'True to ignore case
    regEx.Global = True 'True matches all occurances, False matches the first occurance
    If regEx.test(sentence) Then
        IsFormatHyphenCorrect = True
    Else
        IsFormatHyphenCorrect = False
    End If

End Function


Sub ConvertVerifiedHoursToNumber()
    Windows(distinationName).Activate
    Range("AG7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ConvertText
End Sub



Sub CreatePivotTableForVerifiedHours()

    Windows(distinationName).Activate
    CreateSheet "VerifiedHours"
    
    Worksheets("Student Attachment Records").Select
    Range("A6", Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Dim myCells As Range
    Set myCells = Selection
    
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        myCells.Address, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="VerifiedHours!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("VerifiedHours").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Student ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Verified Hours"), "Count of Verified Hours", _
        xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Count of Verified Hours")
        .Caption = "Sum of Verified Hours"
        .Function = xlSum
    End With
    Range("B5").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Student ID").AutoSort _
        xlAscending, "Sum of Verified Hours", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(1), 1

    Sheets("VerifiedHours").Select
    'Move it to last! My portal read according to Order!
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
End Sub


Sub GenerateStudentPortfolioIndustrialAttachmentBatch()

    If EndsWith(Range("B7").value, ".xls") Then
        sourceName = Range("B7").value
    Else
        sourceName = Range("B7").value & ".xls"
    End If

    distinationName = Range("B8").value & ".xls"
    Dim resultFileName As String
    resultFileName = Range("B8").value & "(Filled).xls"
    'Create a copy first
    CopyExcel "/ExcelTemplate/" & distinationName, resultFileName
    distinationName = resultFileName
    
    Dim needFixAttachmentCategory As Boolean
    needFixAttachmentCategory = CBool(Range("B5").value)
    
    Dim codeOfIndustry As String
    codeOfIndustry = Range("B6").value
    
    Application.Workbooks.Open Application.ActiveWorkbook.Path & "\" & sourceName
    Application.Workbooks.Open Application.ActiveWorkbook.Path & "\" & distinationName
    
    Application.DisplayAlerts = False
    
    Windows(sourceName).Activate
    Sheets("Student Attachment Records").Select
    
    GetLastRowForDisintaion
    CopyCourseField
    CopyColumn sourceName, "D", 2, distinationName, "A", 7
    CopyColumn sourceName, "E", 2, distinationName, "B", 7
    
    'CopyColumn sourceName, "AM", 2, distinationName, "Z", 7
        
    CopyBlock
                
    'CopyColumn sourceName, "N", 2, distinationName, "AG", 7
    
    If needFixAttachmentCategory Then
        FixAttachmentCategory
    Else
        TruncateValue "I", 1
    End If
       
    ReFormatCode "J"
    ReFormatCode "N"
    ReFormatCode "P"
    ReFormatCode "Q"
    ReFormatCode "V"
    ReFormatCode "X"
    ReFormatCode "Y"
    
    TruncateValue "F", 50
    TruncateValue "M", 120
    TruncateValue "W", 120
    TruncateValue "AB", 120
       
    
    'ConvertVerifiedHoursToNumber
    'CreatePivotTableForVerifiedHours
    
    Windows(distinationName).Activate
    ActiveWorkbook.Sheets("Student Attachment Records").Activate
    ActiveSheet.CircleInvalid

    Application.DisplayAlerts = True
    'ActiveWorkbook.Save
    
End Sub



