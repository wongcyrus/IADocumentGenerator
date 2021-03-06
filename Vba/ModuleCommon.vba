Attribute VB_Name = "ModuleCommon"
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

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
 
Sub CreateSheet(sheetname As String)
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = sheetname
End Sub

Sub DeleteSheet(strSheetName As String)
    If SheetExists(strSheetName) Then
        ' deletes a sheet named strSheetName in the active workbook
        Application.DisplayAlerts = False
        Sheets(strSheetName).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Sub ConvertText()
    For Each Cell In Selection
        Cell.value = Val(Cell.value)
    Next
    Selection.NumberFormat = "General"
End Sub

Sub CopyExcel(source As String, distination As String)
 'Create a copy first
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileCopy ActiveWorkbook.Path & "/" & source, ActiveWorkbook.Path & "/" & distination
End Sub


Sub CopyColumn(sourceName As String, sourceCol As String, sourceRow As Integer, _
distinationName As String, distinationCol As String, distinationRow As Integer, _
Optional sourceSheet As String = "", Optional distinationSheet As String = "", _
Optional numberOfSourceRow As Integer = -1, Optional pasteValue As Boolean = True)
    Windows(sourceName).Activate
    
    If sourceSheet <> "" Then
        Worksheets(sourceSheet).Activate
    End If
    
    If numberOfSourceRow = -1 Then
        Range(sourceCol & sourceRow).Select
        Range(Selection, Selection.End(xlDown)).Select
    Else
        Range(sourceCol & sourceRow & ":" & sourceCol & (sourceRow + numberOfSourceRow - 1)).Select
    End If
    

    Selection.Copy
    Windows(distinationName).Activate
    
    If distinationSheet <> "" Then
        Worksheets(distinationSheet).Activate
    End If
    Range(distinationCol & distinationRow).Select
    
    'ActiveSheet.Paste
    If pasteValue Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Else
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
        
End Sub

Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Public Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

'http://www.rondebruin.nl/win/s7/win001.htm
Sub SendEmail(toEmail As String, subject As String, body As String, filePathName As String)
'Create the mail
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
     
        On Error Resume Next
        With OutMail
            .To = toEmail
            .cc = ""
            .BCC = ""
            .subject = subject
            .body = body
            .Attachments.Add filePathName
            .Send   'or use .Display
        End With
        On Error GoTo 0
End Sub

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub
Public Function Zip_All_Files_in_Folder(FolderName As String) As String

    Dim FileNameZip As String
   
   
    FileNameZip = FolderName & ".zip"
    
    DeleteFile FileNameZip

    'Create empty Zip File
    NewZip (FileNameZip)
    
    Dim innerSource ' this should be a Variant
    Dim innerDest ' this should be a Variant
    innerSource = FolderName
    innerDest = FileNameZip
    Dim oApp: Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(innerDest).CopyHere oApp.Namespace(innerSource).items


    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace(innerDest).items.Count = _
       oApp.Namespace(innerSource).items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0

    Zip_All_Files_in_Folder = FileNameZip
End Function

Sub test()
Dim f As String
Dim ff As String
f = "C:\Users\cyrus\Documents\GitHub\IADocumentGenerator\Save\212212432312(nkl)"
ff = Zip_All_Files_in_Folder(f)
SendEmail "cywong@vtc.edu.hk", "Test", "demo", ff
End Sub

Function onlyDigits(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
End Function




