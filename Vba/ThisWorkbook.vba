VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Function VBAIsTrusted() As Boolean
    VBAIsTrusted = False
    On Error GoTo Label1
    a1 = ThisWorkbook.VBProject.VBComponents.Count
    VBAIsTrusted = True
Label1:

End Function

Private Sub Workbook_Open()
    If VBAIsTrusted Then
        ImportCodeModules
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If VBAIsTrusted Then
        SaveCodeModules
    End If
End Sub


'http://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
Sub SaveCodeModules()
    'This code Exports all VBA modules
    Dim i%, sName$
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                .VBComponents(i%).Export Application.ActiveWorkbook.Path & "\Vba\" & sName$ & ".vba"
            End If
        Next i
    End With
End Sub

Sub ImportCodeModules()
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            ModuleName = .VBComponents(i%).CodeModule.Name
            If ModuleName <> "VersionControl" Then
                'MsgBox ModuleName
                If Right(ModuleName, 6) = "Module" Then
                    .VBComponents.Remove .VBComponents(ModuleName)
                    .VBComponents.Import Application.ActiveWorkbook.Path & "\Vba\" & ModuleName & ".vba"
               End If
            End If
        Next i
    End With
End Sub
