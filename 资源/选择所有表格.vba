Sub 选择所有表格()

    Dim tempTable As Table
    Application.ScreenUpdating = False
    If ActiveDocument.ProtectionType = wdAllowOnlyFormFields Then
        MsgBox "文档已保护，此时不能选中多个表格！"
        Exit Sub
    End If
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    For Each tempTable In ActiveDocument.Tables
        tempTable.Range.Editors.Add wdEditorEveryone
    Next
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    Application.ScreenUpdating = True

End Sub
