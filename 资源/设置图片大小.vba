Sub setpicsize() '设置图片大小
Dim n '图片个数
On Error Resume Next '忽略错误
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes类型图片
    ActiveDocument.InlineShapes(n).Height = 400 '设置图片高度为 400px
    ActiveDocument.InlineShapes(n).Width = 300 '设置图片宽度 300px
Next n
For n = 1 To ActiveDocument.Shapes.Count 'Shapes类型图片
    ActiveDocument.Shapes(n).Height = 400 '设置图片高度为 400px
    ActiveDocument.Shapes(n).Width = 300 '设置图片宽度 300px
Next n
End Sub
