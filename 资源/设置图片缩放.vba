Sub 设置图片缩放()
Dim n '图片个数
Dim picwidth
Dim picheight
On Error Resume Next '忽略错误
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes类型图片
    picheight = ActiveDocument.InlineShapes(n).Height
    picwidth = ActiveDocunent.InlineShapes(n).Width
    ActiveDocument.InlineShapes(n).Height = picheight * 1.1 '设置高度为1.1倍
    ActiveDocument.InlineShapes(n).Width = picwidth * 1.1 '设置宽度为1.1倍
Next n
For n = 1 To ActiveDocument.Shapes.Count 'Shapes类型图片
    picheight = ActiveDocument.Shapes(n).Height
    picwidth = ActiveDocument.Shapes(n).Width
    ActiveDocument.Shapes(n).Height = picheight * 1.1 '设置高度为1.1倍
    ActiveDocument.Shapes(n).Width = picwidth * 1.1 '设置宽度为1.1倍
Next n
End Sub
