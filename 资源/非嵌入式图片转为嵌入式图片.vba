Sub 非嵌入式图片转为嵌入式图片()
	Dim oShape As Shape

	For Each oShape In ActiveDocument.Shapes
		oShape.ConvertToInlineShape
	Next
End Sub
