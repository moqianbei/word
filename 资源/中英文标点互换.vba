Sub ToggleInterpunction() '中英文标点互换
'
'
'
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte

ChineseInterpunction = Array("、", "。", "，", ";", "：", "？", "！", "……", "—", "～", "（", "）", "《", "》","￥")'定义一个中文标点的数组对象
EnglishInterpunction = Array(",", ".", ",", ";", ":", "?", "!", "…", "-", "~", "(", ")", "<", ">","$")'定义一个英文标点的数组对象

'提示用户交互的MSGBOX对话框
msgResult = MsgBox("您正在使用中英标点互换功能，按Y将中文标点转为英文标点,按N将英文标点转为中文标点!退出请点击右上角关闭按钮", vbYesNoCancel)

Select Case msgResult

	Case vbCancel
		Exit Sub '如果用户选择了取消按钮,则退出程序运行
	Case vbYes '如果用户选择了YES,则将中文标点转换为英文标点
		myArray1 = ChineseInterpunction
		myArray2 = EnglishInterpunction
		strFind = "“(*)”"
		strRep = """\1"""
	Case vbNo '如果用户选择了NO,则将英文标点转换为中文标点
		myArray1 = EnglishInterpunction
		myArray2 = ChineseInterpunction
		strFind = """(*)"""
		strRep = "“\1”"
End Select

Application.ScreenUpdating = False '关闭屏幕更新

For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
	With ActiveDocument.Content.Find
		.ClearFormatting '不限定查找格式
		.MatchWildcards = False '不使用通配符
		'查找相应的英文标点,替换为对应的中文标点
		.Execute findtext:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
	End With
Next

With ActiveDocument.Content.Find
	.ClearFormatting '不限定查找格式
	.MatchWildcards = True '使用通配符
	.Execute findtext:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
End With

Application.ScreenUpdating = True '恢复屏幕更新
End Sub
