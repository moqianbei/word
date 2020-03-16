Sub 改变字符串数组颜色()
'
'
    Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 12'字号
        .Bold = True'粗体
        .Color = wdColorAutomatic'自动颜色
    End With
  Dim arr()
   arr = Array("<int>", "<float>", "<char>", "<return>", "<double>", "<if>", "<else>", "<for>", "<switch>", "<case>", "<break>", "<do>", "<while>", "<long>", "<typedef>", "<short>", "<unsigned>", "<auto>", "<continue>", "<signed>", "<struct>", "<void>", "<default>", "<goto>", "<sizeof>", "<volatile>", "<static>", "<enum>", "<register>", "<extern>", "<union>", "<const>")'在""中输入想要查找的词语或表达式
   For i = 0 To 31'31为上面的个数-1
     With Selection.Find
        .Text = arr(i)
        .Replacement.Text = "^&amp;"'替换内容为查找框的内容
        .Forward = True'向上匹配
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True'匹配大小写
        .MatchWholeWord = False'全字匹配
        .MatchByte = False
        .MatchWildcards = True'使用通配符
        .MatchSoundsLike = False'同音（英文）
        .MatchAllWordForms = False'查找单词的所有形式（英文）
    End With
Selection.Find.Execute Replace:=wdReplaceAll'全部替换
    Next i
End Sub
