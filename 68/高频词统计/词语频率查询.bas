Attribute VB_Name = "词语频率查询"
Sub 高频词判断()
Dim rng As Range, avar As Variable, str As String, mystr As String
MsgBox "受文档字数和可用内存以及word自身限值，您的操作可能会需要一段时间，" & vbCrLf & _
"也许会出现可能内存不足的情况，您可能需要重启word以便接着下一次的工作！", vbOKOnly + vbExclamation, "warnning"
varclear
bkclear
For Each rng In ActiveDocument.Words
    ActiveDocument.UndoClear '清空word中撤销
    If rng.Characters.Count > 1 Then
        If Asc(rng.Characters(1)) < -2050 And Asc(rng.Characters(1)) > -20319 Then '汉字在编码中是-2050至-20319
            If ActiveDocument.Bookmarks.Exists(rng.Text) = True Then
                On Error Resume Next '出现错误执行下一句，文档变量添加不能重复
                ActiveDocument.Variables.Add rng.Text
                If Err.Number <> 0 Then '一般和 on error resume next 联用，发生错误，err的number的值不在是0。
                    Err.Clear
                    ActiveDocument.Variables(rng.Text).Value = ActiveDocument.Variables(rng.Text).Value + 1
                Else
                    ActiveDocument.Variables(rng.Text).Value = 2 '首次写入变量的value=2，因为已经存在一个书签。
                End If
            Else
                ActiveDocument.Bookmarks.Add rng.Text '只增加书签没有规定插入书签的地点。
            End If
        End If
    End If
Next rng
Application.ScreenUpdating = False
With Selection
    .EndKey wdStory '移到文档的最后。
    For Each avar In ActiveDocument.Variables
        str = """" & avar.Name & "出现频次：" & vbTab & avar.Value & vbCrLf
        mystr = mystr & str
    Next avar
    .InsertAfter mystr
    mystr = ""
    .Sort , "域 1", wdSortFieldNumeric, wdSortOrderDescending '学习sort方法
End With
ActiveDocument.UndoClear
varclear
bkclear
Application.ScreenUpdating = True
End Sub
Sub varclear()
Dim v As Variable
For Each v In ActiveDocument.Variables
    v.Delete
Next v
End Sub
Sub bkclear()
Dim bk As Bookmark
ActiveDocument.UndoClear
For Each bk In ActiveDocument.Bookmarks
    bk.Delete
    ActiveDocument.UndoClear
Next bk
End Sub
