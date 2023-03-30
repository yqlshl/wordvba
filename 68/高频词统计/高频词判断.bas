Attribute VB_Name = "高频词判断"
Sub 高频词判断()
Dim rng As Range, avar As Variable, str As String, mystr As String
MsgBox "受文档字数和可用内存以及word自身限值，您的操作可能会需要一段时间，" & vbCrLf & _
"也许会出现可能内存不足的情况，您可能需要重启word以便接着下一次的工作！", vbOKOnly + vbExclamation, "warnning"
varclear
bkclear
For Each rng In ActiveDocument.Words
    ActiveDocument.UndoClear '不允许使用撤销按钮
    If rng.Characters.Count > 1 Then
        If Asc(rng.Characters(1)) < -2050 And Asc(rng.Characters(1)) > -20319 Then '判断汉字的范围内
            If ActiveDocument.Bookmarks.Exists(rng.Text) = True Then
                On Error Resume Next
                ActiveDocument.Variables.Add rng.Text
                If Err.Number <> 0 Then '添加文档变量发生错误
                    Err.Clear
                    ActiveDocument.Variables(rng.Text).Value = ActiveDocument.Variables(rng.Text).Value + 1
                Else
                    ActiveDocument.Variables(rng.Text).Value = 2
                End If
            Else
                ActiveDocument.Bookmarks.Add rng.Text
            End If
        End If
    End If
Next rng
Application.ScreenUpdating = False
With Selection
    .EndKey wdStory
    For Each avar In ActiveDocument.Variables
        str = """" & avar.Name & "出现频次：" & vbTab & avar.Value & vbCrLf
        mystr = mystr & str
    Next avar
    .InsertAfter mystr
    mystr = ""
    .Sort , "域 1", wdSortFieldNumeric, wdSortOrderDescending
End With
ActiveDocument.UndoClear
varclear
bkclear
Application.ScreenUpdating = True
End Sub
Sub varclear() '清空文档中所有文档变量
Dim v As Variable
For Each v In ActiveDocument.Variables
    v.Delete
Next v

End Sub
Sub bkclear() '删除文档中书签并不允许使用撤销按钮
Dim bk As Bookmark
ActiveDocument.UndoClear
For Each bk In ActiveDocument.Bookmarks
    bk.Delete
    ActiveDocument.UndoClear
Next bk
End Sub
Private Sub Document_Open()
    Dim oV As Variable
    Dim oDoc As Document
    Set oDoc = Word.ActiveDocument
    dDate1 = VBA.Format(DateAdd("d", -1, Date), "yyyy年mm月dd日")
    dDate2 = VBA.Format(DateAdd("d", -1, Date), "yyyy-mm-dd")
    With oDoc
        For Each oV In .Variables
            oV.Delete
        Next
        '添加名为PreDate1，值为变量dDate1的文档变量
        .Variables.Add "PreDate1", dDate1
         '添加名为PreDate2，值为变量dDate2的文档变量
        .Variables.Add "PreDate1", dDate2
    End With
End Sub
Rem 文档变量和域存在关联，用vba代码插入文档变量，在word中用docvariable域进行插入。
