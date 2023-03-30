Attribute VB_Name = "书签集合属性方法"
Sub shl()
ActiveDocument.Bookmarks.Add "myplace", Selection.Range
ActiveDocument.Bookmarks.Add "myplace1", ActiveDocument.Paragraphs(3).Range '在整个第三段插入的标签。
End Sub
Sub BookmarkItem()
If ActiveDocument.Bookmarks.Exists("myplace") = True Then
        ActiveDocument.Bookmarks.Item("myplace").Select
End If
End Sub '标签的选择定位
Sub shl1() '标签的属性
MsgBox ActiveDocument.Bookmarks.Count
'ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation '按照文档中的位置排序。wdSortByName按照书签名称排序。
'Dialogs(wdDialogInsertBookmark).Show '弹出书签对话框。
ActiveDocument.Bookmarks.ShowHidden = True
For Each aBookmark In ActiveDocument.Bookmarks
    If Left(aBookmark.Name, 1) = "_" Then MsgBox aBookmark.Name '隐藏书签是_开头。
Next aBookmark
End Sub
