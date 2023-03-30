Attribute VB_Name = "书签的属性方法"
Sub shl()
ActiveDocument.Bookmarks("myplace2").Copy "myplace1" '讲书签1的位置，变为书签2.
Dim docNew As Document
Dim tableNew As Table
Dim rangeCell As Range

Set docNew = Documents.Add
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 5)
Set rangeCell = tableNew.Cell(3, 5).Range

'rangeCell.InsertAfter "Cell(3,5)"
rangeCell.Text = "123" '两种写法都行
docNew.Bookmarks.Add Name:="BKMK_Cell35", Range:=rangeCell
MsgBox docNew.Bookmarks(1).Column '确认书签是表格的列。

If ActiveDocument.Bookmarks.Exists("temp") = True Then
    If ActiveDocument.Bookmarks("temp").Empty = True Then _
    MsgBox "The Temp bookmark is empty" '查看空书签是不标记任何文本
End If
Set Book1 = ActiveDocument.Bookmarks("myplace")
Set Book2 = ActiveDocument.Bookmarks("myplace3")
If Book2.End > Book1.Start Then Book1.Select 'start 和end 常规性理解就行，就是选定的结束和起始位置。
'书签的name属性不在此进行阐述。
If ActiveDocument.Bookmarks.Exists("temp") = True Then 'storytype属性，文本的类型是正文部分还是其他脚注，页眉页脚等。
    Set myBookmark = ActiveDocument.Bookmarks("temp")
    If myBookmark.StoryType = wdMainTextStory _
        Then myBookmark.Select 'wdmaintextstory表示正文部分。
End If
'书签中range属性不在进一步描述，相当于精灵无处不在。
End Sub
