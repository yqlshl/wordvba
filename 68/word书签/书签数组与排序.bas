Attribute VB_Name = "书签数组与排序"
Option Compare Text '以文本方式比较不区分大小写
Sub shl() 'vbCrLf表示回车+换行。
Dim cel As Cell, str() As String, mystr As String, intcount As Integer, bkmark As Bookmark
Dim first As Integer, last As Integer, i As Integer, j As Integer, temp As String
intcount = 0
With Selection
    For Each cel In .Tables(1).Range.Cells
        mystr = ActiveDocument.Range(cel.Range.Start, cel.Range.End - 1)
        If mystr Like "#*" = True Then
            MsgBox "此数据不能被程序识别，请勿在其首以任何数字形式出现！" & vbCrLf & """" & mystr & """"
            Exit Sub
        Else
            ActiveDocument.Bookmarks.Add mystr '把表格中每个内容插入为标签。
        End If
    Next cel
    ReDim str(0 To ActiveDocument.Bookmarks.Count - 1) '直接定义数组为8个，preserve str（0 to n）也可以。
    ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation
    For Each bkmark In ActiveDocument.Bookmarks
        str(intcount) = bkmark.Name
        intcount = intcount + 1
    Next bkmark
    first = LBound(str)
    last = UBound(str)
    For i = first To last - 1 '冒泡排序法，这个方法很经典，根据字符串首字母的chr排序大小。汉字不行，因为汉字在chr中的顺序不是按照拼音排序的。
        For j = i + 1 To last
            If str(i) > str(j) Then
                temp = str(j)
                str(j) = str(i)
                str(i) = temp
            End If
        Next j
    Next i
    mystr = ""
    temp = ""
    .EndKey wdStory
    .InsertAfter Chr(13)
    For x = first To last
        temp = str(x) & Chr(13)
        mystr = mystr & temp
    Next x
    .InsertAfter mystr
End With
End Sub
