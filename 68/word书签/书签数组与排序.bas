Attribute VB_Name = "��ǩ����������"
Option Compare Text '���ı���ʽ�Ƚϲ����ִ�Сд
Sub shl() 'vbCrLf��ʾ�س�+���С�
Dim cel As Cell, str() As String, mystr As String, intcount As Integer, bkmark As Bookmark
Dim first As Integer, last As Integer, i As Integer, j As Integer, temp As String
intcount = 0
With Selection
    For Each cel In .Tables(1).Range.Cells
        mystr = ActiveDocument.Range(cel.Range.Start, cel.Range.End - 1)
        If mystr Like "#*" = True Then
            MsgBox "�����ݲ��ܱ�����ʶ���������������κ�������ʽ���֣�" & vbCrLf & """" & mystr & """"
            Exit Sub
        Else
            ActiveDocument.Bookmarks.Add mystr '�ѱ����ÿ�����ݲ���Ϊ��ǩ��
        End If
    Next cel
    ReDim str(0 To ActiveDocument.Bookmarks.Count - 1) 'ֱ�Ӷ�������Ϊ8����preserve str��0 to n��Ҳ���ԡ�
    ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation
    For Each bkmark In ActiveDocument.Bookmarks
        str(intcount) = bkmark.Name
        intcount = intcount + 1
    Next bkmark
    first = LBound(str)
    last = UBound(str)
    For i = first To last - 1 'ð�����򷨣���������ܾ��䣬�����ַ�������ĸ��chr�����С�����ֲ��У���Ϊ������chr�е�˳���ǰ���ƴ������ġ�
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
