Attribute VB_Name = "��ǩ�������Է���"
Sub shl()
ActiveDocument.Bookmarks.Add "myplace", Selection.Range
ActiveDocument.Bookmarks.Add "myplace1", ActiveDocument.Paragraphs(3).Range '�����������β���ı�ǩ��
End Sub
Sub BookmarkItem()
If ActiveDocument.Bookmarks.Exists("myplace") = True Then
        ActiveDocument.Bookmarks.Item("myplace").Select
End If
End Sub '��ǩ��ѡ��λ
Sub shl1() '��ǩ������
MsgBox ActiveDocument.Bookmarks.Count
'ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation '�����ĵ��е�λ������wdSortByName������ǩ��������
'Dialogs(wdDialogInsertBookmark).Show '������ǩ�Ի���
ActiveDocument.Bookmarks.ShowHidden = True
For Each aBookmark In ActiveDocument.Bookmarks
    If Left(aBookmark.Name, 1) = "_" Then MsgBox aBookmark.Name '������ǩ��_��ͷ��
Next aBookmark
End Sub
