Attribute VB_Name = "��ǩ�����Է���"
Sub shl()
ActiveDocument.Bookmarks("myplace2").Copy "myplace1" '����ǩ1��λ�ã���Ϊ��ǩ2.
Dim docNew As Document
Dim tableNew As Table
Dim rangeCell As Range

Set docNew = Documents.Add
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 5)
Set rangeCell = tableNew.Cell(3, 5).Range

'rangeCell.InsertAfter "Cell(3,5)"
rangeCell.Text = "123" '����д������
docNew.Bookmarks.Add Name:="BKMK_Cell35", Range:=rangeCell
MsgBox docNew.Bookmarks(1).Column 'ȷ����ǩ�Ǳ����С�

If ActiveDocument.Bookmarks.Exists("temp") = True Then
    If ActiveDocument.Bookmarks("temp").Empty = True Then _
    MsgBox "The Temp bookmark is empty" '�鿴����ǩ�ǲ�����κ��ı�
End If
Set Book1 = ActiveDocument.Bookmarks("myplace")
Set Book2 = ActiveDocument.Bookmarks("myplace3")
If Book2.End > Book1.Start Then Book1.Select 'start ��end �����������У�����ѡ���Ľ�������ʼλ�á�
'��ǩ��name���Բ��ڴ˽��в�����
If ActiveDocument.Bookmarks.Exists("temp") = True Then 'storytype���ԣ��ı������������Ĳ��ֻ���������ע��ҳüҳ�ŵȡ�
    Set myBookmark = ActiveDocument.Bookmarks("temp")
    If myBookmark.StoryType = wdMainTextStory _
        Then myBookmark.Select 'wdmaintextstory��ʾ���Ĳ��֡�
End If
'��ǩ��range���Բ��ڽ�һ���������൱�ھ����޴����ڡ�
End Sub
