Attribute VB_Name = "����Ƶ�ʲ�ѯ"
Sub ��Ƶ���ж�()
Dim rng As Range, avar As Variable, str As String, mystr As String
MsgBox "���ĵ������Ϳ����ڴ��Լ�word������ֵ�����Ĳ������ܻ���Ҫһ��ʱ�䣬" & vbCrLf & _
"Ҳ�����ֿ����ڴ治����������������Ҫ����word�Ա������һ�εĹ�����", vbOKOnly + vbExclamation, "warnning"
varclear
bkclear
For Each rng In ActiveDocument.Words
    ActiveDocument.UndoClear '���word�г���
    If rng.Characters.Count > 1 Then
        If Asc(rng.Characters(1)) < -2050 And Asc(rng.Characters(1)) > -20319 Then '�����ڱ�������-2050��-20319
            If ActiveDocument.Bookmarks.Exists(rng.Text) = True Then
                On Error Resume Next '���ִ���ִ����һ�䣬�ĵ�������Ӳ����ظ�
                ActiveDocument.Variables.Add rng.Text
                If Err.Number <> 0 Then 'һ��� on error resume next ���ã���������err��number��ֵ������0��
                    Err.Clear
                    ActiveDocument.Variables(rng.Text).Value = ActiveDocument.Variables(rng.Text).Value + 1
                Else
                    ActiveDocument.Variables(rng.Text).Value = 2 '�״�д�������value=2����Ϊ�Ѿ�����һ����ǩ��
                End If
            Else
                ActiveDocument.Bookmarks.Add rng.Text 'ֻ������ǩû�й涨������ǩ�ĵص㡣
            End If
        End If
    End If
Next rng
Application.ScreenUpdating = False
With Selection
    .EndKey wdStory '�Ƶ��ĵ������
    For Each avar In ActiveDocument.Variables
        str = """" & avar.Name & "����Ƶ�Σ�" & vbTab & avar.Value & vbCrLf
        mystr = mystr & str
    Next avar
    .InsertAfter mystr
    mystr = ""
    .Sort , "�� 1", wdSortFieldNumeric, wdSortOrderDescending 'ѧϰsort����
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
