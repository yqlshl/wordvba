Attribute VB_Name = "��Ƶ���ж�"
Sub ��Ƶ���ж�()
Dim rng As Range, avar As Variable, str As String, mystr As String
MsgBox "���ĵ������Ϳ����ڴ��Լ�word������ֵ�����Ĳ������ܻ���Ҫһ��ʱ�䣬" & vbCrLf & _
"Ҳ�����ֿ����ڴ治����������������Ҫ����word�Ա������һ�εĹ�����", vbOKOnly + vbExclamation, "warnning"
varclear
bkclear
For Each rng In ActiveDocument.Words
    ActiveDocument.UndoClear '������ʹ�ó�����ť
    If rng.Characters.Count > 1 Then
        If Asc(rng.Characters(1)) < -2050 And Asc(rng.Characters(1)) > -20319 Then '�жϺ��ֵķ�Χ��
            If ActiveDocument.Bookmarks.Exists(rng.Text) = True Then
                On Error Resume Next
                ActiveDocument.Variables.Add rng.Text
                If Err.Number <> 0 Then '����ĵ�������������
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
        str = """" & avar.Name & "����Ƶ�Σ�" & vbTab & avar.Value & vbCrLf
        mystr = mystr & str
    Next avar
    .InsertAfter mystr
    mystr = ""
    .Sort , "�� 1", wdSortFieldNumeric, wdSortOrderDescending
End With
ActiveDocument.UndoClear
varclear
bkclear
Application.ScreenUpdating = True
End Sub
Sub varclear() '����ĵ��������ĵ�����
Dim v As Variable
For Each v In ActiveDocument.Variables
    v.Delete
Next v

End Sub
Sub bkclear() 'ɾ���ĵ�����ǩ��������ʹ�ó�����ť
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
    dDate1 = VBA.Format(DateAdd("d", -1, Date), "yyyy��mm��dd��")
    dDate2 = VBA.Format(DateAdd("d", -1, Date), "yyyy-mm-dd")
    With oDoc
        For Each oV In .Variables
            oV.Delete
        Next
        '�����ΪPreDate1��ֵΪ����dDate1���ĵ�����
        .Variables.Add "PreDate1", dDate1
         '�����ΪPreDate2��ֵΪ����dDate2���ĵ�����
        .Variables.Add "PreDate1", dDate2
    End With
End Sub
Rem �ĵ�����������ڹ�������vba��������ĵ���������word����docvariable����в��롣
