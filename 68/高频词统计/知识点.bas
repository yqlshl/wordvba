Attribute VB_Name = "֪ʶ��"
Sub shl1()
'variable��value�����Լ����塣
ActiveDocument.Variables.Add Name:="Value1", Value:="1"
MsgBox ActiveDocument.Variables("Value1") + 3
For Each myvar In ActiveDocument.Variables
    MsgBox "Name =" & myvar.Name & vbCr & "Value = " & myvar.Value
    myvar.Delete
Next myvar
End Sub
