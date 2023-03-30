Attribute VB_Name = "知识点"
Sub shl1()
'variable的value可以自己定义。
ActiveDocument.Variables.Add Name:="Value1", Value:="1"
MsgBox ActiveDocument.Variables("Value1") + 3
For Each myvar In ActiveDocument.Variables
    MsgBox "Name =" & myvar.Name & vbCr & "Value = " & myvar.Value
    myvar.Delete
Next myvar
End Sub
