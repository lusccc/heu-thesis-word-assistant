Sub fieldcodetotext()
    Dim MyString As String
    ActiveWindow.View.ShowFieldCodes = True
    For Each aField In ActiveDocument.Fields
        aField.Select
        MyString = "{ " & Selection.Fields(1).Code.Text & " }"
        Selection.Text = MyString
    Next aField
    ActiveWindow.View.ShowFieldCodes = False
End Sub