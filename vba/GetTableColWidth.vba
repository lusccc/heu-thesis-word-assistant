Sub 显示表格列宽百分比()
    Dim tbl As Table
    Dim i As Integer
    Dim totalWidth As Single
    Dim msg As String
    Dim objForm As Object
    
    If Selection.Information(wdWithInTable) Then
        Set tbl = Selection.Tables(1)
        
        ' 计算总宽度
        totalWidth = 0
        For i = 1 To tbl.Columns.Count
            totalWidth = totalWidth + tbl.Columns(i).Width
        Next i
        
        ' 生成信息
        msg = "表格列宽百分比：" & vbCrLf & vbCrLf
        
        For i = 1 To tbl.Columns.Count
            msg = msg & "第 " & i & " 列：" & _
                  Round((tbl.Columns(i).Width / totalWidth) * 100, 1) & "%" & vbCrLf
        Next i
        
        ' 复制到剪贴板
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText msg
        dataObj.PutInClipboard
        
        MsgBox "列宽百分比已复制到剪贴板！" & vbCrLf & vbCrLf & msg, vbInformation, "列宽百分比"
    Else
        MsgBox "请先将光标放在表格中", vbExclamation
    End If
End Sub