'===============================================================================
' 续表标签插入宏 v8
' 功能：
'  1) 检测跨页表格，找到分页行
'  2) 在分页行处分割表格
'  3) 在新表格前插入"（续表x.y）"段落（右对齐，宋体五号）
'  4) 不重复表头
'  5) 每次分割后重新从头扫描所有表格，确保分页位置准确
' 注意：此宏会修改表格结构，建议在定稿后运行
' 使用：Alt+F11 → 插入模块 → 粘贴代码 → F5运行 AddContinuedTableLabels
'===============================================================================

Option Explicit

Sub AddContinuedTableLabels()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim processedCount As Long
    Dim skippedCount As Long
    processedCount = 0
    skippedCount = 0
    
    doc.Fields.Update
    
    Dim maxIterations As Long
    maxIterations = 200              ' 安全阀，防止无限循环
    
    Dim iteration As Long
    iteration = 0
    
    ' ===== 每次只分割一处，然后重新从头扫描所有表格 =====
    ' 这样能确保每次插入续表字样后页面布局重新计算，分页位置始终准确
    Do
        iteration = iteration + 1
        If iteration > maxIterations Then
            Debug.Print "达到最大迭代次数 " & maxIterations & "，停止处理"
            Exit Do
        End If
        
        Dim foundAndSplit As Boolean
        foundAndSplit = False
        
        ' 强制重新分页，确保 wdActiveEndPageNumber 返回准确页码
        Application.ScreenUpdating = True
        doc.Fields.Update
        doc.Repaginate
        DoEvents
        Application.ScreenUpdating = False
        
        Dim i As Long
        For i = 1 To doc.Tables.Count
            Dim tbl As Table
            Set tbl = doc.Tables(i)
            
            If TableContainsImage(tbl) Then GoTo NextTable
            If Not IsTableCrossPage(tbl) Then GoTo NextTable
            
            Dim tableNum As String
            tableNum = ""
            ExtractTableNum tbl, tableNum
            
            If tableNum = "" Then GoTo NextTable
            
            Dim splitRow As Long
            splitRow = FindPageBreakRow(tbl)
            If splitRow <= 1 Then GoTo NextTable
            
            ' 记录分割前该行所在页码（用于后续验证）
            Dim targetPage As Long
            targetPage = GetRowPage(tbl, splitRow)
            
            ' 执行一次分割：使用 Rows(splitRow) 定位，兼容合并单元格
            On Error Resume Next
            tbl.Rows(splitRow).Select
            If Err.Number <> 0 Then
                Err.Clear
                ' 后备方案：遍历列找可用单元格
                Dim cc As Long
                For cc = 1 To 30
                    tbl.Cell(splitRow, cc).Range.Select
                    If Err.Number = 0 Then Exit For
                    Err.Clear
                Next cc
            End If
            If Err.Number <> 0 Then
                Debug.Print "定位 " & tableNum & " 第" & splitRow & "行出错"
                Err.Clear
                On Error GoTo 0
                GoTo NextTable
            End If
            Selection.SplitTable
            
            If Err.Number <> 0 Then
                Debug.Print "分割 " & tableNum & " 出错: " & Err.Description
                Err.Clear
                On Error GoTo 0
                GoTo NextTable
            End If
            On Error GoTo 0
            
            ' SplitTable后光标在两个表格之间的空段落中
            ' 直接输入续表文字
            Selection.TypeText ChrW(65288) & "续" & tableNum & ChrW(65289)
            
            ' 设置段落格式：右对齐
            Selection.HomeKey wdLine
            Selection.EndKey wdLine, wdExtend
            With Selection.ParagraphFormat
                .Alignment = wdAlignParagraphRight
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
            ' 设置字体：宋体五号
            With Selection.Font
                .NameFarEast = "宋体"
                .Name = "Times New Roman"
                .Size = 10.5
            End With
            
            ' === 验证续表文字是否在新页面 ===
            ' 如果分割后续表文字仍在上一页，则在续表段落前强制分页
            Application.ScreenUpdating = True
            doc.Repaginate
            DoEvents
            
            Selection.HomeKey wdLine
            Dim labelPage As Long
            labelPage = Selection.Range.Information(wdActiveEndPageNumber)
            
            ' 获取上半部分表格起始页
            Dim upperTblPage As Long
            Dim upperRng As Range
            Set upperRng = doc.Range(tbl.Range.Start, tbl.Range.Start + 1)
            upperTblPage = upperRng.Information(wdActiveEndPageNumber)
            
            If labelPage <= upperTblPage Then
                ' 续表文字和上半部分表格在同一页，需要强制分页
                Selection.HomeKey wdLine
                Selection.ParagraphFormat.PageBreakBefore = True
                Debug.Print "为 " & tableNum & " 续表标签添加了强制分页"
            End If
            
            Application.ScreenUpdating = False
            
            processedCount = processedCount + 1
            foundAndSplit = True
            Exit For                 ' 立即跳出内层循环，重新从头扫描
            
NextTable:
        Next i
        
    Loop While foundAndSplit         ' 没有找到可分割的跨页表格时退出
    
    ' 统计仍然跨页但无法处理的表格（无编号等）
    Application.ScreenUpdating = True
    doc.Repaginate
    Dim j As Long
    For j = 1 To doc.Tables.Count
        If Not TableContainsImage(doc.Tables(j)) Then
            If IsTableCrossPage(doc.Tables(j)) Then
                skippedCount = skippedCount + 1
            End If
        End If
    Next j
    
    doc.Fields.Update
    
    MsgBox "处理完成！" & vbCrLf & vbCrLf & _
           "共分割了 " & processedCount & " 处跨页表格。" & vbCrLf & _
           IIf(skippedCount > 0, "仍有 " & skippedCount & " 个跨页表格未处理（无表编号等）。" & vbCrLf, "") & _
           IIf(iteration > maxIterations, "警告：达到最大迭代次数，可能未完全处理。" & vbCrLf, "") & _
           vbCrLf & "按 Ctrl+Z 可撤销。", vbInformation, "续表处理"
End Sub


'-------------------------------------------------------------------------------
' 找到分页行索引（兼容合并单元格）
'-------------------------------------------------------------------------------
Private Function FindPageBreakRow(tbl As Table) As Long
    FindPageBreakRow = -1
    
    Dim rngStart As Range
    Set rngStart = ActiveDocument.Range(tbl.Range.Start, tbl.Range.Start + 1)
    Dim startPage As Long
    startPage = rngStart.Information(wdActiveEndPageNumber)
    
    Dim maxRow As Long
    maxRow = SafeRowCount(tbl)
    
    Dim r As Long
    For r = 2 To maxRow
        Dim rowPage As Long
        rowPage = GetRowPage(tbl, r)
        
        If rowPage > startPage Then
            FindPageBreakRow = r
            Exit Function
        End If
    Next r
End Function


'-------------------------------------------------------------------------------
' 获取表格指定行的页码（兼容合并单元格）
' 依次尝试: Rows(r).Range.Start → 遍历列找可用Cell
'-------------------------------------------------------------------------------
Private Function GetRowPage(tbl As Table, rowIdx As Long) As Long
    GetRowPage = -1
    On Error Resume Next
    
    ' 方法1：通过 Rows(rowIdx).Range 的起始位置获取页码
    Dim rowRng As Range
    Set rowRng = tbl.Rows(rowIdx).Range
    If Err.Number = 0 Then
        Dim rng1 As Range
        Set rng1 = ActiveDocument.Range(rowRng.Start, rowRng.Start + 1)
        If Err.Number = 0 Then
            GetRowPage = rng1.Information(wdActiveEndPageNumber)
            If Err.Number = 0 Then
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If
    Err.Clear
    
    ' 方法2：遍历列找任意可用单元格
    Dim c As Long
    For c = 1 To 30
        Dim cellRng As Range
        Set cellRng = tbl.Cell(rowIdx, c).Range
        If Err.Number = 0 Then
            Dim rng2 As Range
            Set rng2 = ActiveDocument.Range(cellRng.Start, cellRng.Start + 1)
            If Err.Number = 0 Then
                GetRowPage = rng2.Information(wdActiveEndPageNumber)
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
        Err.Clear
    Next c
    
    On Error GoTo 0
End Function


Private Function SafeRowCount(tbl As Table) As Long
    On Error Resume Next
    SafeRowCount = tbl.Rows.Count
    If Err.Number <> 0 Then
        Err.Clear
        SafeRowCount = 0
        Dim c As Cell
        For Each c In tbl.Range.Cells
            If c.RowIndex > SafeRowCount Then SafeRowCount = c.RowIndex
        Next c
    End If
    On Error GoTo 0
    If SafeRowCount = 0 Then SafeRowCount = 1
End Function


Private Function IsTableCrossPage(tbl As Table) As Boolean
    IsTableCrossPage = False
    Dim rngS As Range, rngE As Range
    Set rngS = ActiveDocument.Range(tbl.Range.Start, tbl.Range.Start + 1)
    Dim ep As Long
    ep = tbl.Range.End - 1
    If ep <= tbl.Range.Start Then ep = tbl.Range.Start + 1
    Set rngE = ActiveDocument.Range(ep - 1, ep)
    Dim sp As Long, endp As Long
    On Error Resume Next
    sp = rngS.Information(wdActiveEndPageNumber)
    endp = rngE.Information(wdActiveEndPageNumber)
    On Error GoTo 0
    If endp > sp Then IsTableCrossPage = True
End Function


Private Sub ExtractTableNum(tbl As Table, ByRef tableNum As String)
    Dim doc As Document
    Set doc = ActiveDocument
    Dim tblStart As Long
    tblStart = tbl.Range.Start
    Dim ssp As Long
    ssp = tblStart - 2000
    If ssp < 0 Then ssp = 0
    Dim sr As Range
    Set sr = doc.Range(ssp, tblStart)
    Dim pc As Long
    pc = sr.Paragraphs.Count
    Dim sf As Long
    sf = pc - 5
    If sf < 1 Then sf = 1
    Dim j As Long
    For j = pc To sf Step -1
        Dim para As Paragraph
        Set para = sr.Paragraphs(j)
        Dim pt As String
        pt = Trim(Replace(para.Range.Text, Chr(13), ""))
        If pt Like "表*.*[ ]*" Then
            Dim pos As Long
            pos = InStr(pt, " ")
            If pos > 2 Then
                tableNum = Left(pt, pos - 1)
                Exit For
            End If
        End If
    Next j
End Sub


Private Function TableContainsImage(tbl As Table) As Boolean
    TableContainsImage = False
    Dim shp As InlineShape
    For Each shp In tbl.Range.InlineShapes
        If shp.Type = wdInlineShapePicture Or shp.Type = wdInlineShapeLinkedPicture Then
            TableContainsImage = True
            Exit Function
        End If
    Next shp
End Function
