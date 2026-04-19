Sub 提取图片宽度百分比()
    ' 提取文档中所有图片相对于页面内容宽度的百分比
    ' 用于将 Word 中手动调整的图片宽度回填到 Quarto (.qmd) 源文件
    '
    ' 输出格式：序号 | 宽度百分比 | fig标签 | 描述(AltText)
    ' 结果自动复制到剪贴板
    
    Dim shp As InlineShape
    Dim contentWidth As Single
    Dim pctExact As Double
    Dim pctRound As Long
    Dim result As String
    Dim idx As Integer
    Dim altText As String
    Dim figLabel As String
    Dim totalImages As Integer
    
    ' ========== 计算页面内容区域宽度（磅） ==========
    ' 注意：多Section文档中，ActiveDocument.PageSetup 在各Section边距
    ' 不统一时会返回 wdUndefined (9999999)，需要通过具体Section获取
    
    Dim pw As Single, lm As Single, rm As Single
    Dim gt As Single  ' 装订线
    pw = 0: lm = 0: rm = 0: gt = 0
    
    ' 先获取页面宽度（各Section通常纸张大小一致）
    On Error Resume Next
    pw = ActiveDocument.PageSetup.PageWidth
    On Error GoTo 0
    
    ' 获取边距：遍历Section找到有效值
    ' （多Section文档中ActiveDocument.PageSetup.LeftMargin可能返回wdUndefined）
    Dim sec As Section
    Dim secLm As Single, secRm As Single, secGt As Single
    Dim marginFound As Boolean
    marginFound = False
    
    On Error Resume Next
    For Each sec In ActiveDocument.Sections
        secLm = sec.PageSetup.LeftMargin
        secRm = sec.PageSetup.RightMargin
        secGt = sec.PageSetup.Gutter
        
        ' 检查值是否合理（排除 wdUndefined=9999999 和异常值）
        If secLm > 0 And secLm < pw And secRm > 0 And secRm < pw Then
            lm = secLm
            rm = secRm
            gt = secGt
            If pw <= 0 Then pw = sec.PageSetup.PageWidth
            marginFound = True
            Exit For
        End If
    Next sec
    On Error GoTo 0
    
    ' 如果Section遍历也失败，尝试通过Selection
    If Not marginFound Then
        On Error Resume Next
        lm = Selection.PageSetup.LeftMargin
        rm = Selection.PageSetup.RightMargin
        gt = Selection.PageSetup.Gutter
        If pw <= 0 Then pw = Selection.PageSetup.PageWidth
        If lm > 0 And lm < pw And rm > 0 And rm < pw Then
            marginFound = True
        End If
        On Error GoTo 0
    End If
    
    contentWidth = pw - lm - rm - gt
    
    If contentWidth <= 0 Or Not marginFound Then
        ' 显示调试信息，提供手动输入
        Dim debugMsg As String
        debugMsg = "自动获取的页面信息：" & vbCrLf & _
                   "PageWidth = " & pw & " pt" & vbCrLf & _
                   "LeftMargin = " & lm & " pt" & vbCrLf & _
                   "RightMargin = " & rm & " pt" & vbCrLf & _
                   "Gutter = " & gt & " pt" & vbCrLf & _
                   "ContentWidth = " & contentWidth & " pt" & vbCrLf & vbCrLf & _
                   "请手动输入页面内容宽度（磅），" & vbCrLf & _
                   "或直接点确定使用默认值 453.6 pt" & vbCrLf & _
                   "（A4纸 595.3pt，内外侧各 2.5cm=70.87pt）"
        
        Dim userInput As String
        userInput = InputBox(debugMsg, "需要手动设置页面内容宽度", "453.6")
        
        If userInput = "" Then
            Exit Sub
        End If
        
        contentWidth = CSng(Val(userInput))
        If contentWidth <= 0 Then contentWidth = 453.6
    End If
    
    ' ========== 预扫描：统计图片总数 ==========
    totalImages = 0
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Or _
           shp.Type = wdInlineShapeLinkedPicture Then
            totalImages = totalImages + 1
        End If
    Next shp
    
    If totalImages = 0 Then
        MsgBox "文档中未找到图片。", vbExclamation
        Exit Sub
    End If
    
    ' ========== 构建输出头 ==========
    result = "# 图片宽度百分比提取结果" & vbCrLf
    result = result & "# 页面内容宽度: " & Round(contentWidth, 1) & " pt (" & _
             Round(contentWidth / 28.35, 1) & " cm)" & vbCrLf
    result = result & "# 图片总数: " & totalImages & vbCrLf
    result = result & "#" & vbCrLf
    result = result & "# 序号" & vbTab & "宽度" & vbTab & _
             "fig标签" & vbTab & "描述(AltText)" & vbCrLf
    result = result & String(80, "-") & vbCrLf
    
    ' ========== 遍历所有图片 ==========
    idx = 0
    For Each shp In ActiveDocument.InlineShapes
        ' 只处理图片类型的 InlineShape
        If shp.Type = wdInlineShapePicture Or _
           shp.Type = wdInlineShapeLinkedPicture Then
            
            idx = idx + 1
            
            ' 计算宽度百分比（四舍五入到整数）
            pctExact = shp.Width / contentWidth * 100
            pctRound = CLng(Int(pctExact + 0.5))
            If pctRound < 1 Then pctRound = 1
            If pctRound > 100 Then pctRound = 100
            
            ' 获取替代文本（Pandoc/Quarto生成的docx中，此为markdown的caption文字）
            altText = ""
            On Error Resume Next
            altText = Trim(shp.AlternativeText)
            On Error GoTo 0
            ' 截断过长的文本
            If Len(altText) > 60 Then altText = Left(altText, 60) & "..."
            
            ' 查找附近的 fig- 书签
            figLabel = FindFigBookmark(shp)
            
            ' 输出行
            result = result & idx & vbTab & _
                     "width=""" & pctRound & "%""" & vbTab & _
                     figLabel & vbTab & _
                     altText & vbCrLf
        End If
    Next shp
    
    result = result & String(80, "-") & vbCrLf
    
    ' ========== 追加便于复制的 QMD width 片段 ==========
    result = result & vbCrLf
    result = result & "# === 便于直接粘贴到 QMD 的 width 属性 ===" & vbCrLf
    
    idx = 0
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Or _
           shp.Type = wdInlineShapeLinkedPicture Then
            idx = idx + 1
            pctExact = shp.Width / contentWidth * 100
            pctRound = CLng(Int(pctExact + 0.5))
            If pctRound < 1 Then pctRound = 1
            If pctRound > 100 Then pctRound = 100
            
            figLabel = FindFigBookmark(shp)
            altText = ""
            On Error Resume Next
            altText = Trim(shp.AlternativeText)
            On Error GoTo 0
            If Len(altText) > 40 Then altText = Left(altText, 40) & "..."
            
            If figLabel <> "" Then
                result = result & "{#" & figLabel & " width=""" & pctRound & "%""}" & _
                         vbTab & "' " & altText & vbCrLf
            Else
                result = result & "#" & idx & " width=""" & pctRound & "%""" & _
                         vbTab & "' " & altText & vbCrLf
            End If
        End If
    Next shp
    
    ' ========== 复制到剪贴板 ==========
    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText result
    dataObj.PutInClipboard
    
    ' ========== 显示摘要 ==========
    Dim summary As String
    summary = "已提取 " & totalImages & " 张图片的宽度信息！" & vbCrLf & vbCrLf
    summary = summary & "页面内容宽度: " & Round(contentWidth / 28.35, 1) & " cm" & vbCrLf
    summary = summary & "结果已复制到剪贴板，可粘贴到文本编辑器查看。" & vbCrLf & vbCrLf
    summary = summary & "使用方式：" & vbCrLf
    summary = summary & "1. 粘贴到文本编辑器查看完整结果" & vbCrLf
    summary = summary & "2. 在 QMD 中搜索对应的 fig 标签或描述文字" & vbCrLf
    summary = summary & "3. 添加或修改 width=""XX%"" 属性"
    
    MsgBox summary, vbInformation, "图片宽度提取完成"
End Sub


Private Function FindFigBookmark(shp As InlineShape) As String
    ' 在图片所在段落及其后续段落中查找 fig- 开头的书签
    ' Quarto/Pandoc 为 {#fig-xxx} 生成对应名称的书签
    
    Dim bk As Bookmark
    Dim searchStart As Long
    Dim searchEnd As Long
    Dim paraRange As Range
    
    FindFigBookmark = ""
    
    On Error Resume Next
    
    ' 获取图片所在段落范围
    Set paraRange = shp.Range.Paragraphs(1).Range.Duplicate
    searchStart = paraRange.Start
    searchEnd = paraRange.End
    
    ' 向后扩展搜索范围（caption/书签通常在图片段落或其后1-3个段落内）
    Dim nextRange As Range
    Set nextRange = paraRange.Duplicate
    nextRange.Collapse wdCollapseEnd
    nextRange.MoveEnd wdParagraph, 3
    searchEnd = nextRange.End
    
    ' 也向前搜索1个段落（某些情况下书签在图片段落之前）
    Dim prevRange As Range
    Set prevRange = paraRange.Duplicate
    prevRange.Collapse wdCollapseStart
    prevRange.MoveStart wdParagraph, -1
    searchStart = prevRange.Start
    
    On Error GoTo 0
    
    ' 在搜索范围内查找 fig- 书签
    For Each bk In ActiveDocument.Bookmarks
        If Left(bk.Name, 4) = "fig-" Or Left(bk.Name, 4) = "fig_" Then
            If bk.Range.Start >= searchStart And bk.Range.Start <= searchEnd Then
                FindFigBookmark = bk.Name
                Exit Function
            End If
        End If
    Next bk
    
End Function
