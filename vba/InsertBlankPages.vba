'===============================================================================
' 学位论文空白页插入宏
' 功能：
'  1) 更新所有域
'  2) 遍历Heading1标题，若标题所在页为偶数，则在前一Section末尾（分节符之前）插入空白页
'     并从头重新遍历，直到不存在“Heading1在偶数页”的情况
'  3) 再次更新所有域
' 使用：Alt+F11打开VBA编辑器 → 插入模块 → 粘贴代码 → F5运行InsertBlankPages
'===============================================================================

Sub InsertBlankPages()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim pageNum As Long
    Dim insertCount As Long
    Dim styleName As String
    Dim txt As String
    Dim foundEvenPage As Boolean
    Dim maxIterations As Long
    Dim iteration As Long
    Dim secIndex As Long
    Dim prevSec As Section
    Dim lastPara As Paragraph
    
    Set doc = ActiveDocument
    insertCount = 0
    maxIterations = 100
    iteration = 0
    
    Application.ScreenUpdating = False

    ' 1) 更新所有域
    doc.Fields.Update
    
    ' 2) 循环处理，直到没有在偶数页的Heading1
    Do
        foundEvenPage = False
        iteration = iteration + 1
        
        If iteration > maxIterations Then
            MsgBox "已达到最大迭代次数(" & maxIterations & ")，请检查文档结构。", vbExclamation
            Exit Do
        End If
        
        ' 遍历所有段落，找到第一个在偶数页的Heading1（找到后立即处理并从头再来）
        For Each para In doc.Paragraphs
            styleName = para.Style
            
            ' 检查是否为Heading1
            If styleName = "1" Or styleName = "Heading 1" Or styleName = "标题 1" Then
                pageNum = para.Range.Information(wdActiveEndPageNumber)
                secIndex = para.Range.Information(wdActiveEndSectionNumber)
                
                ' 如果从偶数页开始
                If pageNum Mod 2 = 0 Then
                    foundEvenPage = True
                    txt = Trim(Replace(para.Range.Text, Chr(13), ""))

                    ' 第一节没有“前一节”，无法插入
                    If secIndex <= 1 Then
                        MsgBox "检测到标题在偶数页，但该标题位于第1节，无法在前一节末尾插入空白页：" & vbCrLf & txt, vbExclamation
                        Exit For
                    End If

                    Set prevSec = doc.Sections(secIndex - 1)
                    
                    ' Word的分节符通常位于上一节最后一个段落的段落属性中，
                    ' 这里取上一节Range的最后一个段落（通常是空段落/分节符段落），
                    ' 在该段落开始位置插入分页符，即可保证分页符在分节符之前（属于上一节）。
                    Set lastPara = prevSec.Range.Paragraphs(prevSec.Range.Paragraphs.Count)
                    Set rng = lastPara.Range
                    rng.Collapse wdCollapseStart
                    rng.InsertBreak wdPageBreak
                    insertCount = insertCount + 1
                    
                    ' 插入后更新域，避免页码信息滞后
                    doc.Fields.Update
                    
                    ' 找到一个就退出，重新检测
                    Exit For
                End If
            End If
        Next para
        
    Loop While foundEvenPage
    
    ' 3) 再次更新所有域
    doc.Fields.Update

    Application.ScreenUpdating = True

    MsgBox "处理完成！" & vbCrLf & vbCrLf & _
           "共插入 " & insertCount & " 个空白页。" & vbCrLf & _
           "迭代次数: " & iteration & vbCrLf & vbCrLf & _
           "按 Ctrl+Z 可撤销。", vbInformation, "空白页插入"
End Sub
