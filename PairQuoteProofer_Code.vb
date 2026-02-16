Sub PairQuoteProofer()
    ' ========================================================
    ' 交互式英文引号校对工具
    ' 核心思路：一次性查找成对的 "..."，确保左右方向绝对正确
    ' 开发者信息：刘钊昊
    ' 开发日期：2026.02.16
    ' ========================================================

    Dim rng As Range
    Dim searchRng As Range
    Dim contentText As String
    Dim userResponse As VbMsgBoxResult
    
    ' 1. 设置查找范围：从当前光标位置开始往后
    Set searchRng = Selection.Range
    searchRng.End = ActiveDocument.Content.End
    
    ' 开启查找配置
    With searchRng.Find
        .ClearFormatting
        ' ========================================================
        ' 核心通配符解释：
        ' [!^13""] 意思是：既不是段落标记(^13)，也不是引号(")的字符
        ' @        意思是：重复1次或多次
        ' 整体逻辑：找一对引号，中间夹着一段不换行的内容
        ' ========================================================
        .Text = "[""][!^13""]@[""]"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        ' 开始循环查找
        Do While .Execute
            ' 将找到的范围赋值给 rng 方便操作
            Set rng = searchRng.Duplicate
            
            ' 选中它，让屏幕滚动到这里
            rng.Select
            ActiveWindow.ScrollIntoView rng, True
            
            ' 提取中间的内容（去掉两头的直引号）
            ' Mid(文本, 2, 长度-2)
            contentText = Mid(rng.Text, 2, Len(rng.Text) - 2)
            
            ' 2. 弹出交互窗口
            ' 直接显示替换后的效果供你确认
            userResponse = MsgBox("发现成对的直引号内容：" & vbCrLf & vbCrLf & _
                                  "原始内容： " & rng.Text & vbCrLf & _
                                  "替换预览： “" & contentText & "”" & vbCrLf & vbCrLf & _
                                  "点击【是(Y)】 -> 替换为 “...” (左右弯引号)" & vbCrLf & _
                                  "点击【否(N)】 -> 跳过此组" & vbCrLf & _
                                  "点击【取消】   -> 停止", _
                                  vbYesNoCancel + vbQuestion, "成对引号校对助手")
            
            ' 3. 处理用户选择
            If userResponse = vbYes Then
                ' 执行替换：手动构建新字符串，确保万无一失
                rng.Text = ChrW(8220) & contentText & ChrW(8221)
                
                ' 替换后，将 searchRng 的起始位置移到替换后的末尾，防止死循环
                searchRng.Collapse Direction:=wdCollapseEnd
                
            ElseIf userResponse = vbNo Then
                ' 跳过，继续找下一个
                searchRng.Collapse Direction:=wdCollapseEnd
                
            ElseIf userResponse = vbCancel Then
                MsgBox "校对已终止。", vbInformation
                Exit Sub
            End If
        Loop
    End With
    
    MsgBox "文档扫描完成！所有成对引号已处理。", vbInformation, "完成"
End Sub


