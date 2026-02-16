Sub InteractiveCitationSuperscript()
    ' ========================================================
    ' 交互式引用上标工具 (v4.0 智能上下文)
    ' 功能：查找文中的引用标记 [数字]，逐个询问是否改为上标
    '
    ' 核心修复：
    ' 1. 终结乱码：改用 "MoveStart wdWord" 智能导航，不再物理切割字符
    '    这能完美跳过隐藏的 REF 域代码，只显示人眼能看到的文字。
    ' 2. 修复了 v3.0 版本中 And 关键字前的空格遗漏问题。
    '
    ' 开发者信息：刘钊昊
    ' 开发日期：2026.02.16
    ' ========================================================

    Dim rng As Range
    Dim searchRng As Range
    Dim userResponse As VbMsgBoxResult
    
    ' 定义专门用于提取上下文的 Range 对象
    Dim contextRngPre As Range
    Dim contextRngPost As Range
    Dim contextPre As String
    Dim contextPost As String
    
    ' === 配置区域 ===
    ' 如果为 True：自动跳过已经是上标格式的 [x]，不弹窗干扰；如果希望每一个已调整的引用上标都显示则设置为False
    Const SKIP_ALREADY_SUPERSCRIPT As Boolean = True
    ' ================

    ' 1. 设置查找范围：从当前光标位置开始往后
    Set searchRng = Selection.Range
    searchRng.End = ActiveDocument.Content.End

    ' 开启查找配置
    With searchRng.Find
        .ClearFormatting
        ' 通配符：匹配 [数字]
        .Text = "\[[0-9]@\]"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        ' 开始循环查找
        Do While .Execute
            ' 将找到的区域赋值给 rng
            Set rng = searchRng.Duplicate
            
            ' === 核心逻辑：检查是否需要跳过 ===
            ' 【修复】补充了 And 前的空格
            If SKIP_ALREADY_SUPERSCRIPT And rng.Font.Superscript = True Then
                ' 折叠范围，继续找下一个
                rng.Collapse Direction:=wdCollapseEnd
                searchRng.Start = rng.End
                GoTo ContinueLoop ' 跳转到循环尾部
            End If

            ' 选中它，让屏幕滚动到这里 (视觉反馈)
            rng.Select
            ActiveWindow.ScrollIntoView rng, True
            
            ' === 1. 获取上下文  ===
            ' 不再用固定区间 Start - 20，而是让 Word 帮我们往回找几个“词”
            ' 这样可以避开切断域代码的问题
            
            ' --- 获取前文 ---
            Set contextRngPre = rng.Duplicate
            contextRngPre.Collapse Direction:=wdCollapseStart ' 塌缩到开头
            ' 往回扩展 6 个词 (wdWord)
            contextRngPre.MoveStart Unit:=wdWord, Count:=-6
            
            ' 双重保险：强制不读取隐藏代码
            contextRngPre.TextRetrievalMode.IncludeFieldCodes = False
            contextRngPre.TextRetrievalMode.IncludeHiddenText = False
            contextPre = contextRngPre.Text
            
            ' 清理回车换行
            contextPre = Replace(contextPre, vbCr, "")
            contextPre = Replace(contextPre, vbLf, "")
            ' 如果太长，只取后 15 个字
            If Len(contextPre) > 15 Then contextPre = Right(contextPre, 15)
            
            ' --- 获取后文 ---
            Set contextRngPost = rng.Duplicate
            contextRngPost.Collapse Direction:=wdCollapseEnd ' 塌缩到结尾
            ' 往后扩展 6 个词
            contextRngPost.MoveEnd Unit:=wdWord, Count:=6
            
            contextRngPost.TextRetrievalMode.IncludeFieldCodes = False
            contextRngPost.TextRetrievalMode.IncludeHiddenText = False
            contextPost = contextRngPost.Text
            
            contextPost = Replace(contextPost, vbCr, "")
            contextPost = Replace(contextPost, vbLf, "")
            If Len(contextPost) > 15 Then contextPost = Left(contextPost, 15)
            
            
            ' === 2. 弹出交互窗口 ===
            userResponse = MsgBox("发现引用标记： " & rng.Text & vbCrLf & vbCrLf & _
                                  "上下文：" & vbCrLf & _
                                  "..." & contextPre & "【 " & rng.Text & " 】" & contextPost & "..." & vbCrLf & vbCrLf & _
                                  "点击【是(Y)】 -> 设为上标 (Superscript)" & vbCrLf & _
                                  "点击【否(N)】 -> 跳过 (保持原样)" & vbCrLf & _
                                  "点击【取消】   -> 终止工具", _
                                  vbYesNoCancel + vbQuestion, "引用上标助手 (刘钊昊 Dev)")
            
            ' === 3. 处理用户选择 ===
            If userResponse = vbYes Then
                ' 执行动作：设为上标
                rng.Font.Superscript = True
                
                ' 关键：处理完后，将 searchRng 的起点移到当前词之后
                rng.Collapse Direction:=wdCollapseEnd
                searchRng.Start = rng.End
                
            ElseIf userResponse = vbNo Then
                ' 跳过
                rng.Collapse Direction:=wdCollapseEnd
                searchRng.Start = rng.End
                
            ElseIf userResponse = vbCancel Then
                MsgBox "工具已终止。", vbInformation
                Exit Sub
            End If
            
ContinueLoop:
        Loop
    End With
    
    MsgBox "文档扫描完成！所有引用标记已检查。", vbInformation, "完成"
End Sub

