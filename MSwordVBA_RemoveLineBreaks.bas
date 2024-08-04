Attribute VB_Name = "NewMacros"
Sub RemoveLineBreaks()
    Dim allInput As String
    Dim line As String
    Dim isFirstLine As Boolean
    Dim selectedText As String
    
    isFirstLine = True ' 标记为第一行 | Mark as the first line
    
    ' 获取选定的文本 | Get the selected text
    selectedText = Selection.text

    ' 统一换行符 | Standardize line breaks
    selectedText = Replace(Replace(selectedText, vbCrLf, vbCr), vbLf, vbCr)

    ' 将选定的文本拆分为多行 | Split the selected text into multiple lines
    Dim arrText() As String ' arrText 为数组 | arrText is an array
    arrText = Split(selectedText, vbCr)

    For i = 0 To UBound(arrText) ' UBound()返回数组中最高下标的值 | UBound() returns the highest index of the array
        
        If arrText(i) <> "" Then
            If Not isFirstLine And allInput <> "" Then ' 非首行且allInput非空 | Not the first line and allInput is not empty
                allInput = allInput & "//"
            End If
            
            allInput = allInput & arrText(i)
            isFirstLine = False ' 已处理至少一行非空内容 | At least one non-empty line has been processed
        End If
        
        Next i

    If MsgBox("点击确认替换文本 | Click OK to replace the text" & vbCrLf & allInput, vbOKCancel) = 1 Then
        ' 如果用户点击 OK，则替换选定文本 | If the user clicks OK, replace the selected text
        Selection.TypeText allInput
    End If
End Sub
