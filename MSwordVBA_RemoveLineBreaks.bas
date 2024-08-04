Attribute VB_Name = "NewMacros"
Sub RemoveLineBreaks()
    Dim allInput As String
    Dim line As String
    Dim isFirstLine As Boolean
    Dim selectedText As String
    
    isFirstLine = True ' ���Ϊ��һ�� | Mark as the first line
    
    ' ��ȡѡ�����ı� | Get the selected text
    selectedText = Selection.text
    
    ' ��ѡ�����ı����Ϊ���� | Split the selected text into multiple lines
    Dim arrText() As String ' arrText Ϊ����
    arrText = Split(selectedText, vbCr)
    ' ֻ��ʶ��vbCr�����Ľ� | Only recognizes vbCr, needs improvement

    For i = 0 To UBound(arrText) ' UBound()��������������±��ֵ | UBound() returns the highest index of the array
        
        arrText(i) = RemoveNewLines(arrText(i))
        
        If arrText(i) <> "" Then
            If Not isFirstLine And allInput <> "" Then ' ��������allInput�ǿ� | Not the first line and allInput is not empty
                allInput = allInput & "//"
            End If
            
            allInput = allInput & arrText(i)
            isFirstLine = False ' �Ѵ�������һ�зǿ����� | At least one non-empty line has been processed
        End If
        
        Next i

    If MsgBox("���ȷ���滻�ı� | Click OK to replace the text" & vbCrLf & allInput, vbOKCancel) = 1 Then
        Selection.TypeText allInput
    End If
End Sub

Function RemoveNewLines(text As String) As String
    RemoveNewLines = Replace(Replace(text, vbLf, ""), vbCrLf, "")
End Function
