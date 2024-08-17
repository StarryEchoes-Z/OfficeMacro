Attribute VB_Name = "NewMacros"
Sub RemoveLineBreaks()
    Dim allInput As String
    Dim line As String
    Dim isFirstLine As Boolean
    Dim selectedText As String
    
    isFirstLine = True ' ���Ϊ��һ�� | Mark as the first line
    
    ' ��ȡѡ�����ı� | Get the selected text
    selectedText = Selection.text
    
    ' ͳһ���з� | Standardize line breaks
    selectedText = Replace(Replace(selectedText, vbCrLf, vbCr), vbLf, vbCr)

    ' ��ѡ�����ı����Ϊ���� | Split the selected text into multiple lines
    Dim arrText() As String ' arrText Ϊ���� | arrText is an array
    arrText = Split(selectedText, vbCr)

    For i = 0 To UBound(arrText) ' UBound()��������������±��ֵ | UBound() returns the highest index value in the array
        
        If arrText(i) <> "" Then
            If Not isFirstLine And allInput <> "" Then ' ��������allInput�ǿ� | Not the first line and allInput is not empty
                allInput = allInput & "//"
            End If
            
            allInput = allInput & arrText(i)
            isFirstLine = False ' �Ѵ�������һ�зǿ����� | At least one non-empty line has been processed
        End If
        
        Next i ' ����ѭ�� | End of loop

    If MsgBox("ȥ�����к���ı���" & vbCrLf & allInput, vbOKCancel) = 1 Then
        ' ����û���� OK�����滻ѡ���ı� | If the user clicks OK, replace the selected text
        Selection.TypeText allInput
    End If
End Sub
