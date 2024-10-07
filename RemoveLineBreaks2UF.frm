VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveLineBreaks2UF 
   Caption         =   "RemoveLineBreaks2UF"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "RemoveLineBreaks2UF.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "RemoveLineBreaks2UF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private inputAddVbcr As Boolean
Private selectedText As String
Private addedText As String

'打开任务窗口时
Private Sub UserForm_Initialize()
    selectedText = Selection.Text
End Sub

'点击命令按钮
Private Sub CommandButtonVbcr_Click()
    inputAddVbcr = True
    Call removeSub
End Sub

'点击命令按钮
Private Sub CommandButtonDoubleSlash_Click()
    inputAddVbcr = False
    addedText = "//"
    Call removeSub
End Sub

'点击命令按钮
Private Sub CommandButtonNone_Click()
    inputAddVbcr = False
    addedText = ""
    Call removeSub
End Sub

'removeSub
Private Sub removeSub()
    Dim allInput As String
    Dim line As String
    Dim isFirstLine As Boolean
    
    isFirstLine = True ' 标记为第一行 | Mark as the first line
    
    ' 获取选定的文本 | Get the selected text
    ' selectedText = Selection.Text
    
    ' 统一换行符 | Standardize line breaks
    selectedText = Replace(Replace(selectedText, vbCrLf, vbCr), vbLf, vbCr)

    ' 将选定的文本拆分为多行 | Split the selected text into multiple lines
    Dim arrText() As String ' arrText 为数组 | arrText is an array
    arrText = Split(selectedText, vbCr)

    For i = 0 To UBound(arrText) ' UBound()返回数组中最高下标的值 | UBound() returns the highest index value in the array
        
        If arrText(i) <> "" Then
            If Not isFirstLine And allInput <> "" Then ' 非首行且allInput非空 | Not the first line and allInput is not empty
                If inputAddVbcr Then
                    allInput = allInput & vbCr
                Else
                    allInput = allInput & addedText
                End If
                '*********************************************
            End If
            
            allInput = allInput & arrText(i)
            isFirstLine = False ' 已处理至少一行非空内容 | At least one non-empty line has been processed
        End If
        
        Next i ' 结束循环 | End of loop

'    If MsgBox("去除换行后的文本：" & vbCrLf & allInput, vbOKCancel) = 1 Then
'        ' 如果用户点击 OK，则替换选定文本 | If the user clicks OK, replace the selected text
        Selection.TypeText allInput
'    End If

    Unload Me

End Sub
