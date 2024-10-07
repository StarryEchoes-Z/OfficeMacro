VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveLineBreaks2UF 
   Caption         =   "RemoveLineBreaks2UF"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "RemoveLineBreaks2UF.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "RemoveLineBreaks2UF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private inputAddVbcr As Boolean
Private selectedText As String
Private addedText As String

'�����񴰿�ʱ
Private Sub UserForm_Initialize()
    selectedText = Selection.Text
End Sub

'������ť
Private Sub CommandButtonVbcr_Click()
    inputAddVbcr = True
    Call removeSub
End Sub

'������ť
Private Sub CommandButtonDoubleSlash_Click()
    inputAddVbcr = False
    addedText = "//"
    Call removeSub
End Sub

'������ť
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
    
    isFirstLine = True ' ���Ϊ��һ�� | Mark as the first line
    
    ' ��ȡѡ�����ı� | Get the selected text
    ' selectedText = Selection.Text
    
    ' ͳһ���з� | Standardize line breaks
    selectedText = Replace(Replace(selectedText, vbCrLf, vbCr), vbLf, vbCr)

    ' ��ѡ�����ı����Ϊ���� | Split the selected text into multiple lines
    Dim arrText() As String ' arrText Ϊ���� | arrText is an array
    arrText = Split(selectedText, vbCr)

    For i = 0 To UBound(arrText) ' UBound()��������������±��ֵ | UBound() returns the highest index value in the array
        
        If arrText(i) <> "" Then
            If Not isFirstLine And allInput <> "" Then ' ��������allInput�ǿ� | Not the first line and allInput is not empty
                If inputAddVbcr Then
                    allInput = allInput & vbCr
                Else
                    allInput = allInput & addedText
                End If
                '*********************************************
            End If
            
            allInput = allInput & arrText(i)
            isFirstLine = False ' �Ѵ�������һ�зǿ����� | At least one non-empty line has been processed
        End If
        
        Next i ' ����ѭ�� | End of loop

'    If MsgBox("ȥ�����к���ı���" & vbCrLf & allInput, vbOKCancel) = 1 Then
'        ' ����û���� OK�����滻ѡ���ı� | If the user clicks OK, replace the selected text
        Selection.TypeText allInput
'    End If

    Unload Me

End Sub
