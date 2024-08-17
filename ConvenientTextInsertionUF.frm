VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvenientTextInsertionUF 
   Caption         =   "ConvenientTextInsertion"
   ClientHeight    =   2604
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3276
   OleObjectBlob   =   "ConvenientTextInsertionUF.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ConvenientTextInsertionUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 在 UserForm1 的模块级别声明一个布尔常量
Private Const closeAfterClick As Boolean = True

' 子程序用于关闭用户窗体（如果需要）
Private Sub CloseUserFormIfRequired()
    If closeAfterClick Then
        DoEvents
        Unload Me
    End If
End Sub

' 通用过程来处理按钮点击事件
Private Sub Button_ClickHandler(ByRef btn As CommandButton)
    Dim cbc As CommandButton
    Set cbc = btn
    Selection.TypeText cbc.Caption
    Call CloseUserFormIfRequired
End Sub

' CommandButton1 的 Click 事件处理程序
Private Sub CommandButton1_Click()
    Call Button_ClickHandler(Me.CommandButton1)
End Sub

' CommandButton2 的 Click 事件处理程序
Private Sub CommandButton2_Click()
    Call Button_ClickHandler(Me.CommandButton2)
End Sub

' CommandButton3 的 Click 事件处理程序
Private Sub CommandButton3_Click()
    Call Button_ClickHandler(Me.CommandButton3)
End Sub

Private Sub UserForm_Click()

End Sub
