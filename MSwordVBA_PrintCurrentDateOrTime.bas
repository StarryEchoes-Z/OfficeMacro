Attribute VB_Name = "NewMacros"
Sub PrintCurrentDateOrTime()

    ' �������ڡ�ʱ�����
    ' Define date and time variables
    Dim currentDate As Date
    Dim currentTime As Variant
    
    '��ȡ���ڡ�ʱ��
    ' Get the current date and time
    currentDate = Now
    currentTime = Time()
    
    ' ��ʽ�����ڡ�ʱ��
    ' Format the date and time
    Dim formattedDate As String
    Dim formattedTime As String
    
    formattedDate = Format(currentDate, "yyyy.mm.dd")
    formattedTime = Format(currentTime, "HH:MM")
    
    '���
    'Output
    
    'Selection.TypeText formattedDate & "  " & formattedTime
    Selection.TypeText formattedDate
    'Selection.TypeText & formattedTime
    
End Sub
