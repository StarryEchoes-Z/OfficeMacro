Attribute VB_Name = "NewMacros"
Sub PrintCurrentDateOrTime()

    ' 定义日期、时间变量
    ' Define date and time variables
    Dim currentDate As Date
    Dim currentTime As Variant
    
    '获取日期、时间
    ' Get the current date and time
    currentDate = Now
    currentTime = Time()
    
    ' 格式化日期、时间
    ' Format the date and time
    Dim formattedDate As String
    Dim formattedTime As String
    
    formattedDate = Format(currentDate, "yyyy.mm.dd")
    formattedTime = Format(currentTime, "HH:MM")
    
    '输出
    'Output
    
    'Selection.TypeText formattedDate & "  " & formattedTime
    Selection.TypeText formattedDate
    'Selection.TypeText & formattedTime
    
End Sub
