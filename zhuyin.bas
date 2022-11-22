'判断传入的Unicode是否为中文字符
Function isChinese(uniChar As Integer) As Boolean
    isChinese = uniChar >= 19968 And uniChar <= 40869
End Function

'从Json字符串中提取data字段的数据
Function getDataFromJSON(s As String) As String
    With CreateObject("VBScript.Regexp")
        .Pattern = """data"":""(.*)"""
        getDataFromJSON = .Execute(s)(0).SubMatches(0)
    End With
End Function

'使用http组件调用拼音转换服务获取拼音字符
Function GetPhonetic(strWord As String) As String
    Dim myURL As String
    Dim winHttpReq As Object

    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    myURL = "http://localhost:8080/pinyin1"
    myURL = myURL & "?han=" & strWord

    winHttpReq.Open "GET", myURL, False
    winHttpReq.Send

    GetPhonetic = getDataFromJSON(winHttpReq.responseText)
End Function

'测试GetPhonetic方法
Sub testGetPhonetic()
    ret = GetPhonetic("汗")
    MsgBox ret
End Sub

'Word批量拼音注音
Sub BatchAddPinYin()
    Application.ScreenUpdating = False
    Dim SelectText As String
    Dim PinYinText As String
    Selection.WholeStory
    TextLength = Selection.Characters.Count
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1
    For i = 0 To TextLength
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        With Selection
            SelectText = .Text'基准文字
            If isChinese(AscW(SelectText)) Then'判断是否为中文字符
                PinYinText = GetPhonetic(SelectText)'基准文字 转换为 拼音文字
                If PinYinText <> "" Then
                    .Range.PhoneticGuide _
                        Text:=PinYinText, _'拼音文本
                        Alignment:=wdPhoneticGuideAlignmentCenter, _'对齐方式
                        Raise:=0, _'偏移量（磅）
                        FontSize:=10, _'字号（磅）
                        FontName:="等线"'字体（磅）
                End If
            End If
        End With
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next
    Selection.WholeStory
    Application.ScreenUpdating = True
End Sub

'Word批量使用默认样式加注拼音
Sub BatchAddPinYinByDefaultStyle()
    Application.ScreenUpdating = False
    On Error Resume Next
    Selection.WholeStory
    TextLength = Selection.Characters.Count
    Selection.EndKey
    For i = TextLength To 0 Step -30
        If i <= 30 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=i
            SelectText = Selection.MoveRight(Unit:=wdCharacter, Count:=i,Extend:=wdExtend)
        Else
            Selection.MoveLeft Unit:=wdCharacter, Count:=30
            SelectText = Selection.MoveRight(Unit:=wdCharacter, Count:=30,Extend:=wdExtend)
        End If
        SendKeys "{Enter}"
        Application.Run "FormatPhoneticGuide"
    Next
    Selection.WholeStory
    Application.ScreenUpdating = True
End Sub

'Word批量清除拼音注音
Sub CleanPinYin()
    Application.ScreenUpdating = False
    Selection.WholeStory
    TextLength = Selection.Characters.Count
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1
    For i = 0 To TextLength
        With Selection
             .Range.PhoneticGuide Text:=""
        End With
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next
    Selection.WholeStory
    Application.ScreenUpdating = True
End Sub
