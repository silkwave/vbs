Function ConvertSQLToolToMyBatis(inputText As String) As String
    Dim modifiedText As String
    Dim regexPattern As Object
    Dim matchedItems As Object
    Dim currentMatch As Object

    ' 정규표현식 패턴 설정
    Set regexPattern = CreateObject("VBScript.RegExp")
    regexPattern.Global = True
    regexPattern.Pattern = ":\w+" ' ANSI 형식 패턴 (: 뒤에 영문, 숫자, 밑줄)

    ' 정규표현식을 이용해 매치 찾기
    Set matchedItems = regexPattern.Execute(inputText)

    ' 매치된 각각의 문자열을 처리
    modifiedText = inputText
    For Each currentMatch In matchedItems
        ' 매치된 문자를 #{...} 형식으로 변경
        modifiedText = Replace(modifiedText, currentMatch.Value, "#{" & Mid(currentMatch.Value, 2) & "}")
    Next currentMatch

    ConvertSQLToolToMyBatis = modifiedText
End Function

Function ConvertMyBatisToSQLTool(inputText As String) As String
    Dim modifiedText As String
    Dim regexPattern As Object
    Dim matchedItems As Object
    Dim currentMatch As Object

    ' 정규표현식 패턴 설정
    Set regexPattern = CreateObject("VBScript.RegExp")
    regexPattern.Global = True
    regexPattern.Pattern = "#\{(\w+)\}" ' MyBatis 형식 패턴 (#{...} 형식)

    ' 정규표현식을 이용해 매치 찾기
    Set matchedItems = regexPattern.Execute(inputText)

    ' 매치된 각각의 문자열을 처리
    modifiedText = inputText
    For Each currentMatch In matchedItems
        ' 매치된 문자를 :... 형식으로 변경
        modifiedText = Replace(modifiedText, currentMatch.Value, ":" & Mid(currentMatch.Submatches(0), 1))
    Next currentMatch

    ConvertMyBatisToSQLTool = modifiedText
End Function

Sub ClearColumnsBandB()
    ' B열의 내용을 지웁니다.
    Range("B2:B1000").ClearContents
End Sub

Sub ClearColumnsAandA()
    ' A열의 내용을 지웁니다.
    Range("A2:A1000").ClearContents
End Sub

Sub ApplyMyBatisToSQLToolConversion()
    Dim totalRows As Long
    Dim rowIndex As Long
    Dim inputText As String
    
    ClearColumnsBandB

    totalRows = Cells(Rows.Count, 1).End(xlUp).Row ' 데이터가 있는 마지막 행 번호를 가져옴

    ' 각 행에 대해 ConvertMyBatisToSQLTool 함수 적용 (2번째 행부터)
    For rowIndex = 2 To totalRows
        inputText = Cells(rowIndex, 1).Value
        Cells(rowIndex, 2).Value = ConvertMyBatisToSQLTool(inputText)
    Next rowIndex

    Application.ScreenUpdating = True ' 화면 업데이트를 다시 활성화
End Sub

Sub ApplySQLToolToMyBatisConversion()
    Dim totalRows As Long
    Dim rowIndex As Long
    Dim inputText As String
    
    ClearColumnsBandB

    totalRows = Cells(Rows.Count, 1).End(xlUp).Row ' 데이터가 있는 마지막 행 번호를 가져옴

    ' 각 행에 대해 ConvertSQLToolToMyBatis 함수 적용 (2번째 행부터)
    For rowIndex = 2 To totalRows
        inputText = Cells(rowIndex, 1).Value
        Cells(rowIndex, 2).Value = ConvertSQLToolToMyBatis(inputText)
    Next rowIndex

    Application.ScreenUpdating = True ' 화면 업데이트를 다시 활성화
End Sub
