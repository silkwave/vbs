Option Explicit

' 상수 선언
Const HEADER_ROW As Long = 1                       ' 헤더가 위치하는 행
Const DATA_START_ROW As Long = 2                   ' 데이터 시작 행
Const SAMPLE_DATA_MAX_ROWS As Long = 20            ' 샘플 데이터 최대 행 수
Const BUTTON_WIDTH As Single = 60                   ' 버튼 너비
Const BUTTON_HEIGHT As Single = 30                  ' 버튼 높이

' 다중 바이트 문자 길이를 반환하는 함수
Function LenMbcs(ByVal Mesg As String) As Long
    LenMbcs = LenB(StrConv(Mesg, vbFromUnicode)) ' 입력 문자열의 길이를 바이트 단위로 반환
End Function

' 입력 문자열의 지정된 위치에서 길이에 따라 부분 문자열을 반환하는 함수
Function h_mid(ByVal Mesg As String, ByVal startIdx As Long, ByVal Mesglen As Long) As String
    ' 입력 문자열을 바이트 단위로 변환 후 부분 문자열을 반환하고 다시 유니코드로 변환
    h_mid = StrConv(MidB(StrConv(Mesg, vbFromUnicode), startIdx, Mesglen), vbUnicode)
End Function

Sub SetBorders(rng As Range)
    ' 지정된 범위에 테두리 설정
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
End Sub

Sub CreateSampleData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim dataRange As Range
    Set dataRange = ws.Range("A1:D" & SAMPLE_DATA_MAX_ROWS)

    ' 헤더 생성 및 색상 서식 지정
    With dataRange
        .Cells(HEADER_ROW, 1).Value = "영문명"
        .Cells(HEADER_ROW, 2).Value = "한글명"
        .Cells(HEADER_ROW, 3).Value = "길이"
        .Cells(HEADER_ROW, 4).Value = "전문내용"
        .Interior.Color = RGB(200, 200, 200) ' 회색 배경색
        .Font.Bold = True ' 굵은 글꼴

        ' 테두리 서식 지정
        SetBorders dataRange

        ' 샘플 데이터 입력
        Dim rowIndex As Long
        For rowIndex = DATA_START_ROW To SAMPLE_DATA_MAX_ROWS
            With .Cells(rowIndex, 1)
                .Value = "ENG " & rowIndex
                .Offset(0, 1).Value = "한글 " & rowIndex
                .Offset(0, 2).Value = 10 ' Segment Length
                .Offset(0, 3).Value = "가나다123" ' 메시지
                .Offset(0, 3).NumberFormat = "@" ' 텍스트 형식
            End With
        Next rowIndex
    End With

    ' 1행의 행 높이 설정
    ws.Rows(HEADER_ROW).RowHeight = 30

    ' 분석전문 및 조립전문 열 추가
    Dim highlightRange As Range
    Set highlightRange = ws.Range("E" & HEADER_ROW & ":E" & (HEADER_ROW + 1))

    With highlightRange
        .Cells(2, 1).Interior.Color = RGB(255, 192, 203) ' 분홍색
        .Cells(3, 1).Interior.Color = RGB(255, 165, 0) ' 오렌지색
        .Font.Bold = True ' 굵은 글꼴

        ' 테두리 서식 지정
        SetBorders highlightRange
    End With

    ' 버튼 생성
    CreateButton ws, "전문분석기", "SplitSeg", HEADER_ROW, 5
    CreateButton ws, "전문합치기", "MergeSeg", HEADER_ROW, 6

    ' 열 너비 설정
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 20

    ' MergeSeg 실행
    MergeSeg

    ' 조립전문 열에 값 할당
    highlightRange.Cells(2, 1).Value = highlightRange.Cells(3, 1).Value
End Sub

Sub CreateButton(ws As Worksheet, btnText As String, macroName As String, topRow As Long, leftCol As Integer)
    ' 기존 버튼 삭제
    Dim btn As Button
    For Each btn In ws.Buttons
        If btn.Caption = btnText Then btn.Delete
    Next btn

    ' 버튼 생성
    Dim newButton As Button
    On Error Resume Next ' 오류 발생 시 코드 실행 지속
    Set newButton = ws.Buttons.Add(ws.Cells(topRow, leftCol).Left, ws.Cells(topRow, leftCol).Top, BUTTON_WIDTH, BUTTON_HEIGHT)
    On Error GoTo 0 ' 오류 처리 종료

    If Not newButton Is Nothing Then
        With newButton
            .Caption = btnText
            .OnAction = "'" & macroName & "'" ' 매크로 이름
        End With
    Else
        MsgBox "버튼 생성에 실패했습니다.", vbExclamation
    End If
End Sub

Sub SplitSeg()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim anStr As String
    Dim curRow As Long
    Dim maxRow As Long
    Dim segLen As Long
    Dim curLen As Long

    anStr = Trim(ws.Cells(DATA_START_ROW, 5).Value)
    maxRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    curLen = 1

    ' 데이터 처리 및 서식 적용
    For curRow = DATA_START_ROW To maxRow
        segLen = ws.Cells(curRow, 3).Value
        ws.Cells(curRow, 4).NumberFormat = "@" ' 텍스트 형식
        ws.Cells(curRow, 4).Value = h_mid(anStr, curLen, segLen)
        curLen = curLen + segLen
    Next curRow

    ' 완료 메시지
    MsgBox "(" & ws.Name & ") 전문분석완료", , "전문분석완료/silkwave"

CleanUp:
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, , "오류 번호 : " & Err.Number
    Resume CleanUp
End Sub



Sub MergeSeg()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endRow As Long
    Dim curRow As Long
    Dim segLen As Long
    Dim comMsg As String
    Dim msgPart As String

    Set ws = ActiveSheet
    startRow = DATA_START_ROW
    endRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 각 Segment Length에 따라 메시지를 결합
    For curRow = startRow To endRow
        segLen = ws.Cells(curRow, 3).Value
        msgPart = Left(ws.Cells(curRow, 4).Value, segLen)
        If Len(msgPart) < segLen Then
            msgPart = msgPart & Space(segLen - LenMbcs(msgPart))
        End If
        comMsg = comMsg & msgPart
    Next curRow

    ' 결과 출력
    ws.Cells(3, 5).Value = comMsg
    MsgBox "(" & ws.Name & ") 전문합치기완료", , "전문합치기완료/silkwave"
End Sub


