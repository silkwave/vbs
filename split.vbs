Option Explicit

' ============================================================
' [상수 정의 영역]
' ============================================================
Private Const HEADER_ROW      As Long = 1    ' 헤더 행
Private Const DATA_START_ROW  As Long = 2    ' 데이터 시작 행
Private Const BUTTON_WIDTH    As Single = 80 ' 버튼 너비
Private Const BUTTON_HEIGHT   As Single = 25 ' 버튼 높이

' 컬럼 정의 (매직넘버 제거)
Private Const COL_ENG_NAME  As Long = 1  ' A열: 영문명
Private Const COL_KOR_NAME  As Long = 2  ' B열: 한글명
Private Const COL_BYTE_LEN  As Long = 3  ' C열: 길이(Byte)
Private Const COL_DATA      As Long = 4  ' D열: 데이터
Private Const COL_RESULT    As Long = 5  ' E열: 전문 입출력

' E열 내 역할 정의
Private Const ROW_LABEL     As Long = 1  ' E1: 라벨
Private Const ROW_INPUT     As Long = 2  ' E2: 전문 입력 (분석용)
Private Const ROW_OUTPUT    As Long = 3  ' E3: 전문 출력 (조립 결과)

' ============================================================
' [유틸] 문자열의 바이트 길이를 반환
' - 한글/영문 혼합 시 정확한 길이 계산
' ============================================================
Public Function LenMbcs(ByVal s As String) As Long
    LenMbcs = LenB(StrConv(s, vbFromUnicode))
End Function

' ============================================================
' [유틸] 바이트 기준 문자열 자르기
' - startByte : 시작 위치 (1부터 시작)
' - byteLen   : 자를 바이트 길이
' - 내부적으로 문자열 → 바이트 변환 후 처리
' ============================================================
Public Function MidMbcs(ByVal s As String, _
                       ByVal startByte As Long, _
                       ByVal byteLen As Long) As String

    Dim byteStr As String

    ' 유니코드 문자열을 바이트 문자열로 변환
    byteStr = StrConv(s, vbFromUnicode)

    ' 시작 위치가 전체 길이를 초과하면 빈 문자열 반환
    If startByte > LenB(byteStr) Then
        MidMbcs = ""
        Exit Function
    End If

    ' MidB로 바이트 단위 추출 후 다시 문자열로 변환
    MidMbcs = StrConv(MidB(byteStr, startByte, byteLen), vbUnicode)

End Function

' ============================================================
' [유틸] 바이트 기준 오른쪽 패딩
' - totalLen에 맞게 공백(Space) 채움
' - 초과 시 바이트 기준으로 자름
' ============================================================
Public Function PadRightMbcs(ByVal s As String, ByVal totalLen As Long) As String

    Dim curLen As Long
    curLen = LenMbcs(s)

    ' 현재 길이가 더 크면 잘라냄
    If curLen > totalLen Then
        PadRightMbcs = MidMbcs(s, 1, totalLen)
    Else
        ' 부족하면 공백으로 채움
        PadRightMbcs = s & Space(totalLen - curLen)
    End If

End Function

' ============================================================
' [유틸] 테두리 적용
' ============================================================
Private Sub SetBorders(rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

' ============================================================
' [유틸] 마지막 데이터 행 찾기 (C열 기준)
' ============================================================
Private Function GetLastDataRow(ws As Worksheet) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, COL_BYTE_LEN).End(xlUp).Row
End Function

' ============================================================
' [유틸] 전체 정의된 바이트 길이 합계 계산
' ============================================================
Private Function GetTotalDefinedBytes(ws As Worksheet) As Long

    Dim r As Long
    Dim total As Long

    For r = DATA_START_ROW To GetLastDataRow(ws)
        If IsNumeric(ws.Cells(r, COL_BYTE_LEN).Value) Then
            total = total + CLng(ws.Cells(r, COL_BYTE_LEN).Value)
        End If
    Next r

    GetTotalDefinedBytes = total

End Function

' ============================================================
' [메인] 샘플 데이터 생성
' - 테스트용 필드 정의 자동 생성
' ============================================================
Public Sub CreateSampleData()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 기존 데이터 초기화
    ws.Cells.Clear

    ' 헤더 생성
    ws.Cells(1, 1).Value = "영문명"
    ws.Cells(1, 2).Value = "한글명"
    ws.Cells(1, 3).Value = "길이(Byte)"
    ws.Cells(1, 4).Value = "데이터 내용"

    ' 헤더 스타일
    With ws.Range("A1:D1")
        .Interior.Color = RGB(220, 230, 241)
        .Font.Bold = True
    End With

    ' 샘플 데이터 (금융 전문 구조 예시)
    Dim sampleData As Variant
    sampleData = Array( _
        Array("FLD_001", "거래구분", 4, "1001"), _
        Array("FLD_002", "기관코드", 6, "000001"), _
        Array("FLD_003", "계좌번호", 20, "12345678901234567890"), _
        Array("FLD_004", "고객명", 20, "홍길동"), _
        Array("FLD_005", "거래금액", 15, "1000000") _
    )

    Dim i As Long
    For i = 0 To UBound(sampleData)
        ws.Cells(DATA_START_ROW + i, COL_ENG_NAME).Value = sampleData(i)(0)
        ws.Cells(DATA_START_ROW + i, COL_KOR_NAME).Value = sampleData(i)(1)
        ws.Cells(DATA_START_ROW + i, COL_BYTE_LEN).Value = sampleData(i)(2)
        ws.Cells(DATA_START_ROW + i, COL_DATA).Value = sampleData(i)(3)
        ws.Cells(DATA_START_ROW + i, COL_DATA).NumberFormat = "@"
    Next i

    ' 입력/출력 영역 설정
    ws.Cells(ROW_LABEL, COL_RESULT).Value = "전문 입출력"
    ws.Cells(ROW_INPUT, COL_RESULT).Interior.Color = RGB(255, 228, 235)
    ws.Cells(ROW_OUTPUT, COL_RESULT).Interior.Color = RGB(255, 250, 205)

    ' 테두리 적용
    SetBorders ws.Range("A1:D6")
    SetBorders ws.Range("E1:E3")

    ' 버튼 생성
    CreateButton ws, "전문분석기", "SplitSeg", 1, 6
    CreateButton ws, "전문합치기", "MergeSeg", 2, 6

    ws.Columns("A:G").AutoFit

    MsgBox "샘플 데이터 생성 완료"

End Sub

' ============================================================
' [유틸] 버튼 생성 (중복 제거 포함)
' ============================================================
Private Sub CreateButton(ws As Worksheet, _
                         btnText As String, _
                         macroName As String, _
                         topRow As Long, _
                         leftCol As Long)

    Dim btn As Button

    ' 기존 동일 버튼 제거
    For Each btn In ws.Buttons
        If btn.Caption = btnText Then btn.Delete
    Next btn

    Dim c As Range
    Set c = ws.Cells(topRow, leftCol)

    Set btn = ws.Buttons.Add(c.Left, c.Top, BUTTON_WIDTH, BUTTON_HEIGHT)

    With btn
        .Caption = btnText
        .OnAction = macroName
    End With

End Sub

' ============================================================
' [기능] 전문 분석
' - E2 전체 전문 → D열 필드별 분해
' ============================================================
Public Sub SplitSeg()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim fullStr As String
    fullStr = ws.Cells(ROW_INPUT, COL_RESULT).Value

    ' 입력값 검증
    If Len(Trim(fullStr)) = 0 Then
        MsgBox "E2에 전문을 입력하세요.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)

    Dim startByte As Long
    startByte = 1

    Dim r As Long
    For r = DATA_START_ROW To lastRow

        Dim segLen As Long
        segLen = CLng(ws.Cells(r, COL_BYTE_LEN).Value)

        ' 바이트 기준 분해
        ws.Cells(r, COL_DATA).Value = MidMbcs(fullStr, startByte, segLen)

        startByte = startByte + segLen

    Next r

    MsgBox "전문 분석 완료"
    Exit Sub

ErrHandler:
    MsgBox "오류: " & Err.Description

End Sub

' ============================================================
' [기능] 전문 조립
' - D열 필드 → E3 전체 전문 생성
' ============================================================
Public Sub MergeSeg()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)

    Dim resultMsg As String
    Dim warnings As String

    Dim r As Long
    For r = DATA_START_ROW To lastRow

        Dim segLen As Long
        segLen = CLng(ws.Cells(r, COL_BYTE_LEN).Value)

        Dim val As String
        val = CStr(ws.Cells(r, COL_DATA).Value)

        ' 길이 초과 체크
        If LenMbcs(val) > segLen Then
            warnings = warnings & "행 " & r & " 초과 → 잘림" & vbCrLf
        End If

        ' 바이트 기준 패딩 후 결합
        resultMsg = resultMsg & PadRightMbcs(val, segLen)

    Next r

    ws.Cells(ROW_OUTPUT, COL_RESULT).Value = resultMsg

    Dim msg As String
    msg = "전문 조립 완료 (" & LenMbcs(resultMsg) & " byte)"

    If warnings <> "" Then
        msg = msg & vbCrLf & vbCrLf & "※ 길이 초과 필드:" & vbCrLf & warnings
    End If

    MsgBox msg

    Exit Sub

ErrHandler:
    MsgBox "오류: " & Err.Description

End Sub