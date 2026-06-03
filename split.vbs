Option Explicit

'=============================================================================
' [마스터 매크로] 전문분석기 시트 자동 생성 및 UI 디자인 세팅
'=============================================================================
Public Sub Create_Parser_Sheet()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim btnArea As Range
    
    Set wb = ThisWorkbook
    
    ' 1. 기존에 '전문분석기' 시트가 있다면 경고 후 초기화 혹은 새로 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("전문분석기").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = wb.Sheets.Add(Before:=wb.Sheets(1))
    ws.Name = "전문분석기"
    
    ' 화면 그리드 라인 활성화
    ActiveWindow.DisplayGridlines = True
    
    ' 2. 상단 마스터 입력/출력 영역 디자인 (B2, B3 정밀 세팅)
    With ws
        ' 행 높이 설정
        .Rows("1:3").RowHeight = 26
        .Rows("4").RowHeight = 24
        
        ' 타이틀 및 가이드 라벨 레이아웃
        .Range("A2").Value = "최종 전문 결과 (B2)"
        .Range("A3").Value = "원본 전문 소스 (B3)"
        .Range("A2:A3").Font.Bold = True
        .Range("A2:A3").HorizontalAlignment = xlCenter
        .Range("A2:A3").Interior.Color = RGB(240, 242, 245)
        
        ' 핵심 데이터 셀 서식 및 테두리 (텍스트 서식 @ 지정으로 숫자 잘림 방지)
        With .Range("B2:F2")
            .Merge
            .NumberFormat = "@"
            .Interior.Color = RGB(255, 251, 230) ' 결과창: 연한 노란색
        End With
        
        With .Range("B3:F3")
            .Merge
            .NumberFormat = "@"
            .Interior.Color = RGB(240, 248, 255) ' 소스창: 연한 파란색
        End With
        
        ' 테두리 일괄 적용
        .Range("A2:F3").Borders.LineStyle = xlContinuous
        .Range("A2:F3").Borders.Color = RGB(180, 180, 180)
        
        ' 3. 하단 데이터 정의 테이블 헤더 생성 (4행)
        .Range("A4").Value = "순번"
        .Range("B4").Value = "필드명"
        .Range("C4").Value = "길이(Byte)"
        .Range("D4").Value = "데이터 결과 값"
        
        With .Range("A4:D4")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196) ' 신뢰감을 주는 비즈니스 블루
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        ' 4. 실무 가이드용 데이터 샘플 3줄 기본 제공
        .Range("A5:D5").Value = Array(1, "은행코드", 3, "")
        .Range("A6:D6").Value = Array(2, "고객성명", 10, "")
        .Range("A7:D7").Value = Array(3, "거래금액", 12, "")
        .Range("D5:D7").NumberFormat = "@"
        .Range("A5:D7").Borders.LineStyle = xlContinuous
        .Range("A5:A7").HorizontalAlignment = xlCenter
        .Range("C5:C7").HorizontalAlignment = xlCenter
        
        ' 열 너비 최적화 기본값
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 40
        .Columns("E").ColumnWidth = 5
        .Columns("F").ColumnWidth = 5
    End With
    
    ' 5. 실무용 원클릭 매크로 버튼(도형) 자동 생성 및 링크 설정
    ' 위치 지정용 가이드셀 지정 (G열에 나란히 배치)
    CreateVBAButton ws, "H2", "BTN_전문분석", "전문 분석", RGB(46, 117, 182)
    CreateVBAButton ws, "H3", "BTN_전문합치기", "전문 합치기", RGB(112, 173, 71)
    CreateVBAButton ws, "H4", "BTN_초기화", "데이터 초기화", RGB(165, 165, 165)
    CreateVBAButton ws, "H5", "BTN_전문복사", "결과 복사", RGB(255, 192, 0)
    
    MsgBox "'전문분석기' 워크시트와 매크로 버튼이 완벽하게 자동 구축되었습니다!", vbInformation, "구축 완료"
End Sub

' [서브 함수] UI 버튼(도형)을 생성하는 내부 로직
Private Sub CreateVBAButton(ws As Worksheet, cellAddr As String, macroName As String, btnText As String, bgColor As Long)
    Dim targetCell As Range
    Dim btnShape As Shape
    
    Set targetCell = ws.Range(cellAddr)
    
    ' 셀 위치에 맞춰 사각형 도형 생성
    Set btnShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetCell.Left + 2, targetCell.Top + 2, 100, 22)
    
    With btnShape
        .TextFrame.Characters.Text = btnText
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.Solid
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        .OnAction = macroName ' 클릭 시 구동할 매크로 연결
    End With
End Sub


'=============================================================================
' [공통 엔진] 한글 바이트 대응 문자열 추출 (환경 오류 예방 표준 버전)
'=============================================================================
Public Function MidB_Fix(ByVal strInput As String, ByVal startByte As Long, ByVal lengthByte As Long) As String
    Dim arrByte() As Byte
    Dim arrSub() As Byte
    Dim i As Long
    Dim maxIdx As Long
    
    If strInput = "" Or lengthByte <= 0 Then MidB_Fix = "": Exit Function
    
    arrByte = StrConv(strInput, vbFromUnicode)
    maxIdx = UBound(arrByte)
    
    If (startByte - 1) > maxIdx Then MidB_Fix = "": Exit Function
    If (startByte - 1 + lengthByte) > (maxIdx + 1) Then lengthByte = (maxIdx + 1) - (startByte - 1)
    
    ReDim arrSub(lengthByte - 1)
    For i = 0 To lengthByte - 1
        arrSub(i) = arrByte(startByte - 1 + i)
    Next i
    
    MidB_Fix = StrConv(arrSub, vbUnicode)
End Function

'=============================================================================
' [공통 엔진] 한글 바이트 대응 패딩(우측 공백 채우기) 처리
'=============================================================================
Public Function PadRightB(ByVal strInput As String, ByVal totalByte As Long) As String
    Dim arrByte() As Byte
    Dim currentByteLen As Long
    
    arrByte = StrConv(strInput, vbFromUnicode)
    currentByteLen = UBound(arrByte) + 1
    
    If currentByteLen >= totalByte Then
        PadRightB = MidB_Fix(strInput, 1, totalByte)
    Else
        PadRightB = strInput & Space$(totalByte - currentByteLen)
    End If
End Function


'=============================================================================
' 기능 1. 전문 분석 (Parser) - [B3 셀 소스 ?? D열 결과 파싱]
'=============================================================================
Public Sub BTN_전문분석()
    Dim ws As Worksheet
    Dim msg As String
    Dim i As Long
    Dim offset As Long
    Dim lastRow As Long
    Dim fieldLen As Long
    
    Set ws = ThisWorkbook.Sheets("전문분석기")
    msg = CStr(ws.Range("B3").MergeArea.Cells(1, 1).Value)
    
    If Trim$(msg) = "" Then
        MsgBox "B3 셀에 분석할 원본 전문 데이터가 없습니다.", vbExclamation, "알림"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 5 Then Exit Sub
    
    offset = 1
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 5 To lastRow
        If IsNumeric(ws.Cells(i, 3).Value) Then
            fieldLen = CLng(ws.Cells(i, 3).Value)
        Else
            fieldLen = 0
        End If
        
        If fieldLen > 0 Then
            ws.Cells(i, 4).NumberFormat = "@"
            ws.Cells(i, 4).Value = MidB_Fix(msg, offset, fieldLen)
            offset = offset + fieldLen
        Else
            ws.Cells(i, 4).Value = ""
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "전문분석이 정상 완료되었습니다!", vbInformation, "성공"
End Sub


'=============================================================================
' 기능 2. 전문 합치기 (Builder) - [D열 소스 ?? B2 셀 결과 출력]
'=============================================================================
Public Sub BTN_전문합치기()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fieldLen As Long
    Dim fieldData As String
    Dim resultMsg As String
    
    Set ws = ThisWorkbook.Sheets("전문분석기")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 5 Then Exit Sub
    
    Application.ScreenUpdating = False
    resultMsg = ""
    
    For i = 5 To lastRow
        If IsNumeric(ws.Cells(i, 3).Value) Then
            fieldLen = CLng(ws.Cells(i, 3).Value)
        Else
            fieldLen = 0
        End If
        
        If fieldLen > 0 Then
            fieldData = CStr(ws.Cells(i, 4).Value)
            resultMsg = resultMsg & PadRightB(fieldData, fieldLen)
        End If
    Next i
    
    ws.Range("B2").MergeArea.Cells(1, 1).NumberFormat = "@"
    ws.Range("B2").MergeArea.Cells(1, 1).Value = resultMsg
    
    Application.ScreenUpdating = True
    
    MsgBox "전문합치기가 완료되어 B2 셀에 반영되었습니다.", vbInformation, "성공"
End Sub


'=============================================================================
' 기능 3. 데이터 초기화 (병합 오류 완전 차단)
'=============================================================================
Public Sub BTN_초기화()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("전문분석기")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    If lastRow >= 5 Then
        ws.Range("D5:D" & lastRow).ClearContents
    End If
    
    ' 병합 셀 구조를 안전하게 유지하며 값만 소거 (.Value = "")
    ws.Range("B2").MergeArea.Cells(1, 1).Value = ""
    ws.Range("B3").MergeArea.Cells(1, 1).Value = ""
    
    Application.ScreenUpdating = True
    MsgBox "초기화가 완료되었습니다.", vbInformation, "초기화 완료"
End Sub


'=============================================================================
' 기능 4. 클립보드 복사 (백그라운드 복사)
'=============================================================================
Public Sub BTN_전문복사()
    Dim ws As Worksheet
    Dim objData As Object
    Dim targetValue As String
    
    Set ws = ThisWorkbook.Sheets("전문분석기")
    targetValue = CStr(ws.Range("B2").MergeArea.Cells(1, 1).Value)
    
    If Trim$(targetValue) = "" Then
        MsgBox "복사할 전문 결과물이 B2 셀에 존재하지 않습니다.", vbExclamation, "알림"
        Exit Sub
    End If
    
    On Error Resume Next
    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objData.SetText targetValue
    objData.PutInClipboard
    On Error GoTo 0
    
    MsgBox "최종 조립 전문이 클립보드에 복사되었습니다. (Ctrl+V 가능)", vbInformation, "복사 완료"
End Sub

