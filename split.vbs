Option Explicit

' 상수 정의
Const HEADER_ROW As Long = 1
Const DATA_START_ROW As Long = 2
Const SAMPLE_DATA_MAX_ROWS As Long = 10 ' 테스트용 10행
Const BUTTON_WIDTH As Single = 80
Const BUTTON_HEIGHT As Single = 25

' 바이트 단위 문자열 길이 반환
Function LenMbcs(ByVal Mesg As String) As Long
    LenMbcs = LenB(StrConv(Mesg, vbFromUnicode))
End Function

' 바이트 단위 Mid 함수
Function h_mid(ByVal Mesg As String, ByVal startIdx As Long, ByVal Mesglen As Long) As String
    Dim b() As Byte
    b = StrConv(Mesg, vbFromUnicode)
    ' 시작 위치가 전체 바이트 길이보다 크면 빈값 반환
    If startIdx > UBound(b) + 1 Then
        h_mid = ""
    Else
        h_mid = StrConv(MidB(b, startIdx, Mesglen), vbUnicode)
    End If
End Function

' 바이트 단위로 문자열을 맞추고 부족하면 공백을 채우는 함수 (Pad Right)
Function PadRightMbcs(ByVal Mesg As String, ByVal TotalLen As Long) As String
    Dim curLen As Long
    curLen = LenMbcs(Mesg)
    
    If curLen > TotalLen Then
        ' 설정된 길이보다 길면 바이트 단위로 자름
        PadRightMbcs = StrConv(MidB(StrConv(Mesg, vbFromUnicode), 1, TotalLen), vbUnicode)
    Else
        ' 부족하면 그만큼 공백(Space) 추가
        PadRightMbcs = Mesg & Space(TotalLen - curLen)
    End If
End Function

Sub SetBorders(rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Sub CreateSampleData()
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    ws.Cells.Clear ' 시트 초기화
    
    ' 헤더 구성
    Dim headers As Variant
    headers = Array("영문명", "한글명", "길이(Byte)", "데이터 내용")
    ws.Range("A1:D1").Value = headers
    ws.Range("A1:D1").Interior.Color = RGB(240, 240, 240)
    ws.Range("A1:D1").Font.Bold = True
    
    ' 샘플 데이터 채우기
    Dim i As Long
    For i = 0 To 4
        ws.Cells(DATA_START_ROW + i, 1).Value = "FLD_00" & i + 1
        ws.Cells(DATA_START_ROW + i, 2).Value = "필드_" & i + 1
        ws.Cells(DATA_START_ROW + i, 3).Value = 10 ' 각 필드 10바이트 고정
        ws.Cells(DATA_START_ROW + i, 4).Value = "Data" & (i + 1)
        ws.Cells(DATA_START_ROW + i, 4).NumberFormat = "@"
    Next i
    
    ' 결과창 영역 설정 (E2: 입력용, E3: 출력용)
    ws.Range("E1").Value = "전문 입출력"
    ws.Range("E2").Interior.Color = RGB(255, 240, 245) ' 분석용 입력창
    ws.Range("E3").Interior.Color = RGB(255, 250, 205) ' 조립용 결과창
    SetBorders ws.Range("A1:D6")
    SetBorders ws.Range("E1:E3")
    
    ' 버튼 생성
    CreateButton ws, "전문분석기", "SplitSeg", 1, 6
    CreateButton ws, "전문합치기", "MergeSeg", 1, 7
    
    ws.Columns("A:G").AutoFit
    MsgBox "샘플 데이터가 생성되었습니다."
End Sub

Sub CreateButton(ws As Worksheet, btnText As String, macroName As String, topRow As Long, leftCol As Integer)
    Dim btn As Button
    For Each btn In ws.Buttons
        If btn.Caption = btnText Then btn.Delete
    Next btn

    Dim targetCell As Range: Set targetCell = ws.Cells(topRow, leftCol)
    Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, BUTTON_WIDTH, BUTTON_HEIGHT)
    With btn
        .Caption = btnText
        .OnAction = macroName
    End With
End Sub

' [전문분석기] : E2 셀의 전문을 필드별로 쪼개기
Sub SplitSeg()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim fullStr As String: fullStr = Trim(ws.Range("E2").Value)
    Dim curRow As Long, lastRow As Long
    Dim segLen As Long, startByte As Long: startByte = 1

    If fullStr = "" Then
        MsgBox "분석할 전문 내용이 E2 셀에 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    For curRow = DATA_START_ROW To lastRow
        segLen = ws.Cells(curRow, 3).Value
        ws.Cells(curRow, 4).Value = h_mid(fullStr, startByte, segLen)
        startByte = startByte + segLen
    Next curRow

    MsgBox "전문 분석이 완료되었습니다."
End Sub

' [전문합치기] : D열의 내용을 바이트 길이에 맞춰 합쳐서 E3에 출력
Sub MergeSeg()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim curRow As Long, lastRow As Long
    Dim segLen As Long
    Dim resultMsg As String, cellVal As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For curRow = DATA_START_ROW To lastRow
        segLen = ws.Cells(curRow, 3).Value
        cellVal = CStr(ws.Cells(curRow, 4).Value)
        
        ' 바이트 길이에 맞춰 패딩 처리 후 합침
        resultMsg = resultMsg & PadRightMbcs(cellVal, segLen)
    Next curRow

    ws.Range("E3").Value = resultMsg
    MsgBox "전문 조립이 완료되었습니다."
End Sub