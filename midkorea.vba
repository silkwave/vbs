'// Left 함수를 대신한 한글처리 가능한 H_Left 함수
Public Function H_Left(SrcString As String, iLength As Integer) As String
    H_Left = StrConv(LeftB(StrConv(SrcString, vbFromUnicode), iLength), vbUnicode)
End Function

'// Right 함수를 대신한 한글처리 가능한 H_Right 함수
Public Function H_Right(SrcString As String, iLength As Integer) As String
    H_Right = StrConv(RightB(StrConv(SrcString, vbFromUnicode), iLength), vbUnicode)
End Function

'// Mid 함수를 대신한 한글처리 가능한 H_Mid 함수
Public Function H_Mid(SrcString As String, idxStr As Integer, iLength As Integer) As String
    H_Mid = StrConv(MidB(StrConv(SrcString, vbFromUnicode), idxStr, iLength), vbUnicode)
End Function

'// Len 함수를 대신한 한글처리 가능한 H_Len 함수
Public Function H_Len(SrcString As String) As Integer
    H_Len = LenB(StrConv(SrcString, vbFromUnicode))
End Function
