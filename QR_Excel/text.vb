Private m_KeyCodeDict As Object
Private m_DefaultCode As String

Public Sub InitKeyCodeDict(keyRange As Range, codeRange As Range)
    ' 엑셀 시트에서 키-코드 매핑 딕셔너리 초기화
    ' ARG:
        ' keyRange: 키 이름이 있는 셀 범위
        ' codeRange: 코드가 있는 셀 범위
        '            codeRange가 keyRange보다 크면 초과된 첫 번째 셀을 디폴트 코드로 사용

    Set m_KeyCodeDict = CreateObject("Scripting.Dictionary")
    m_DefaultCode = "mr"  ' 기본값 (엑셀에서 지정하지 않은 경우)

    Dim i As Long
    Dim k As String
    For i = 1 To keyRange.Cells.Count
        k = CleanKey(CStr(keyRange.Cells(i).Value))
        If k <> "" Then
            m_KeyCodeDict(k) = CStr(codeRange.Cells(i).Value)
        End If
    Next i

    If codeRange.Cells.Count > keyRange.Cells.Count Then
        Dim extra As String
        extra = CStr(codeRange.Cells(keyRange.Cells.Count + 1).Value)
        If extra <> "" Then
            m_DefaultCode = extra
        End If
    End If
End Sub

Public Function GetKeyCode(keyName As String) As String
    ' 키 이름에 대응하는 코드 반환
    ' ARG:
        ' keyName: 코드를 조회할 키 이름
    ' RETURN:
        ' 키에 대응하는 코드 (없으면 원래 키 이름 반환)

    If m_KeyCodeDict Is Nothing Then
        GetKeyCode = keyName
        Exit Function
    End If

    Dim trimmedKey As String
    trimmedKey = CleanKey(keyName)

    If m_KeyCodeDict.Exists(trimmedKey) Then
        GetKeyCode = m_KeyCodeDict(trimmedKey)
    Else
        GetKeyCode = keyName
    End If
End Function

Public Function MyTextJoin(Delimiter As String, IgnoreEmpty As Boolean, ParamArray TargetRange() As Variant) As String
    ' 구분자를 이용해서 텍스트 분할하는 함수    Dim r As Variant, cell As Variant
    ' ARG:
        ' Delimiter: 텍스트를 구분할 때 사용할 구분자
        ' IgnoreEmpty: 빈 셀을 무시할지 여부 (True/False)
        ' TargetRange: 텍스트를 추출할 셀 범위 (여러 범위를 지원)
    ' RETURN:
        ' 구분자로 연결된 텍스트 문자열

    Dim result As String

    For Each r In TargetRange
        For Each cell In r
            If Not (IgnoreEmpty And IsEmpty(cell)) Then
                result = result & cell.Value & Delimiter
            End If
        Next cell
    Next r

    If Len(result) > 0 Then
        MyTextJoin = Left(result, Len(result) - Len(Delimiter))
    End If
End Function

Public Function MyTextJson(rootKey As String, ParamArray args() As Variant) As String
    ' 루트키를 이용해서 텍스트를 JSON 형식으로 변환하는 함수
    ' ARG:
        ' rootKey: JSON 객체의 루트 키
        ' args: 텍스트를 추출할 셀 범위 또는 직접 입력된 텍스트 (여러 범위를 지원)
        '       홀수 개인 경우 마지막 항목은 디폴트 코드("mr")를 키로 사용
    ' RETURN:
        ' JSON 형식의 문자열

    Const DEFAULT_CODE As String = "mr"

    Dim cellVals As New Collection
    Dim a As Variant, c As Range
    Dim i As Long

    For Each a In args
        If TypeOf a Is Range Then
            For Each c In a
                cellVals.Add CStr(c.Value)
            Next c
        Else
            cellVals.Add CStr(a)
        End If
    Next a

    Dim innerDict As Object
    Set innerDict = CreateObject("Scripting.Dictionary")
    i = 1
    Do While i <= cellVals.Count
        If i + 1 <= cellVals.Count Then
            innerDict.Add GetKeyCode(cellVals(i)), cellVals(i + 1)
            i = i + 2
        Else
            innerDict.Add DEFAULT_CODE, cellVals(i)
            i = i + 1
        End If
    Loop

    If rootKey = "" Then
        MyTextJson = ConvertToJson(innerDict)
    Else
        Dim rootDict As Object
        Set rootDict = CreateObject("Scripting.Dictionary")
        rootDict.Add rootKey, innerDict
        MyTextJson = ConvertToJson(rootDict)
    End If
End Function

Private Function CleanKey(ByVal s As String) As String
    ' 키 이름에서 공백/제어문자/비표시 문자를 모두 제거
    Dim i As Long
    Dim code As Long
    Dim result As String

    For i = 1 To Len(s)
        code = AscW(Mid(s, i, 1))
        If code < 0 Then code = code + 65536
        Select Case code
            Case 0 To 32    ' 제어문자 + 일반 공백
            Case 127        ' DEL
            Case 160        ' Non-breaking space
            Case 8203       ' Zero-width space
            Case 8204       ' Zero-width non-joiner
            Case 8205       ' Zero-width joiner
            Case 65279      ' BOM
            Case Else
                result = result & Mid(s, i, 1)
        End Select
    Next i

    CleanKey = result
End Function




