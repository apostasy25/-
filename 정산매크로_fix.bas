' =====================================================================
' 정산매크로 수정 패치 (claude/fix-settlement-mapping-PXjjf)
'
' 아래 5개 함수를 기존 .bas / VBA 모듈에서 동일 이름으로 교체하세요.
'   1) ShouldSkipSourceRow   - 문제1 해결 (계약금 오스킵 제거)
'   2) IsSubtotalRow         - 문제1 해결 (중간 Exit For 오탐 제거)
'   3) ExtractPlatformFromBigo - 문제2 해결 (플랫폼 사전 기반 스캔)
'   4) ParseBigoGrossAndRS   - 신규 (비고에서 gross/RS 파싱)
'   5) WriteRawDataRow       - 문제3 해결 (금액 분해 로직)
' =====================================================================

' ---------- 1) 문제1: ShouldSkipSourceRow ----------
' 변경점: "계" 로 시작하는 값을 스킵하던 조건 제거 (계약금 오스킵 방지)
'         주석/화살표 휴리스틱을 '핵심 식별정보가 모두 비어있을 때'에만 적용
Private Function ShouldSkipSourceRow(ByVal ws As Worksheet, ByVal headers As Object, ByVal r As Long) As Boolean
    Dim s As String, i As Long
    Dim pieces(1 To 7) As String
    Dim sTitle As String, sCode As String, sAuthor As String, sReal As String

    pieces(1) = CStr(ReadByAliases(ws, headers, r, Array("플랫폼명", "플랫폼", "거래처", "거래처명")))
    pieces(2) = CStr(ReadByAliases(ws, headers, r, Array("작품명", "작품 명", "타이틀", "제목")))
    pieces(3) = CStr(ReadByAliases(ws, headers, r, Array("작품코드", "코드")))
    pieces(4) = CStr(ReadByAliases(ws, headers, r, Array("구분", "정산용구분")))
    pieces(5) = CStr(ReadByAliases(ws, headers, r, Array("유형", "정산용유형", "사업구분")))
    pieces(6) = CStr(ReadByAliases(ws, headers, r, Array("필명", "작가명")))
    pieces(7) = CStr(ReadByAliases(ws, headers, r, Array("저자명", "실명", "작가실명")))

    ' 합계/소계 명시 텍스트만 필터 ("계"로 시작 조건 제거 → 계약금 보존)
    For i = LBound(pieces) To UBound(pieces)
        s = NormalizeHeader(pieces(i))
        If Len(s) > 0 Then
            If s = "합계" Or s = "총합계" Or s = "소계" Or s = "누계" Then
                ShouldSkipSourceRow = True
                Exit Function
            End If
        End If
    Next i

    ' 핵심 식별정보(작품명/코드/필명/저자명) 전부 비어있으면 요약/주석행
    sTitle = Trim$(CStr(ReadByAliases(ws, headers, r, Array("작품명", "작품 명", "타이틀", "제목"))))
    sCode = Trim$(CStr(ReadByAliases(ws, headers, r, Array("작품코드", "코드"))))
    sAuthor = Trim$(CStr(ReadByAliases(ws, headers, r, Array("필명", "작가명"))))
    sReal = Trim$(CStr(ReadByAliases(ws, headers, r, Array("저자명", "실명", "작가실명"))))

    If Len(sTitle) = 0 And Len(sCode) = 0 And Len(sAuthor) = 0 And Len(sReal) = 0 Then
        ShouldSkipSourceRow = True
    End If
End Function

' ---------- 2) 문제1: IsSubtotalRow ----------
' 변경점: 첫 5개 열의 "정확 일치" 합계 텍스트만 인정.
'         화살표/"실수령액 기준" 등 주석 휴리스틱 제거 (중간 데이터 오탐 방지)
Private Function IsSubtotalRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long, v As String
    For c = 1 To 5
        v = Trim$(CStr(ws.Cells(r, c).Value))
        If v = "합계" Or v = "총합계" Or v = "소계" Or v = "누계" Then
            IsSubtotalRow = True
            Exit Function
        End If
    Next c
End Function

' ---------- 3) 문제2: ExtractPlatformFromBigo (플랫폼 사전 스캔) ----------
Private Function ExtractPlatformFromBigo(ByVal bigo As String, ByVal gubun As String) As String
    Dim s As String, i As Long
    Dim known As Variant
    known = Array( _
        "네이버시리즈", "네이버웹툰", "네이버", _
        "카카오페이지", "카카오웹툰", "카카오", _
        "리디북스", "리디", _
        "블라이스", "레드피치", "테라핀", _
        "스토리위즈", "원스토리", "문피아", "조아라", _
        "밀리의서재", "밀리", "교보문고", "교보", _
        "예스24", "예스", "알라딘", "북큐브", _
        "미스터블루", "코핀커뮤니케이션즈", "코핀재팬", "코핀", _
        "탑툰", "투믹스", "봄툰", "피너툰", "레진", _
        "북팔", "조이라이드", "스토리잇", "버프툰" _
    )
    s = Trim$(bigo)
    If Len(s) = 0 Then
        ExtractPlatformFromBigo = "매입(" & gubun & ")"
        Exit Function
    End If

    ' 괄호 안 우선 (예: "스토리위즈(블라이스) ..." → 블라이스)
    Dim p1 As Long, p2 As Long, inParen As String
    p1 = InStr(s, "("): p2 = InStr(s, ")")
    If p1 > 0 And p2 > p1 Then
        inParen = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
        For i = LBound(known) To UBound(known)
            If InStr(1, inParen, CStr(known(i)), vbTextCompare) > 0 Then
                ExtractPlatformFromBigo = CStr(known(i))
                Exit Function
            End If
        Next i
    End If

    ' 본문 스캔 (긴 이름 우선 순서로 배열 구성함)
    For i = LBound(known) To UBound(known)
        If InStr(1, s, CStr(known(i)), vbTextCompare) > 0 Then
            ExtractPlatformFromBigo = CStr(known(i))
            Exit Function
        End If
    Next i

    ExtractPlatformFromBigo = "매입(" & gubun & ")"
End Function

' ---------- 4) 신규: ParseBigoGrossAndRS ----------
' 비고 예시:
'   "블라이스 50만원*0.7"          → gross=500000, rs=0.7
'   "MG_네이버 1천*0.75"            → gross=1000000, rs=0.75 (천=천원? → 천=1000000 관례)
'   "MG_레드피치 웹툰선인세 200만원*0.7"
'   "네이버 2천만원*0.8"
' 반환: grossOut(총매출), rsOut(작가RS 0~1). 파싱 실패시 False.
Private Function ParseBigoGrossAndRS(ByVal bigo As String, _
                                     ByRef grossOut As Double, _
                                     ByRef rsOut As Double) As Boolean
    Dim s As String, i As Long, ch As String
    Dim numStart As Long, numStr As String, num As Double
    Dim unit As String, hasUnit As Boolean
    Dim rsPos As Long, rsStr As String

    s = Replace(bigo, " ", "")
    s = Replace(s, ",", "")

    ' --- RS 파싱: "*0.7" 또는 "*70%" ---
    rsPos = InStr(s, "*")
    If rsPos = 0 Then Exit Function
    rsStr = Mid$(s, rsPos + 1)
    ' 뒤에 다른 문자가 붙어있을 수 있으므로 숫자+.+% 만 추출
    Dim j As Long, rsClean As String
    For j = 1 To Len(rsStr)
        ch = Mid$(rsStr, j, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "%" Then
            rsClean = rsClean & ch
        Else
            Exit For
        End If
    Next j
    If Len(rsClean) = 0 Then Exit Function
    If Right$(rsClean, 1) = "%" Then
        rsOut = CDbl(Left$(rsClean, Len(rsClean) - 1)) / 100#
    Else
        rsOut = CDbl(rsClean)
        If rsOut > 1 Then rsOut = rsOut / 100#
    End If
    If rsOut <= 0 Or rsOut > 1 Then Exit Function

    ' --- gross 파싱: "*" 바로 앞의 "숫자+단위" ---
    ' 예: "50만원", "200만원", "1천", "1천만원", "2천만원"
    Dim head As String
    head = Left$(s, rsPos - 1)
    ' 끝에서부터 숫자/단위 토큰 역추출
    Dim endPos As Long
    endPos = Len(head)
    ' 단위 판정: "원"으로 끝나면 한 글자 더 포함
    Dim tail As String, unitLen As Long, tokenLen As Long
    unitLen = 0
    If Right$(head, 1) = "원" Then unitLen = 1
    ' "만원"/"천원"/"억원"
    If unitLen = 1 And Len(head) >= 2 Then
        Dim prev As String
        prev = Mid$(head, Len(head) - 1, 1)
        If prev = "만" Or prev = "천" Or prev = "억" Then unitLen = 2
    ElseIf Len(head) >= 1 Then
        ' "원" 없이 "만"/"천"/"억"으로 끝나는 경우
        Dim lastCh As String
        lastCh = Right$(head, 1)
        If lastCh = "만" Or lastCh = "천" Or lastCh = "억" Then unitLen = 1
    End If

    If unitLen = 0 Then Exit Function
    unit = Right$(head, unitLen)
    Dim numPart As String
    numPart = Left$(head, Len(head) - unitLen)
    ' numPart 끝에서부터 숫자+. 만 역추출
    Dim k As Long, digits As String
    For k = Len(numPart) To 1 Step -1
        ch = Mid$(numPart, k, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            digits = ch & digits
        Else
            Exit For
        End If
    Next k
    If Len(digits) = 0 Then Exit Function
    num = CDbl(digits)

    ' 단위 환산
    Select Case unit
        Case "원": grossOut = num
        Case "만", "만원": grossOut = num * 10000#
        Case "천", "천원": grossOut = num * 1000000#     ' 관례: "1천"=100만원
        Case "억", "억원": grossOut = num * 100000000#
        Case Else: Exit Function
    End Select

    If grossOut > 0 Then ParseBigoGrossAndRS = True
End Function

' ---------- 5) 문제3: WriteRawDataRow (금액 분해) ----------
Private Sub WriteRawDataRow(ByVal wsSrc As Worksheet, ByVal srcRow As Long, _
                            ByVal wsOut As Worksheet, ByVal outRow As Long)
    Dim rawGubun As String, rawBigo As String, rawAmount As Double
    Dim gross As Double, rs As Double, authorAmt As Double, netProfit As Double
    Dim parsed As Boolean

    rawGubun = Trim$(CStr(SafeReadCell(wsSrc, srcRow, 7)))
    rawBigo = CStr(SafeReadCell(wsSrc, srcRow, 12))
    rawAmount = NzNumber(SafeReadCell(wsSrc, srcRow, 11))   ' K열 = 작가 지급액(세전)

    parsed = ParseBigoGrossAndRS(rawBigo, gross, rs)
    If parsed Then
        ' gross가 비고 기준. 작가 지급액은 원본 K(세전) 우선, 없으면 gross*rs
        If rawAmount > 0 Then
            authorAmt = rawAmount
        Else
            authorAmt = gross * rs
        End If
        netProfit = gross - authorAmt
    Else
        ' 파싱 실패 시: K열(세전지급액) 하나만 존재. 역산 불가 → 기존 동작 유지
        gross = rawAmount
        rs = 0
        authorAmt = rawAmount
        netProfit = 0
    End If

    ' B열 플랫폼명 ← 비고 사전 스캔
    wsOut.Cells(outRow, 2).Value = ExtractPlatformFromBigo(rawBigo, rawGubun)
    ' E 서비스월 ← raw B(기준월)
    wsOut.Cells(outRow, 5).Value = SafeReadCell(wsSrc, srcRow, 2)
    ' F 회계귀속월 ← raw E
    wsOut.Cells(outRow, 6).Value = SafeReadCell(wsSrc, srcRow, 5)
    ' G 필명 / H 저자명 ← raw H
    wsOut.Cells(outRow, 7).Value = SafeReadCell(wsSrc, srcRow, 8)
    wsOut.Cells(outRow, 8).Value = SafeReadCell(wsSrc, srcRow, 8)
    ' I 작품코드 / J 작품명
    wsOut.Cells(outRow, 9).Value = SafeReadCell(wsSrc, srcRow, 3)
    wsOut.Cells(outRow, 10).Value = SafeReadCell(wsSrc, srcRow, 4)
    ' K 수익모델 ← 매입 구분 매핑
    wsOut.Cells(outRow, 11).Value = MapRawGubunToModel(rawGubun)
    ' P 구분 / Q 유형
    wsOut.Cells(outRow, 16).Value = rawGubun
    wsOut.Cells(outRow, 17).Value = rawGubun
    ' S 플랫폼 총매출 / U 정산기준매출 / Y 테라핀 순매출 / AA 실수령액 ← gross
    wsOut.Cells(outRow, 19).Value = gross
    wsOut.Cells(outRow, 21).Value = gross
    wsOut.Cells(outRow, 25).Value = gross
    wsOut.Cells(outRow, 27).Value = gross
    ' AB 지급일자
    wsOut.Cells(outRow, 28).Value = SafeReadCell(wsSrc, srcRow, 6)
    ' AC 작가 RS(%)
    If rs > 0 Then wsOut.Cells(outRow, 29).Value = rs
    ' AD 작가 금액(세전)
    wsOut.Cells(outRow, 30).Value = authorAmt
    ' AE 테라핀 순이익
    wsOut.Cells(outRow, 31).Value = netProfit
    ' AF 비고 (파싱 성공 여부 태그 포함)
    wsOut.Cells(outRow, 32).Value = rawBigo & " [매입Raw:" & rawGubun & _
                                    IIf(parsed, "/gross=" & Format$(gross, "#,##0"), "/파싱실패") & "]"
End Sub
