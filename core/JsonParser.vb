' ===== Core/JsonParser.vb =====
' JSON 파싱 + 체결시간 변환 — 수정 불필요

Public Class JsonParser

    ' ══════════════════════════════════════════
    '  JSON → DataTable (분봉/틱봉 공통)
    ' ══════════════════════════════════════════
    Public Shared Function ParseCandles(json As String) As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("시간", GetType(String))
        dt.Columns.Add("시가", GetType(Long))
        dt.Columns.Add("고가", GetType(Long))
        dt.Columns.Add("저가", GetType(Long))
        dt.Columns.Add("종가", GetType(Long))
        dt.Columns.Add("거래량", GetType(Long))
        dt.Columns.Add("등락%", GetType(Decimal))

        Dim dataStart = json.IndexOf("[")
        Dim dataEnd = json.LastIndexOf("]")
        If dataStart < 0 OrElse dataEnd < 0 Then
            Throw New Exception("API 응답에 데이터 배열이 없습니다." & vbCrLf &
                                json.Substring(0, Math.Min(300, json.Length)))
        End If

        Dim arrayStr = json.Substring(dataStart + 1, dataEnd - dataStart - 1)
        Dim objects = SplitJsonObjects(arrayStr)

        If objects.Count = 0 Then Return dt

        ' 필드명 자동 감지
        Dim first = objects(0)
        Dim fTime = DetectField(first, {"체결시간", "time", "Time", "날짜", "date", "datetime"})
        Dim fOpen = DetectField(first, {"시가", "open", "Open", "openPrice"})
        Dim fHigh = DetectField(first, {"고가", "high", "High", "highPrice"})
        Dim fLow = DetectField(first, {"저가", "low", "Low", "lowPrice"})
        Dim fClose = DetectField(first, {"현재가", "종가", "close", "Close", "price", "closePrice"})
        Dim fVol = DetectField(first, {"거래량", "volume", "Volume", "vol"})

        If String.IsNullOrEmpty(fOpen) OrElse String.IsNullOrEmpty(fClose) Then
            Throw New Exception("JSON 필드명 감지 실패." & vbCrLf & "객체: " & first)
        End If

        For Each obj In objects
            Dim timeVal = FormatKiwoomTime(ExtractString(obj, fTime))
            Dim o = ExtractLong(obj, fOpen)
            Dim h = ExtractLong(obj, fHigh)
            Dim l = ExtractLong(obj, fLow)
            Dim c = ExtractLong(obj, fClose)
            Dim v = ExtractLong(obj, fVol)
            Dim pct As Decimal = If(o > 0, Math.Round(CDec(c - o) / CDec(o) * 100, 2), 0D)

            dt.Rows.Add(timeVal, o, h, l, c, v, pct)
        Next

        Return dt
    End Function

    ' ── {} 객체 분리 ──
    Private Shared Function SplitJsonObjects(arrayStr As String) As List(Of String)
        Dim objects As New List(Of String)
        Dim depth = 0, objStart = -1
        For i = 0 To arrayStr.Length - 1
            If arrayStr(i) = "{"c Then
                If depth = 0 Then objStart = i
                depth += 1
            ElseIf arrayStr(i) = "}"c Then
                depth -= 1
                If depth = 0 AndAlso objStart >= 0 Then
                    objects.Add(arrayStr.Substring(objStart, i - objStart + 1))
                    objStart = -1
                End If
            End If
        Next
        Return objects
    End Function

    ' ── 필드명 감지 ──
    Private Shared Function DetectField(jsonObj As String, candidates() As String) As String
        ' 정확 매칭
        For Each c In candidates
            If jsonObj.IndexOf($"""{c}""") >= 0 Then Return c
        Next
        ' 대소문자 무시
        For Each c In candidates
            Dim idx = jsonObj.IndexOf($"""{c}""", StringComparison.OrdinalIgnoreCase)
            If idx >= 0 Then Return jsonObj.Substring(idx + 1, c.Length)
        Next
        Return ""
    End Function

    ' ── 문자열 값 추출 ──
    Public Shared Function ExtractString(json As String, key As String) As String
        If String.IsNullOrEmpty(key) Then Return ""
        Dim pattern = $"""{key}"""
        Dim idx = json.IndexOf(pattern)
        If idx < 0 Then idx = json.IndexOf(pattern, StringComparison.OrdinalIgnoreCase)
        If idx < 0 Then Return ""

        Dim colonIdx = json.IndexOf(":"c, idx + pattern.Length)
        If colonIdx < 0 Then Return ""

        Dim valStart = colonIdx + 1
        While valStart < json.Length AndAlso " 	".IndexOf(json(valStart)) >= 0
            valStart += 1
        End While
        If valStart >= json.Length Then Return ""

        If json(valStart) = """"c Then
            Dim valEnd = json.IndexOf(""""c, valStart + 1)
            If valEnd < 0 Then Return ""
            Return json.Substring(valStart + 1, valEnd - valStart - 1)
        ElseIf json(valStart) = "n"c Then
            Return ""
        Else
            Dim valEnd = valStart
            While valEnd < json.Length AndAlso ","c <> json(valEnd) AndAlso "}"c <> json(valEnd)
                valEnd += 1
            End While
            Return json.Substring(valStart, valEnd - valStart).Trim()
        End If
    End Function

    ' ── Long 값 추출 (부호·쉼표 포함 문자열 대응) ──
    Public Shared Function ExtractLong(json As String, key As String) As Long
        Dim s = ExtractString(json, key)
        If String.IsNullOrEmpty(s) Then Return 0
        s = s.Replace(",", "").Replace(" ", "").Trim()

        Dim neg = False
        If s.StartsWith("-") Then : neg = True : s = s.Substring(1)
        ElseIf s.StartsWith("+") Then : s = s.Substring(1) : End If

        Dim result As Long
        If Long.TryParse(s, result) Then Return If(neg, -result, result)
        Dim d As Double
        If Double.TryParse(s, d) Then Return CLng(If(neg, -d, d))
        Return 0
    End Function

    ' ── 키움 체결시간 변환 ──
    Public Shared Function FormatKiwoomTime(raw As String) As String
        If String.IsNullOrEmpty(raw) Then Return raw
        Dim digits = raw.Replace("-", "").Replace(":", "").Replace(" ", "").Trim()
        If digits.Length = 14 Then
            Return $"{digits.Substring(0, 4)}-{digits.Substring(4, 2)}-{digits.Substring(6, 2)} {digits.Substring(8, 2)}:{digits.Substring(10, 2)}:{digits.Substring(12, 2)}"
        ElseIf digits.Length = 12 Then
            Return $"{digits.Substring(0, 4)}-{digits.Substring(4, 2)}-{digits.Substring(6, 2)} {digits.Substring(8, 2)}:{digits.Substring(10, 2)}"
        ElseIf digits.Length = 8 Then
            Return $"{digits.Substring(0, 4)}-{digits.Substring(4, 2)}-{digits.Substring(6, 2)}"
        End If
        Return raw
    End Function

    ' ── 클립보드 % 파싱 ──
    Public Shared Function ParsePct(raw As String) As Object
        If String.IsNullOrEmpty(raw) Then Return DBNull.Value
        Dim s = raw.Replace("""", "").Replace("%", "").Replace("+", "").Replace(",", "").Trim()
        If String.IsNullOrEmpty(s) Then Return DBNull.Value
        Dim result As Decimal
        If Decimal.TryParse(s, result) Then Return result
        Return DBNull.Value
    End Function

    ' ── 클립보드 거래량 파싱 ──
    Public Shared Function ParseVolume(raw As String) As Object
        If String.IsNullOrEmpty(raw) Then Return DBNull.Value
        Dim s = raw.Replace("""", "").Replace(",", "").Trim()
        If String.IsNullOrEmpty(s) Then Return DBNull.Value
        Dim result As Long
        If Long.TryParse(s, result) Then Return result
        Return DBNull.Value
    End Function

End Class
