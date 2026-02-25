' ===== Strategy/TickIntensity.vb =====
' 분봉 구간별 틱봉 개수 합산

Public Class TickIntensity

    ''' <summary>분봉 DataTable에 틱강도/틱거래량 컬럼 추가</summary>
    Public Shared Sub Calculate(dtMinute As DataTable, dtTick As DataTable, minuteInterval As Integer)
        If dtMinute Is Nothing OrElse dtTick Is Nothing Then Return

        If Not dtMinute.Columns.Contains("틱강도") Then dtMinute.Columns.Add("틱강도", GetType(Integer))
        If Not dtMinute.Columns.Contains("틱거래량") Then dtMinute.Columns.Add("틱거래량", GetType(Long))

        ' 틱 시간/거래량 리스트
        Dim tickTimes As New List(Of DateTime)
        Dim tickVols As New List(Of Long)
        For Each row As DataRow In dtTick.Rows
            Dim ts = row("시간").ToString()
            Dim dt2 As DateTime
            If DateTime.TryParse(ts, dt2) Then
                tickTimes.Add(dt2)
                Dim vol As Long = 0
                If row("거래량") IsNot Nothing AndAlso Not IsDBNull(row("거래량")) Then vol = CLng(row("거래량"))
                tickVols.Add(vol)
            End If
        Next

        For Each mRow As DataRow In dtMinute.Rows
            Dim endTimeStr = mRow("시간").ToString()
            Dim endTime As DateTime
            If Not DateTime.TryParse(endTimeStr, endTime) Then
                mRow("틱강도") = 0 : mRow("틱거래량") = 0L : Continue For
            End If
            Dim startTime = endTime.AddMinutes(-minuteInterval)
            Dim cnt = 0, vol As Long = 0
            For i = 0 To tickTimes.Count - 1
                If tickTimes(i) > startTime AndAlso tickTimes(i) <= endTime Then
                    cnt += 1 : vol += tickVols(i)
                End If
            Next
            mRow("틱강도") = cnt : mRow("틱거래량") = vol
        Next
    End Sub

    ''' <summary>틱강도 요약 문자열</summary>
    Public Shared Function GetSummary(dtMinute As DataTable) As String
        If dtMinute Is Nothing OrElse Not dtMinute.Columns.Contains("틱강도") Then Return ""
        Dim total = 0, cnt = 0, maxV = 0, minV = Integer.MaxValue, cnt20 = 0, cnt10 = 0
        For Each row As DataRow In dtMinute.Rows
            If row("틱강도") Is Nothing OrElse IsDBNull(row("틱강도")) Then Continue For
            Dim v = CInt(row("틱강도"))
            total += v : cnt += 1
            If v > maxV Then maxV = v
            If v < minV Then minV = v
            If v >= 20 Then cnt20 += 1
            If v >= 10 Then cnt10 += 1
        Next
        If cnt = 0 Then Return "틱강도: 데이터 없음"
        Dim avg = total / CDbl(cnt)
        Return $"틱강도: 평균 {avg:N1}, 최대 {maxV}, 최소 {minV} | ≥20: {cnt20}건, ≥10: {cnt10}건 / {cnt}건"
    End Function

End Class
