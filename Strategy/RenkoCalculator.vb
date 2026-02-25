Imports System.Math

Public Class RenkoBrick
    Public Property Direction As Integer   ' 1=상승, -1=하락
    Public Property OpenPrice As Double
    Public Property ClosePrice As Double
    Public Property StartTime As String
    Public Property EndTime As String
    Public Property BarCount As Integer    ' 이 벽돌을 만드는데 소요된 봉 수
End Class

Public Class RenkoResult
    Public Property Bricks As New List(Of RenkoBrick)
    Public Property ConsecutiveUp As Integer = 0       ' 마지막 연속 상승 벽돌 수
    Public Property ConsecutiveDown As Integer = 0
    Public Property BullRatio As Double = 0            ' 최근 N개 중 상승 비율
    Public Property Speed As Double = 0                ' 시간당 벽돌 생성 수
    Public Property BrickRange As Double = 0           ' 연속 상승 구간 누적 %
    Public Property TotalBricks As Integer = 0
End Class

Public Class RenkoCalculator

    ''' <summary>
    ''' 종가 기반 렌코 벽돌 계산 (퍼센트 기반 벽돌 크기)
    ''' </summary>
    Public Shared Function Calculate(chartData As DataTable, brickPct As Double) As RenkoResult
        Dim result As New RenkoResult()
        If chartData Is Nothing OrElse chartData.Rows.Count < 2 Then Return result

        ' 종가 컬럼명 탐지
        Dim closeCol As String = ""
        For Each cn As String In {"현재가", "종가", "Close", "close"}
            If chartData.Columns.Contains(cn) Then closeCol = cn : Exit For
        Next
        If closeCol = "" Then Return result

        ' 시간 컬럼명 탐지
        Dim timeCol As String = ""
        For Each cn As String In {"체결시간", "시간", "Time", "time", "일자"}
            If chartData.Columns.Contains(cn) Then timeCol = cn : Exit For
        Next

        ' 종가 배열 추출
        Dim prices As New List(Of Double)
        Dim times As New List(Of String)
        For i As Integer = 0 To chartData.Rows.Count - 1
            Dim v As Double = 0
            Double.TryParse(chartData.Rows(i)(closeCol).ToString().Replace(",", "").Replace("+", "").Replace("-", ""), v)
            If v > 0 Then
                prices.Add(v)
                If timeCol <> "" Then times.Add(chartData.Rows(i)(timeCol).ToString()) Else times.Add(i.ToString())
            End If
        Next
        If prices.Count < 2 Then Return result

        ' 첫 번째 벽돌 기준가
        Dim brickSize As Double = prices(0) * brickPct / 100.0
        If brickSize <= 0 Then brickSize = 1
        Dim currentLevel As Double = Math.Floor(prices(0) / brickSize) * brickSize
        Dim bricks As New List(Of RenkoBrick)
        Dim lastBrickBar As Integer = 0

        For i As Integer = 1 To prices.Count - 1
            Dim price As Double = prices(i)
            ' 동적 벽돌 크기 (퍼센트 기반이므로 현재 레벨에 따라 변동)
            brickSize = currentLevel * brickPct / 100.0
            If brickSize <= 0 Then brickSize = 1

            ' 상승 벽돌
            While price >= currentLevel + brickSize
                Dim b As New RenkoBrick()
                b.Direction = 1
                b.OpenPrice = currentLevel
                currentLevel += brickSize
                b.ClosePrice = currentLevel
                b.StartTime = If(lastBrickBar < times.Count, times(lastBrickBar), "")
                b.EndTime = If(i < times.Count, times(i), "")
                b.BarCount = i - lastBrickBar
                bricks.Add(b)
                lastBrickBar = i
                brickSize = currentLevel * brickPct / 100.0
                If brickSize <= 0 Then brickSize = 1
            End While

            ' 하락 벽돌
            While price <= currentLevel - brickSize
                Dim b As New RenkoBrick()
                b.Direction = -1
                b.OpenPrice = currentLevel
                currentLevel -= brickSize
                b.ClosePrice = currentLevel
                b.StartTime = If(lastBrickBar < times.Count, times(lastBrickBar), "")
                b.EndTime = If(i < times.Count, times(i), "")
                b.BarCount = i - lastBrickBar
                bricks.Add(b)
                lastBrickBar = i
                brickSize = currentLevel * brickPct / 100.0
                If brickSize <= 0 Then brickSize = 1
            End While
        Next

        result.Bricks = bricks
        result.TotalBricks = bricks.Count

        ' 마지막 연속 상승/하락 계산
        If bricks.Count > 0 Then
            Dim lastDir As Integer = bricks(bricks.Count - 1).Direction
            Dim cnt As Integer = 0
            For j As Integer = bricks.Count - 1 To 0 Step -1
                If bricks(j).Direction = lastDir Then cnt += 1 Else Exit For
            Next
            If lastDir = 1 Then result.ConsecutiveUp = cnt Else result.ConsecutiveDown = cnt
        End If

        ' 최근 20개 벽돌 중 상승 비율
        Dim lookback As Integer = Math.Min(20, bricks.Count)
        If lookback > 0 Then
            Dim bullCnt As Integer = 0
            For j As Integer = bricks.Count - lookback To bricks.Count - 1
                If bricks(j).Direction = 1 Then bullCnt += 1
            Next
            result.BullRatio = bullCnt / CDbl(lookback)
        End If

        ' 속도 (마지막 연속 상승 구간의 봉당 벽돌 수)
        If result.ConsecutiveUp > 0 Then
            Dim totalBars As Integer = 0
            For j As Integer = bricks.Count - result.ConsecutiveUp To bricks.Count - 1
                totalBars += bricks(j).BarCount
            Next
            If totalBars > 0 Then result.Speed = result.ConsecutiveUp / CDbl(totalBars)
        End If

        ' 연속 상승 구간 누적 수익률 (%)
        If result.ConsecutiveUp >= 1 Then
            Dim startIdx As Integer = bricks.Count - result.ConsecutiveUp
            Dim startPrice As Double = bricks(startIdx).OpenPrice
            Dim endPrice As Double = bricks(bricks.Count - 1).ClosePrice
            If startPrice > 0 Then result.BrickRange = (endPrice - startPrice) / startPrice * 100.0
        End If

        Return result
    End Function

    ''' <summary>
    ''' 여러 벽돌 크기로 계산하여 딕셔너리 반환
    ''' </summary>
    Public Shared Function CalculateMultiple(chartData As DataTable, brickPcts As Double()) As Dictionary(Of Double, RenkoResult)
        Dim dic As New Dictionary(Of Double, RenkoResult)
        For Each pct In brickPcts
            dic(pct) = Calculate(chartData, pct)
        Next
        Return dic
    End Function

    ''' <summary>
    ''' StrategyEngine의 indicators 딕셔너리에 렌코 지표 추가
    ''' 모든 봉에 동일한 값을 넣음 (렌코는 시점 독립적 요약 지표)
    ''' </summary>
    Public Shared Sub AddRenkoIndicators(chartData As DataTable, indicators As Dictionary(Of String, Double()), brickPct As Double)
        Dim renko As RenkoResult = Calculate(chartData, brickPct)
        Dim n As Integer = chartData.Rows.Count
        If n = 0 Then Exit Sub

        Dim arrConsec(n - 1) As Double
        Dim arrRatio(n - 1) As Double
        Dim arrSpeed(n - 1) As Double
        Dim arrRange(n - 1) As Double
        Dim arrTotal(n - 1) As Double

        For i As Integer = 0 To n - 1
            arrConsec(i) = renko.ConsecutiveUp
            arrRatio(i) = renko.BullRatio
            arrSpeed(i) = renko.Speed
            arrRange(i) = renko.BrickRange
            arrTotal(i) = renko.TotalBricks
        Next

        indicators("RENKO_CONSEC") = arrConsec
        indicators("RENKO_RATIO") = arrRatio
        indicators("RENKO_SPEED") = arrSpeed
        indicators("RENKO_RANGE") = arrRange
        indicators("RENKO_TOTAL") = arrTotal
    End Sub

End Class