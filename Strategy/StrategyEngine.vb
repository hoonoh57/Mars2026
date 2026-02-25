' ===== Strategy/StrategyEngine.vb =====
' 전략 조건 평가 + 백테스트 엔진
' TurnUp/TurnDown/Rising/Falling + (N) 이전봉 참조 지원

Public Class StrategyEngine

    ''' <summary>차트 데이터에 전략 적용 → 신호 + 성과</summary>
    Public Shared Function Evaluate(strategy As TradingStrategy,
                                     chartData As DataTable) As StrategyPerformance
        Dim perf As New StrategyPerformance()
        If chartData Is Nothing OrElse chartData.Rows.Count < 3 Then Return perf

        Dim count = chartData.Rows.Count
        Dim rows = chartData.Rows

        ' 지표 빌드
        Dim indicators As New Dictionary(Of String, Double())
        BuildIndicators(chartData, indicators)

        ' 시뮬레이션
        Dim inPosition = False
        Dim entryIdx = 0, entryPrice As Double = 0

        For i = 2 To count - 1  ' 최소 2봉 이전 참조 필요
            If Not inPosition Then
                If EvalConditions(strategy.BuyConditions, indicators, i, chartData) Then
                    inPosition = True : entryIdx = i
                    entryPrice = CDbl(rows(i)("종가"))
                    perf.Signals.Add(New TradeSignal() With {
                        .BarIndex = i, .Type = SignalType.Buy, .Price = entryPrice,
                        .Time = rows(i)("시간").ToString(), .Reason = "매수조건 충족"})
                End If
            Else
                Dim curPrice = CDbl(rows(i)("종가"))
                Dim retPct = (curPrice - entryPrice) / entryPrice * 100
                Dim holdBars = i - entryIdx
                Dim doExit = False, exitReason = ""

                ' 손절
                If retPct <= strategy.StopLossPct Then
                    doExit = True : exitReason = $"손절({retPct:N2}%)"
                    perf.Signals.Add(New TradeSignal() With {
                        .BarIndex = i, .Type = SignalType.StopLoss, .Price = curPrice,
                        .Time = rows(i)("시간").ToString(), .Reason = exitReason})
                End If

                ' 익절
                If Not doExit AndAlso retPct >= strategy.TakeProfitPct Then
                    doExit = True : exitReason = $"익절({retPct:N2}%)"
                    perf.Signals.Add(New TradeSignal() With {
                        .BarIndex = i, .Type = SignalType.TakeProfit, .Price = curPrice,
                        .Time = rows(i)("시간").ToString(), .Reason = exitReason})
                End If

                ' 최대 보유
                If Not doExit AndAlso holdBars >= strategy.MaxHoldBars Then
                    doExit = True : exitReason = $"보유한도({holdBars}봉)"
                    perf.Signals.Add(New TradeSignal() With {
                        .BarIndex = i, .Type = SignalType.Sell, .Price = curPrice,
                        .Time = rows(i)("시간").ToString(), .Reason = exitReason})
                End If

                ' 매도 조건
                If Not doExit AndAlso strategy.SellConditions.Count > 0 AndAlso
                   EvalConditions(strategy.SellConditions, indicators, i, chartData) Then
                    doExit = True : exitReason = "매도조건 충족"
                    perf.Signals.Add(New TradeSignal() With {
                        .BarIndex = i, .Type = SignalType.Sell, .Price = curPrice,
                        .Time = rows(i)("시간").ToString(), .Reason = exitReason})
                End If

                If doExit Then
                    perf.TradeDetails.Add(New TradeDetail() With {
                        .EntryIndex = entryIdx, .ExitIndex = i,
                        .EntryPrice = entryPrice, .ExitPrice = curPrice,
                        .ReturnPct = retPct, .HoldBars = holdBars,
                        .EntryTime = rows(entryIdx)("시간").ToString(),
                        .ExitTime = rows(i)("시간").ToString(), .ExitReason = exitReason})
                    inPosition = False
                End If
            End If
        Next

        ' 미청산
        If inPosition Then
            Dim lp = CDbl(rows(count - 1)("종가"))
            perf.TradeDetails.Add(New TradeDetail() With {
                .EntryIndex = entryIdx, .ExitIndex = count - 1,
                .EntryPrice = entryPrice, .ExitPrice = lp,
                .ReturnPct = (lp - entryPrice) / entryPrice * 100,
                .HoldBars = count - 1 - entryIdx,
                .EntryTime = rows(entryIdx)("시간").ToString(),
                .ExitTime = rows(count - 1)("시간").ToString(), .ExitReason = "미청산"})
        End If

        ' 성과 집계
        CalcPerformance(perf)
        Return perf
    End Function

    ' ── 성과 집계 ──
    Private Shared Sub CalcPerformance(perf As StrategyPerformance)
        perf.TotalTrades = perf.TradeDetails.Count
        perf.WinTrades = perf.TradeDetails.AsEnumerable.Where(Function(t) t.ReturnPct > 0).Count
        perf.LossTrades = perf.TradeDetails.AsEnumerable.Where(Function(t) t.ReturnPct <= 0).Count
        perf.WinRate = If(perf.TotalTrades > 0, perf.WinTrades / CDbl(perf.TotalTrades) * 100, 0)
        perf.TotalReturnPct = perf.TradeDetails.Sum(Function(t) t.ReturnPct)
        perf.AvgReturnPct = If(perf.TotalTrades > 0, perf.TotalReturnPct / perf.TotalTrades, 0)
        perf.AvgHoldBars = If(perf.TotalTrades > 0, perf.TradeDetails.Average(Function(t) CDbl(t.HoldBars)), 0)
        Dim tG = perf.TradeDetails.Where(Function(t) t.ReturnPct > 0).Sum(Function(t) t.ReturnPct)
        Dim tL = Math.Abs(perf.TradeDetails.Where(Function(t) t.ReturnPct <= 0).Sum(Function(t) t.ReturnPct))
        perf.ProfitFactor = If(tL > 0, tG / tL, If(tG > 0, 999, 0))
        Dim cum As Double = 0, pk As Double = 0, dd As Double = 0
        For Each t In perf.TradeDetails
            cum += t.ReturnPct : If cum > pk Then pk = cum
            Dim d = pk - cum : If d > dd Then dd = d
        Next
        perf.MaxDrawdownPct = dd
    End Sub

    ' ── 조건 평가 (AND) ──
    Private Shared Function EvalConditions(conditions As List(Of StrategyCondition),
                                            indicators As Dictionary(Of String, Double()),
                                            idx As Integer,
                                            chartData As DataTable) As Boolean
        If conditions.Count = 0 Then Return False
        For Each cond In conditions
            If Not EvalSingle(cond, indicators, idx, chartData) Then Return False
        Next
        Return True
    End Function

    ' ── 개별 조건 평가 ──
    Private Shared Function EvalSingle(cond As StrategyCondition,
                                        indicators As Dictionary(Of String, Double()),
                                        idx As Integer,
                                        chartData As DataTable) As Boolean

        Dim op = cond.Operator.Trim()

        Select Case op.ToUpper()

            Case "TURNUP"
                ' 지표가 하락→상승 전환 (prev2 >= prev1 AND prev1 < current)
                Dim cur = ResolveValue(cond.Indicator, indicators, idx, chartData)
                Dim prev1 = ResolveValue(cond.Indicator, indicators, idx - 1, chartData)
                Dim prev2 = ResolveValue(cond.Indicator, indicators, idx - 2, chartData)
                If Double.IsNaN(cur) OrElse Double.IsNaN(prev1) OrElse Double.IsNaN(prev2) Then Return False
                Return prev2 >= prev1 AndAlso cur > prev1

            Case "TURNDOWN"
                ' 지표가 상승→하락 전환 (prev2 <= prev1 AND prev1 > current)
                Dim cur = ResolveValue(cond.Indicator, indicators, idx, chartData)
                Dim prev1 = ResolveValue(cond.Indicator, indicators, idx - 1, chartData)
                Dim prev2 = ResolveValue(cond.Indicator, indicators, idx - 2, chartData)
                If Double.IsNaN(cur) OrElse Double.IsNaN(prev1) OrElse Double.IsNaN(prev2) Then Return False
                Return prev2 <= prev1 AndAlso cur < prev1

            Case "RISING"
                ' N봉 연속 상승
                Dim n = CInt(cond.Value)
                If n < 1 Then n = 2
                If idx < n Then Return False
                For j = idx - n + 1 To idx
                    Dim cur = ResolveValue(cond.Indicator, indicators, j, chartData)
                    Dim prev = ResolveValue(cond.Indicator, indicators, j - 1, chartData)
                    If Double.IsNaN(cur) OrElse Double.IsNaN(prev) Then Return False
                    If cur <= prev Then Return False
                Next
                Return True

            Case "FALLING"
                ' N봉 연속 하락
                Dim n = CInt(cond.Value)
                If n < 1 Then n = 2
                If idx < n Then Return False
                For j = idx - n + 1 To idx
                    Dim cur = ResolveValue(cond.Indicator, indicators, j, chartData)
                    Dim prev = ResolveValue(cond.Indicator, indicators, j - 1, chartData)
                    If Double.IsNaN(cur) OrElse Double.IsNaN(prev) Then Return False
                    If cur >= prev Then Return False
                Next
                Return True

            Case "CROSSUP"
                ' left가 right를 상향 돌파
                Dim leftCur = ResolveValue(cond.Indicator, indicators, idx, chartData)
                Dim leftPrev = ResolveValue(cond.Indicator, indicators, idx - 1, chartData)
                Dim rightCur = ResolveRight(cond, indicators, idx, chartData)
                Dim rightPrev = ResolveRight(cond, indicators, idx - 1, chartData)
                If Double.IsNaN(leftCur) OrElse Double.IsNaN(leftPrev) OrElse
                   Double.IsNaN(rightCur) OrElse Double.IsNaN(rightPrev) Then Return False
                Return leftPrev <= rightPrev AndAlso leftCur > rightCur

            Case "CROSSDOWN"
                ' left가 right를 하향 돌파
                Dim leftCur = ResolveValue(cond.Indicator, indicators, idx, chartData)
                Dim leftPrev = ResolveValue(cond.Indicator, indicators, idx - 1, chartData)
                Dim rightCur = ResolveRight(cond, indicators, idx, chartData)
                Dim rightPrev = ResolveRight(cond, indicators, idx - 1, chartData)
                If Double.IsNaN(leftCur) OrElse Double.IsNaN(leftPrev) OrElse
                   Double.IsNaN(rightCur) OrElse Double.IsNaN(rightPrev) Then Return False
                Return leftPrev >= rightPrev AndAlso leftCur < rightCur

            Case ">", "<", ">=", "<=", "=", "=="
                Dim leftVal = ResolveValue(cond.Indicator, indicators, idx, chartData)
                Dim rightVal = ResolveRight(cond, indicators, idx, chartData)
                If Double.IsNaN(leftVal) OrElse Double.IsNaN(rightVal) Then Return False
                Select Case op
                    Case ">" : Return leftVal > rightVal
                    Case "<" : Return leftVal < rightVal
                    Case ">=" : Return leftVal >= rightVal
                    Case "<=" : Return leftVal <= rightVal
                    Case "=", "==" : Return Math.Abs(leftVal - rightVal) < 0.001
                End Select

            Case Else
                Return False
        End Select

        Return False
    End Function

    ' ── 이름에서 값 해석 (이전 봉 참조 지원) ──
    ' "JMA" → 현재 봉 JMA, "JMA(1)" → 1봉 전 JMA, "JMA(2)" → 2봉 전
    Private Shared Function ResolveValue(name As String,
                                          indicators As Dictionary(Of String, Double()),
                                          idx As Integer,
                                          chartData As DataTable) As Double
        If String.IsNullOrEmpty(name) Then Return Double.NaN

        ' 숫자인 경우
        Dim numVal As Double
        If Double.TryParse(name.Trim(), numVal) Then Return numVal

        ' (N) 접미사 파싱
        Dim baseName = name.Trim()
        Dim offset = 0
        If baseName.EndsWith(")") Then
            Dim pOpen = baseName.LastIndexOf("("c)
            If pOpen > 0 Then
                Dim offStr = baseName.Substring(pOpen + 1, baseName.Length - pOpen - 2)
                If Integer.TryParse(offStr, offset) Then
                    baseName = baseName.Substring(0, pOpen)
                End If
            End If
        End If

        Dim actualIdx = idx - offset
        If actualIdx < 0 Then Return Double.NaN

        ' 지표 딕셔너리 검색
        Dim key = baseName.ToUpper()
        For Each kv In indicators
            If kv.Key.ToUpper() = key Then
                If actualIdx >= 0 AndAlso actualIdx < kv.Value.Length Then Return kv.Value(actualIdx)
                Return Double.NaN
            End If
        Next

        ' DataTable 직접 참조
        If actualIdx >= chartData.Rows.Count Then Return Double.NaN
        Select Case key
            Case "PRICE", "CLOSE", "종가" : Return CDbl(chartData.Rows(actualIdx)("종가"))
            Case "OPEN", "시가" : Return CDbl(chartData.Rows(actualIdx)("시가"))
            Case "HIGH", "고가" : Return CDbl(chartData.Rows(actualIdx)("고가"))
            Case "LOW", "저가" : Return CDbl(chartData.Rows(actualIdx)("저가"))
            Case "VOLUME", "거래량"
                Dim v = chartData.Rows(actualIdx)("거래량")
                If v IsNot Nothing AndAlso Not IsDBNull(v) Then Return CDbl(CLng(v))
        End Select
        Return Double.NaN
    End Function

    ' ── 우변 해석 (Target 또는 Value) ──
    Private Shared Function ResolveRight(cond As StrategyCondition,
                                          indicators As Dictionary(Of String, Double()),
                                          idx As Integer,
                                          chartData As DataTable) As Double
        If Not String.IsNullOrEmpty(cond.Target) Then
            ' Target이 숫자인지 확인
            Dim numVal As Double
            If Double.TryParse(cond.Target.Trim(), numVal) Then Return numVal
            ' 지표 참조
            Return ResolveValue(cond.Target, indicators, idx, chartData)
        End If
        Return cond.Value
    End Function

    ' ── 지표 빌드 ──
    Private Shared Sub BuildIndicators(chartData As DataTable,
                                        indicators As Dictionary(Of String, Double()))
        Dim count = chartData.Rows.Count
        Dim closes(count - 1) As Double
        Dim highs(count - 1) As Double, lows(count - 1) As Double
        For i = 0 To count - 1
            closes(i) = CDbl(chartData.Rows(i)("종가"))
            highs(i) = CDbl(chartData.Rows(i)("고가"))
            lows(i) = CDbl(chartData.Rows(i)("저가"))
        Next

        indicators("Close") = closes
        indicators("High") = highs
        indicators("Low") = lows
        indicators("MA5") = CalcSMA(closes, 5)
        indicators("MA20") = CalcSMA(closes, 20)
        indicators("MA120") = CalcSMA(closes, 120)
        indicators("RSI14") = CalcRSI(closes, 14)
        indicators("RSI5") = CalcRSI(closes, 5)
        indicators("RSI50") = CalcRSI(closes, 50)
        indicators("JMA") = CalcJMA(closes, 7, 50, 2)

        Dim stUp As Double() = Nothing, stDn As Double() = Nothing, stDir As Integer() = Nothing
        CalcSuperTrend(highs, lows, closes, 14, 2.0, stUp, stDn, stDir)
        Dim stVal(count - 1) As Double, stDirD(count - 1) As Double
        For i = 0 To count - 1
            stVal(i) = If(stDir(i) = 1, stUp(i), stDn(i))
            stDirD(i) = stDir(i)
        Next
        indicators("SuperTrend") = stVal
        indicators("STDir") = stDirD

        If chartData.Columns.Contains("틱강도") Then
            Dim ti(count - 1) As Double
            For i = 0 To count - 1
                Dim v = chartData.Rows(i)("틱강도")
                If v IsNot Nothing AndAlso Not IsDBNull(v) Then ti(i) = CDbl(CInt(v))
            Next
            indicators("TickIntensity") = ti
        End If
    End Sub

    ' ═══════ 지표 계산 함수들 (기존 동일) ═══════

    Private Shared Function CalcSMA(data() As Double, period As Integer) As Double()
        Dim n = data.Length, result(n - 1) As Double, sum As Double = 0
        For i = 0 To n - 1
            sum += data(i) : If i >= period Then sum -= data(i - period)
            result(i) = If(i >= period - 1, sum / period, Double.NaN)
        Next : Return result
    End Function

    Private Shared Function CalcRSI(closes() As Double, period As Integer) As Double()
        Dim n = closes.Length, result(n - 1) As Double
        For i = 0 To n - 1 : result(i) = Double.NaN : Next
        If n <= period Then Return result
        Dim gS As Double = 0, lS As Double = 0
        For i = 1 To period
            Dim chg = closes(i) - closes(i - 1)
            If chg > 0 Then gS += chg Else lS += Math.Abs(chg)
        Next
        Dim aG = gS / period, aL = lS / period
        result(period) = If(aL = 0, 100, 100 - 100 / (1 + aG / aL))
        For i = period + 1 To n - 1
            Dim chg = closes(i) - closes(i - 1)
            aG = (aG * (period - 1) + If(chg > 0, chg, 0)) / period
            aL = (aL * (period - 1) + If(chg < 0, Math.Abs(chg), 0)) / period
            result(i) = If(aL = 0, 100, 100 - 100 / (1 + aG / aL))
        Next : Return result
    End Function

    Private Shared Function CalcJMA(closes() As Double, length As Integer, phase As Integer, power As Integer) As Double()
        Dim n = closes.Length, result(n - 1) As Double
        Dim phaseRatio = If(phase < -100, 0.5, If(phase > 100, 2.5, CDbl(phase) / 100.0 + 1.5))
        Dim beta = 0.45 * (length - 1) / (0.45 * (length - 1) + 2.0)
        Dim alpha = Math.Pow(beta, CDbl(power))
        Dim e0 = closes(0), e1 = 0.0, e2 = 0.0, jma = closes(0) : result(0) = closes(0)
        For i = 1 To n - 1
            e0 = (1 - alpha) * closes(i) + alpha * e0
            e1 = (closes(i) - e0) * (1 - beta) + beta * e1
            e2 = (e0 + phaseRatio * e1 - jma) * Math.Pow(1 - alpha, 2) + Math.Pow(alpha, 2) * e2
            jma += e2 : result(i) = jma
        Next : Return result
    End Function

    Private Shared Sub CalcSuperTrend(highs() As Double, lows() As Double, closes() As Double,
                                       atrP As Integer, mult As Double,
                                       ByRef upB() As Double, ByRef dnB() As Double, ByRef dir() As Integer)
        Dim n = closes.Length
        upB = New Double(n - 1) {} : dnB = New Double(n - 1) {} : dir = New Integer(n - 1) {}
        Dim tr(n - 1) As Double, atr(n - 1) As Double
        tr(0) = highs(0) - lows(0)
        For i = 1 To n - 1
            tr(i) = Math.Max(highs(i) - lows(i), Math.Max(Math.Abs(highs(i) - closes(i - 1)), Math.Abs(lows(i) - closes(i - 1))))
        Next
        Dim s As Double = 0
        For i = 0 To Math.Min(atrP - 1, n - 1) : s += tr(i) : Next
        If atrP <= n Then atr(atrP - 1) = s / atrP
        For i = atrP To n - 1 : atr(i) = (atr(i - 1) * (atrP - 1) + tr(i)) / atrP : Next
        Dim bU(n - 1) As Double, bD(n - 1) As Double, fU(n - 1) As Double, fD(n - 1) As Double
        For i = 0 To n - 1 : Dim m = (highs(i) + lows(i)) / 2 : bU(i) = m - mult * atr(i) : bD(i) = m + mult * atr(i) : Next
        fU(0) = bU(0) : fD(0) = bD(0) : dir(0) = 1 : upB(0) = fU(0) : dnB(0) = fD(0)
        For i = 1 To n - 1
            fU(i) = If(bU(i) > fU(i - 1) OrElse closes(i - 1) > fU(i - 1), Math.Max(bU(i), fU(i - 1)), bU(i))
            fD(i) = If(bD(i) < fD(i - 1) OrElse closes(i - 1) < fD(i - 1), Math.Min(bD(i), fD(i - 1)), bD(i))
            dir(i) = If(dir(i - 1) = 1, If(closes(i) < fU(i), -1, 1), If(closes(i) > fD(i), 1, -1))
            upB(i) = fU(i) : dnB(i) = fD(i)
        Next
    End Sub

End Class
