' ===== Strategy/ChartRenderer.vb =====
' 캔들차트 + 틱강도(MA5/MA20) + 가격MA + SuperTrend + JMA + RSI + 크로스헤어
' 지표 토글: IndicatorFlags 로 제어

Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class ChartRenderer

    ' ══════════════════════════════════════════
    '  지표 표시 플래그 (폼에서 토글)
    ' ══════════════════════════════════════════
    <Flags>
    Public Enum IndicatorFlags
        None = 0
        MA5 = 1
        MA20 = 2
        MA120 = 4
        SuperTrend = 8
        JMA = 16
        RSI14 = 32
        RSI5 = 64
        RSI50 = 128
        TickIntensity = 256
        TickMA = 512
    End Enum

    ' ── 클래스 상단에 추가 ──
    ' 현재 적용된 전략 성과 (폼에서 설정)
    Public Shared CurrentPerformance As StrategyPerformance = Nothing

    ' 현재 활성 지표 (기본값)
    Public Shared ActiveIndicators As IndicatorFlags =
        IndicatorFlags.MA5 Or IndicatorFlags.MA20 Or
        IndicatorFlags.TickIntensity Or IndicatorFlags.TickMA

    Public Shared Function IsActive(flag As IndicatorFlags) As Boolean
        Return (ActiveIndicators And flag) = flag
    End Function

    Public Shared Sub Toggle(flag As IndicatorFlags)
        ActiveIndicators = ActiveIndicators Xor flag
    End Sub

    ' ══════════════════════════════════════════
    '  메인 Draw
    ' ══════════════════════════════════════════
    Public Shared Sub Draw(g As Graphics, rect As Rectangle,
                           chartData As DataTable, mousePoint As Point)

        If chartData Is Nothing OrElse chartData.Rows.Count = 0 Then Return

        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Text.TextRenderingHint.ClearTypeGridFit

        Dim count = chartData.Rows.Count
        Dim rows = chartData.Rows

        ' ── RSI 표시 여부 ──
        Dim showRsi = IsActive(IndicatorFlags.RSI14) OrElse
                      IsActive(IndicatorFlags.RSI5) OrElse
                      IsActive(IndicatorFlags.RSI50)
        Dim hasTick = chartData.Columns.Contains("틱강도") AndAlso IsActive(IndicatorFlags.TickIntensity)

        ' ── 레이아웃 ──
        Dim mL = 62, mR = 22, mT = 22, mB = 28
        Dim histH = If(hasTick, 60, 0)
        Dim rsiH = If(showRsi, 55, 0)
        Dim gapH = 8

        Dim chartLeft = rect.X + mL
        Dim chartRight = rect.X + rect.Width - mR
        Dim chartWidth = chartRight - chartLeft

        Dim priceTop = rect.Y + mT
        Dim totalGaps = If(hasTick, gapH, 0) + If(showRsi, gapH, 0)
        Dim priceBottom = rect.Y + rect.Height - mB - histH - rsiH - totalGaps
        Dim priceHeight = priceBottom - priceTop

        Dim histTop = priceBottom + gapH
        Dim histBottom = histTop + histH

        Dim rsiTop = If(hasTick, histBottom + gapH, priceBottom + gapH)
        Dim rsiBottom = rsiTop + rsiH

        If priceHeight < 40 OrElse chartWidth < 40 Then Return

        Dim gap As Double = chartWidth / CDbl(count)
        Dim candleW As Double = Math.Max(1, gap * 0.7)

        Dim idxToX = Function(idx As Integer) As Integer
                         Return CInt(chartLeft + gap * idx + gap / 2)
                     End Function

        ' ── 가격 범위 ──
        Dim pMin As Long = Long.MaxValue, pMax As Long = Long.MinValue
        For i = 0 To count - 1
            Dim lo = CLng(rows(i)("저가")), hi = CLng(rows(i)("고가"))
            If lo < pMin Then pMin = lo
            If hi > pMax Then pMax = hi
        Next
        If pMin = pMax Then pMax = pMin + 1
        Dim pRange As Long = pMax - pMin

        Dim priceToY = Function(price As Double) As Integer
                           Return CInt(priceTop + priceHeight * (1.0 - (price - pMin) / pRange))
                       End Function

        ' ── 종가 배열 ──
        Dim closes(count - 1) As Double
        Dim highs(count - 1) As Double, lows(count - 1) As Double
        For i = 0 To count - 1
            closes(i) = CDbl(rows(i)("종가"))
            highs(i) = CDbl(rows(i)("고가"))
            lows(i) = CDbl(rows(i)("저가"))
        Next

        ' ── 이동평균 ──
        Dim ma5 = CalcSMA(closes, 5)
        Dim ma20 = CalcSMA(closes, 20)
        Dim ma120 = CalcSMA(closes, 120)

        ' ── SuperTrend(14,2) ──
        Dim stUp As Double() = Nothing, stDn As Double() = Nothing, stDir As Integer() = Nothing
        CalcSuperTrend(highs, lows, closes, 14, 2.0, stUp, stDn, stDir)

        ' ── JMA(7,50,2) ──
        Dim jma = CalcJMA(closes, 7, 50, 2)

        ' ── RSI ──
        Dim rsi14 = CalcRSI(closes, 14)
        Dim rsi5 = CalcRSI(closes, 5)
        Dim rsi50 = CalcRSI(closes, 50)

        ' ── 틱강도 ──
        Dim tickVals(count - 1) As Integer
        Dim maxTick As Integer = 1
        If hasTick Then
            For i = 0 To count - 1
                Dim v = rows(i)("틱강도")
                If v IsNot Nothing AndAlso Not IsDBNull(v) Then
                    tickVals(i) = CInt(v)
                    If tickVals(i) > maxTick Then maxTick = tickVals(i)
                End If
            Next
        End If
        Dim tickToY = Function(tv As Double) As Integer
                          Return CInt(histBottom - histH * tv / Math.Max(maxTick, 1))
                      End Function
        Dim tMa5 = CalcSMAFromInt(tickVals, 5)
        Dim tMa20 = CalcSMAFromInt(tickVals, 20)

        ' ═════════════════════════════════════
        '  1) 배경 그리드
        ' ═════════════════════════════════════
        Using gridPen As New Pen(Color.FromArgb(40, 40, 55), 1)
            gridPen.DashStyle = DashStyle.Dot
            For i = 0 To 5
                Dim y = priceTop + CInt(priceHeight * CDbl(i) / 5)
                g.DrawLine(gridPen, chartLeft, y, chartRight, y)
                Dim pLabel = pMax - CLng(pRange * CDbl(i) / 5)
                Using fnt As New Font("Consolas", 7)
                    g.DrawString(pLabel.ToString("N0"), fnt, Brushes.Gray, rect.X + 2, y - 6)
                End Using
            Next
        End Using

        ' ═════════════════════════════════════
        '  2) 캔들
        ' ═════════════════════════════════════
        For i = 0 To count - 1
            Dim o = CLng(rows(i)("시가")), h = CLng(rows(i)("고가"))
            Dim l = CLng(rows(i)("저가")), c = CLng(rows(i)("종가"))
            If o = 0 AndAlso c = 0 Then Continue For
            Dim cx = idxToX(i)
            Dim yO = priceToY(o), yC = priceToY(c), yH = priceToY(h), yL = priceToY(l)
            Dim isUp = (c >= o)
            Dim cc = If(isUp, Color.FromArgb(220, 60, 60), Color.FromArgb(60, 60, 220))
            Using pen As New Pen(cc, 1) : g.DrawLine(pen, cx, yH, cx, yL) : End Using
            Dim bTop = Math.Min(yO, yC), bH = Math.Max(Math.Abs(yO - yC), 1)
            Using br As New SolidBrush(cc)
                g.FillRectangle(br, CInt(cx - candleW / 2), bTop, CInt(candleW), bH)
            End Using
        Next

        ' ═════════════════════════════════════
        '  3) 가격 지표 오버레이
        ' ═════════════════════════════════════
        If IsActive(IndicatorFlags.MA5) Then DrawLine(g, ma5, idxToX, priceToY, Color.FromArgb(255, 200, 50), 1.5F, count)
        If IsActive(IndicatorFlags.MA20) Then DrawLine(g, ma20, idxToX, priceToY, Color.FromArgb(100, 200, 255), 1.5F, count)
        If IsActive(IndicatorFlags.MA120) Then DrawLine(g, ma120, idxToX, priceToY, Color.FromArgb(200, 100, 255), 1.5F, count)

        ' SuperTrend
        If IsActive(IndicatorFlags.SuperTrend) AndAlso stDir IsNot Nothing Then
            For i = 1 To count - 1
                If stDir(i) = 0 Then Continue For
                Dim val1 = If(stDir(i) = 1, stUp(i - 1), stDn(i - 1))
                Dim val2 = If(stDir(i) = 1, stUp(i), stDn(i))
                If Double.IsNaN(val1) OrElse Double.IsNaN(val2) Then Continue For
                If val1 <= 0 OrElse val2 <= 0 Then Continue For
                Dim col = If(stDir(i) = 1, Color.FromArgb(0, 200, 100), Color.FromArgb(255, 80, 80))
                Using pen As New Pen(col, 2.0F)
                    g.DrawLine(pen, idxToX(i - 1), priceToY(val1), idxToX(i), priceToY(val2))
                End Using
            Next
        End If

        ' JMA
        'If IsActive(IndicatorFlags.JMA) Then DrawLine(g, jma, idxToX, priceToY, Color.FromArgb(255, 100, 200), 2.0F, count)
        ' JMA (상승=초록, 하락=분홍)
        If IsActive(IndicatorFlags.JMA) Then
            For i = 1 To count - 1
                If Double.IsNaN(jma(i)) OrElse jma(i) <= 0 OrElse Double.IsNaN(jma(i - 1)) OrElse jma(i - 1) <= 0 Then Continue For
                Dim col = If(jma(i) >= jma(i - 1),
                             Color.FromArgb(0, 220, 120),
                             Color.FromArgb(255, 80, 180))
                Using pen As New Pen(col, 2.0F)
                    g.DrawLine(pen, idxToX(i - 1), priceToY(jma(i - 1)), idxToX(i), priceToY(jma(i)))
                End Using
            Next
        End If

        ' ═════════════════════════════════════
        '  4) 틱강도 히스토그램
        ' ═════════════════════════════════════
        If hasTick Then
            ' 기준선
            For Each bl In {5, 10, 20}
                If bl <= maxTick Then
                    Dim y = tickToY(bl)
                    Using dp As New Pen(Color.FromArgb(40, 40, 55), 1)
                        dp.DashStyle = DashStyle.Dot
                        g.DrawLine(dp, chartLeft, y, chartRight, y)
                    End Using
                    Using fnt As New Font("Consolas", 7)
                        g.DrawString(bl.ToString(), fnt, Brushes.Gray, rect.X + 2, y - 6)
                    End Using
                End If
            Next
            ' 기준선 10/20
            For Each rv In {10, 20}
                If rv <= maxTick Then
                    Dim y = tickToY(rv)
                    Dim rc = If(rv = 20, Color.FromArgb(120, 255, 60, 60), Color.FromArgb(120, 255, 165, 0))
                    Using dp As New Pen(rc, 1) : dp.DashStyle = DashStyle.Dash : g.DrawLine(dp, chartLeft, y, chartRight, y) : End Using
                End If
            Next

            ' 바
            For i = 0 To count - 1
                Dim ti = tickVals(i) : If ti <= 0 Then Continue For
                Dim barH = CInt(histH * CDbl(ti) / maxTick)
                Dim cx = idxToX(i)
                Dim barColor = If(ti >= 20, Color.FromArgb(200, 255, 80, 80),
                               If(ti >= 10, Color.FromArgb(200, 255, 160, 60),
                               If(ti >= 5, Color.FromArgb(200, 255, 220, 100),
                                           Color.FromArgb(150, 100, 100, 100))))
                Using br As New SolidBrush(barColor)
                    g.FillRectangle(br, CInt(cx - candleW / 2), histBottom - barH, CInt(candleW), barH)
                End Using
            Next

            ' 틱MA
            If IsActive(IndicatorFlags.TickMA) Then
                DrawLineFunc(g, tMa5, idxToX, tickToY, Color.FromArgb(220, 255, 200, 50), 1.5F, count, DashStyle.Dash)
                DrawLineFunc(g, tMa20, idxToX, tickToY, Color.FromArgb(220, 100, 200, 255), 1.5F, count, DashStyle.Dash)
            End If

            Using fnt As New Font("맑은 고딕", 7.5F, FontStyle.Bold)
                g.DrawString("틱강도", fnt, New SolidBrush(Color.FromArgb(255, 160, 60)), chartLeft + 3, histTop + 2)
            End Using
        End If

        ' ═════════════════════════════════════
        '  5) RSI 패널
        ' ═════════════════════════════════════
        If showRsi Then
            ' 배경 그리드
            Using gridPen As New Pen(Color.FromArgb(35, 35, 50), 1)
                gridPen.DashStyle = DashStyle.Dot
                For Each lv In {30, 50, 70}
                    Dim y = RsiToY(lv, rsiTop, rsiH)
                    g.DrawLine(gridPen, chartLeft, y, chartRight, y)
                    Using fnt As New Font("Consolas", 7)
                        g.DrawString(lv.ToString(), fnt, Brushes.Gray, rect.X + 2, y - 6)
                    End Using
                Next
            End Using
            ' 기준선 30/70
            Using dp30 As New Pen(Color.FromArgb(100, 100, 200, 100), 1)
                dp30.DashStyle = DashStyle.Dash
                g.DrawLine(dp30, chartLeft, RsiToY(30, rsiTop, rsiH), chartRight, RsiToY(30, rsiTop, rsiH))
                g.DrawLine(dp30, chartLeft, RsiToY(70, rsiTop, rsiH), chartRight, RsiToY(70, rsiTop, rsiH))
            End Using
            Using dp50 As New Pen(Color.FromArgb(80, 150, 150, 150), 1)
                dp50.DashStyle = DashStyle.Dash
                g.DrawLine(dp50, chartLeft, RsiToY(50, rsiTop, rsiH), chartRight, RsiToY(50, rsiTop, rsiH))
            End Using

            Dim rsiToYFunc = Function(v As Double) As Integer
                                 Return RsiToY(v, rsiTop, rsiH)
                             End Function

            If IsActive(IndicatorFlags.RSI14) Then DrawLineFunc(g, rsi14, idxToX, rsiToYFunc, Color.FromArgb(255, 220, 100), 1.5F, count, DashStyle.Solid)
            If IsActive(IndicatorFlags.RSI5) Then DrawLineFunc(g, rsi5, idxToX, rsiToYFunc, Color.FromArgb(100, 255, 100), 1.2F, count, DashStyle.Solid)
            If IsActive(IndicatorFlags.RSI50) Then DrawLineFunc(g, rsi50, idxToX, rsiToYFunc, Color.FromArgb(255, 130, 130), 1.2F, count, DashStyle.Solid)

            Using fnt As New Font("맑은 고딕", 7.5F, FontStyle.Bold)
                g.DrawString("RSI", fnt, New SolidBrush(Color.FromArgb(255, 220, 100)), chartLeft + 3, rsiTop + 2)
            End Using
        End If

        ' ═════════════════════════════════════
        '  6) X축 시간 라벨
        ' ═════════════════════════════════════
        Dim bottomY = If(showRsi, rsiBottom, If(hasTick, histBottom, priceBottom))
        Dim lblInterval = Math.Max(1, count \ 12)
        Using fnt As New Font("맑은 고딕", 7)
            For i = 0 To count - 1 Step lblInterval
                Dim ts = ExtractTime(rows(i)("시간").ToString())
                Dim cx = idxToX(i)
                Dim sz = g.MeasureString(ts, fnt)
                g.DrawString(ts, fnt, Brushes.LightGray, cx - sz.Width / 2, bottomY + 3)
            Next
        End Using

        ' ═════════════════════════════════════
        '  7) 크로스헤어
        ' ═════════════════════════════════════
        If mousePoint <> Point.Empty AndAlso
           mousePoint.X >= chartLeft AndAlso mousePoint.X <= chartRight AndAlso
           mousePoint.Y >= priceTop AndAlso mousePoint.Y <= bottomY Then

            Dim nearIdx = CInt(Math.Round((mousePoint.X - chartLeft - gap / 2) / gap))
            nearIdx = Math.Max(0, Math.Min(count - 1, nearIdx))
            Dim nearX = idxToX(nearIdx)
            Dim nearRow = rows(nearIdx)

            ' 십자선
            Using crossPen As New Pen(Color.FromArgb(140, 255, 255, 255), 1)
                crossPen.DashStyle = DashStyle.Dash
                g.DrawLine(crossPen, chartLeft, mousePoint.Y, chartRight, mousePoint.Y)
                g.DrawLine(crossPen, nearX, priceTop, nearX, bottomY)
            End Using

            ' Y축 가격 라벨
            If mousePoint.Y >= priceTop AndAlso mousePoint.Y <= priceBottom Then
                Dim crossPrice = pMin + CLng(pRange * (1.0 - CDbl(mousePoint.Y - priceTop) / priceHeight))
                DrawYLabel(g, rect.X + 1, mousePoint.Y, crossPrice.ToString("N0"),
                           Color.FromArgb(220, 30, 30, 50), Brushes.Yellow)
            End If

            ' Y축 틱강도 라벨
            If hasTick AndAlso mousePoint.Y >= histTop AndAlso mousePoint.Y <= histBottom Then
                Dim crossTick = Math.Max(0, CInt(maxTick * (1.0 - CDbl(mousePoint.Y - histTop) / histH)))
                DrawYLabel(g, rect.X + 1, mousePoint.Y, $"틱:{crossTick}",
                           Color.FromArgb(220, 60, 30, 0), Brushes.Orange)
            End If

            ' Y축 RSI 라벨
            If showRsi AndAlso mousePoint.Y >= rsiTop AndAlso mousePoint.Y <= rsiBottom Then
                Dim crossRsi = Math.Max(0, Math.Min(100, CInt(100 * (1.0 - CDbl(mousePoint.Y - rsiTop) / rsiH))))
                DrawYLabel(g, rect.X + 1, mousePoint.Y, $"RSI:{crossRsi}",
                           Color.FromArgb(220, 50, 50, 0), Brushes.Yellow)
            End If

            ' X축 시간 라벨
            Dim nearTime = ExtractTime(nearRow("시간").ToString())
            Using fnt As New Font("맑은 고딕", 8, FontStyle.Bold)
                Dim sz = g.MeasureString(nearTime, fnt)
                Dim tx = nearX - sz.Width / 2
                Dim tRect As New RectangleF(CSng(tx) - 2, bottomY + 1, sz.Width + 4, sz.Height + 2)
                g.FillRectangle(New SolidBrush(Color.FromArgb(220, 30, 30, 50)), tRect)
                g.DrawString(nearTime, fnt, Brushes.Yellow, tRect.X + 2, tRect.Y + 1)
            End Using

            ' 하이라이트
            Using hlPen As New Pen(Color.FromArgb(80, 255, 255, 0), 1)
                g.DrawRectangle(hlPen, CInt(nearX - candleW / 2 - 2), priceTop,
                                CInt(candleW + 4), priceBottom - priceTop)
            End Using

            ' ── 정보 패널 ──
            Dim nO = CLng(nearRow("시가")), nH = CLng(nearRow("고가"))
            Dim nL = CLng(nearRow("저가")), nC = CLng(nearRow("종가"))
            Dim nPct As Decimal = 0
            If chartData.Columns.Contains("등락%") AndAlso Not IsDBNull(nearRow("등락%")) Then nPct = CDec(nearRow("등락%"))
            Dim nTi = If(hasTick AndAlso Not IsDBNull(nearRow("틱강도")), CInt(nearRow("틱강도")), 0)

            Dim line1 = $"{nearRow("시간")}  시:{nO:N0} 고:{nH:N0} 저:{nL:N0} 종:{nC:N0}  등락:{nPct:+0.00;-0.00}%"
            Dim line2 = ""
            If IsActive(IndicatorFlags.MA5) Then line2 &= $"MA5:{ma5(nearIdx):N0} "
            If IsActive(IndicatorFlags.MA20) Then line2 &= $"MA20:{ma20(nearIdx):N0} "
            If IsActive(IndicatorFlags.MA120) AndAlso ma120(nearIdx) > 0 Then line2 &= $"MA120:{ma120(nearIdx):N0} "
            If IsActive(IndicatorFlags.SuperTrend) AndAlso stDir(nearIdx) <> 0 Then
                Dim stVal = If(stDir(nearIdx) = 1, stUp(nearIdx), stDn(nearIdx))
                line2 &= $"ST:{stVal:N0}({If(stDir(nearIdx) = 1, "▲", "▼")}) "
            End If
            If IsActive(IndicatorFlags.JMA) AndAlso jma(nearIdx) > 0 Then line2 &= $"JMA:{jma(nearIdx):N0} "

            Dim line3 = ""
            If hasTick Then line3 &= $"틱강도:{nTi} tMA5:{tMa5(nearIdx):N1} tMA20:{tMa20(nearIdx):N1} "
            If showRsi Then
                If IsActive(IndicatorFlags.RSI14) AndAlso rsi14(nearIdx) >= 0 Then line3 &= $"RSI14:{rsi14(nearIdx):N1} "
                If IsActive(IndicatorFlags.RSI5) AndAlso rsi5(nearIdx) >= 0 Then line3 &= $"RSI5:{rsi5(nearIdx):N1} "
                If IsActive(IndicatorFlags.RSI50) AndAlso rsi50(nearIdx) >= 0 Then line3 &= $"RSI50:{rsi50(nearIdx):N1} "
            End If

            Dim lines = {line1, line2.TrimEnd(), line3.TrimEnd()}.Where(Function(s) s.Length > 0).ToArray()
            Using infoFont As New Font("맑은 고딕", 8.5F)
                Dim lineH = 16
                Dim panelW = 0
                For Each ln In lines
                    Dim w = CInt(g.MeasureString(ln, infoFont).Width) + 12
                    If w > panelW Then panelW = w
                Next
                Dim panelH = lines.Length * lineH + 8
                g.FillRectangle(New SolidBrush(Color.FromArgb(210, 10, 10, 25)),
                                chartLeft + 3, priceTop + 3, panelW, panelH)
                For j = 0 To lines.Length - 1
                    Dim br As Brush = Brushes.White
                    If j = 0 AndAlso nPct <> 0 Then br = New SolidBrush(If(nPct > 0, Color.FromArgb(255, 120, 120), Color.FromArgb(120, 120, 255)))
                    If j = lines.Length - 1 Then br = New SolidBrush(Color.FromArgb(255, 200, 100))
                    g.DrawString(lines(j), infoFont, br, chartLeft + 8, priceTop + 7 + j * lineH)
                Next
            End Using
        End If

        ' ═════════════════════════════════════
        '  8) 범례
        ' ═════════════════════════════════════
        DrawLegend(g, chartRight, priceTop)

        ' ═════════════════════════════════════
        '  9) 매매 신호
        ' ═════════════════════════════════════
        If CurrentPerformance IsNot Nothing AndAlso CurrentPerformance.Signals.Count > 0 Then
            For Each sig In CurrentPerformance.Signals
                If sig.BarIndex < 0 OrElse sig.BarIndex >= count Then Continue For
                Dim sx = idxToX(sig.BarIndex)
                Dim sy = priceToY(sig.Price)

                Select Case sig.Type
                    Case SignalType.Buy
                        ' 빨간 ▲
                        Dim pts = {New Point(sx, sy - 12), New Point(sx - 7, sy + 2), New Point(sx + 7, sy + 2)}
                        g.FillPolygon(New SolidBrush(Color.FromArgb(230, 255, 50, 50)), pts)
                        g.DrawPolygon(New Pen(Color.DarkRed, 1), pts)
                        Using fnt As New Font("Consolas", 7, FontStyle.Bold)
                            g.DrawString("B", fnt, Brushes.White, sx - 4, sy - 10)
                        End Using

                    Case SignalType.Sell, SignalType.StopLoss, SignalType.TakeProfit
                        ' 파란 ▼
                        Dim col = If(sig.Type = SignalType.StopLoss, Color.FromArgb(230, 255, 100, 0),
                                  If(sig.Type = SignalType.TakeProfit, Color.FromArgb(230, 0, 200, 100),
                                     Color.FromArgb(230, 50, 50, 255)))
                        Dim pts = {New Point(sx, sy + 12), New Point(sx - 7, sy - 2), New Point(sx + 7, sy - 2)}
                        g.FillPolygon(New SolidBrush(col), pts)
                        g.DrawPolygon(New Pen(Color.Black, 1), pts)
                        Dim lbl = If(sig.Type = SignalType.StopLoss, "SL", If(sig.Type = SignalType.TakeProfit, "TP", "S"))
                        Using fnt As New Font("Consolas", 7, FontStyle.Bold)
                            g.DrawString(lbl, fnt, Brushes.White, sx - 5, sy - 1)
                        End Using
                End Select
            Next

            ' 성과 요약 (우하단)
            Dim p = CurrentPerformance
            Dim perfText = $"거래:{p.TotalTrades} 승:{p.WinTrades} 패:{p.LossTrades} 승률:{p.WinRate:N1}% " &
                           $"총수익:{p.TotalReturnPct:N2}% 평균:{p.AvgReturnPct:N2}% MDD:{p.MaxDrawdownPct:N2}% PF:{p.ProfitFactor:N2}"
            Using fnt As New Font("맑은 고딕", 8.5F, FontStyle.Bold)
                Dim sz = g.MeasureString(perfText, fnt)
                Dim px = chartRight - sz.Width - 5
                Dim py = priceBottom - sz.Height - 5
                g.FillRectangle(New SolidBrush(Color.FromArgb(210, 10, 10, 25)), px - 3, py - 2, sz.Width + 6, sz.Height + 4)
                Dim tc = If(p.TotalReturnPct >= 0, Color.FromArgb(100, 255, 100), Color.FromArgb(255, 100, 100))
                g.DrawString(perfText, fnt, New SolidBrush(tc), px, py)
            End Using
        End If



    End Sub

    ' ══════════════════════════════════════════
    '  헬퍼: 범례
    ' ══════════════════════════════════════════
    Private Shared Sub DrawLegend(g As Graphics, chartRight As Integer, priceTop As Integer)
        Using fnt As New Font("맑은 고딕", 7)
            Dim items As New List(Of Tuple(Of String, Color, DashStyle))
            If IsActive(IndicatorFlags.MA5) Then items.Add(Tuple.Create("MA5", Color.FromArgb(255, 200, 50), DashStyle.Solid))
            If IsActive(IndicatorFlags.MA20) Then items.Add(Tuple.Create("MA20", Color.FromArgb(100, 200, 255), DashStyle.Solid))
            If IsActive(IndicatorFlags.MA120) Then items.Add(Tuple.Create("MA120", Color.FromArgb(200, 100, 255), DashStyle.Solid))
            If IsActive(IndicatorFlags.SuperTrend) Then items.Add(Tuple.Create("ST▲", Color.FromArgb(0, 200, 100), DashStyle.Solid))
            If IsActive(IndicatorFlags.SuperTrend) Then items.Add(Tuple.Create("ST▼", Color.FromArgb(255, 80, 80), DashStyle.Solid))
            '//If IsActive(IndicatorFlags.JMA) Then items.Add(Tuple.Create("JMA", Color.FromArgb(255, 100, 200), DashStyle.Solid))
            If IsActive(IndicatorFlags.JMA) Then
                items.Add(Tuple.Create("JMA▲", Color.FromArgb(0, 220, 120), DashStyle.Solid))
                items.Add(Tuple.Create("JMA▼", Color.FromArgb(255, 80, 180), DashStyle.Solid))
            End If

            If IsActive(IndicatorFlags.TickIntensity) Then items.Add(Tuple.Create("틱강도", Color.FromArgb(255, 160, 60), DashStyle.Solid))
            If IsActive(IndicatorFlags.RSI14) Then items.Add(Tuple.Create("RSI14", Color.FromArgb(255, 220, 100), DashStyle.Solid))
            If IsActive(IndicatorFlags.RSI5) Then items.Add(Tuple.Create("RSI5", Color.FromArgb(100, 255, 100), DashStyle.Solid))
            If IsActive(IndicatorFlags.RSI50) Then items.Add(Tuple.Create("RSI50", Color.FromArgb(255, 130, 130), DashStyle.Solid))

            If items.Count = 0 Then Return

            Dim colW = 75, rowH = 14, cols = 3
            Dim totalRows = CInt(Math.Ceiling(items.Count / CDbl(cols)))
            Dim lW = colW * cols + 10, lH = totalRows * rowH + 6
            Dim lx = chartRight - lW - 5, ly = priceTop + 4

            g.FillRectangle(New SolidBrush(Color.FromArgb(190, 10, 10, 20)), lx, ly, lW, lH)
            For idx = 0 To items.Count - 1
                Dim col = idx Mod cols, row = idx \ cols
                Dim ix = lx + 5 + col * colW, iy = ly + 3 + row * rowH
                Using pen As New Pen(items(idx).Item2, 2) : pen.DashStyle = items(idx).Item3
                    g.DrawLine(pen, ix, iy + 6, ix + 18, iy + 6)
                End Using
                g.DrawString(items(idx).Item1, fnt, Brushes.LightGray, ix + 20, iy)
            Next
        End Using



    End Sub





    ' ══════════════════════════════════════════
    '  헬퍼: 라인 그리기
    ' ══════════════════════════════════════════
    Private Shared Sub DrawLine(g As Graphics, values() As Double,
                                 idxToX As Func(Of Integer, Integer),
                                 valToY As Func(Of Double, Integer),
                                 col As Color, width As Single, count As Integer)
        DrawLineFunc(g, values, idxToX, valToY, col, width, count, DashStyle.Solid)
    End Sub

    Private Shared Sub DrawLineFunc(g As Graphics, values() As Double,
                                     idxToX As Func(Of Integer, Integer),
                                     valToY As Func(Of Double, Integer),
                                     col As Color, width As Single, count As Integer,
                                     style As DashStyle)
        If values Is Nothing Then Return
        Using pen As New Pen(col, width)
            pen.DashStyle = style
            pen.LineJoin = LineJoin.Round
            Dim prev As Point = Nothing
            Dim hasPrev = False
            For i = 0 To count - 1
                If i >= values.Length Then Exit For
                If Double.IsNaN(values(i)) OrElse values(i) <= 0 Then hasPrev = False : Continue For
                Dim pt As New Point(idxToX(i), valToY(values(i)))
                If hasPrev Then g.DrawLine(pen, prev, pt)
                prev = pt : hasPrev = True
            Next
        End Using
    End Sub

    ' ══════════════════════════════════════════
    '  헬퍼: Y축 라벨
    ' ══════════════════════════════════════════
    Private Shared Sub DrawYLabel(g As Graphics, x As Integer, y As Integer,
                                   txt As String, bgColor As Color, textBrush As Brush)
        Using fnt As New Font("Consolas", 8, FontStyle.Bold)
            Dim sz = g.MeasureString(txt, fnt)
            g.FillRectangle(New SolidBrush(bgColor), x, y - sz.Height / 2, sz.Width + 6, sz.Height + 2)
            g.DrawString(txt, fnt, textBrush, x + 3, y - sz.Height / 2 + 1)
        End Using
    End Sub

    Private Shared Function RsiToY(rsiVal As Double, rsiTop As Integer, rsiH As Integer) As Integer
        Return CInt(rsiTop + rsiH * (1.0 - rsiVal / 100.0))
    End Function

    Private Shared Function ExtractTime(fullTime As String) As String
        If String.IsNullOrEmpty(fullTime) Then Return ""
        If fullTime.Length >= 16 Then Return fullTime.Substring(11, 5)
        If fullTime.Length >= 5 AndAlso fullTime.Contains(":") Then Return fullTime.Substring(0, 5)
        Return fullTime
    End Function

    ' ══════════════════════════════════════════
    '  기술적 지표 계산
    ' ══════════════════════════════════════════

    ''' <summary>단순이동평균</summary>
    Private Shared Function CalcSMA(data() As Double, period As Integer) As Double()
        Dim n = data.Length
        Dim result(n - 1) As Double
        Dim sum As Double = 0
        For i = 0 To n - 1
            sum += data(i)
            If i >= period Then sum -= data(i - period)
            If i >= period - 1 Then
                result(i) = sum / period
            Else
                result(i) = Double.NaN
            End If
        Next
        Return result
    End Function

    Private Shared Function CalcSMAFromInt(data() As Integer, period As Integer) As Double()
        Dim dbl(data.Length - 1) As Double
        For i = 0 To data.Length - 1 : dbl(i) = data(i) : Next
        Return CalcSMA(dbl, period)
    End Function

    ''' <summary>RSI(period)</summary>
    Private Shared Function CalcRSI(closes() As Double, period As Integer) As Double()
        Dim n = closes.Length
        Dim result(n - 1) As Double
        For i = 0 To n - 1 : result(i) = Double.NaN : Next
        If n <= period Then Return result

        Dim gainSum As Double = 0, lossSum As Double = 0
        For i = 1 To period
            Dim chg = closes(i) - closes(i - 1)
            If chg > 0 Then gainSum += chg Else lossSum += Math.Abs(chg)
        Next
        Dim avgGain = gainSum / period
        Dim avgLoss = lossSum / period
        If avgLoss = 0 Then result(period) = 100 Else result(period) = 100 - (100 / (1 + avgGain / avgLoss))

        For i = period + 1 To n - 1
            Dim chg = closes(i) - closes(i - 1)
            Dim gain = If(chg > 0, chg, 0)
            Dim loss = If(chg < 0, Math.Abs(chg), 0)
            avgGain = (avgGain * (period - 1) + gain) / period
            avgLoss = (avgLoss * (period - 1) + loss) / period
            If avgLoss = 0 Then result(i) = 100 Else result(i) = 100 - (100 / (1 + avgGain / avgLoss))
        Next
        Return result
    End Function

    ''' <summary>SuperTrend(atrPeriod, multiplier)</summary>
    Private Shared Sub CalcSuperTrend(highs() As Double, lows() As Double, closes() As Double,
                                       atrPeriod As Integer, multiplier As Double,
                                       ByRef upperBand() As Double, ByRef lowerBand() As Double,
                                       ByRef direction() As Integer)
        Dim n = closes.Length
        upperBand = New Double(n - 1) {}
        lowerBand = New Double(n - 1) {}
        direction = New Integer(n - 1) {}

        ' ATR 계산 (Wilder smoothing)
        Dim tr(n - 1) As Double
        Dim atr(n - 1) As Double
        tr(0) = highs(0) - lows(0)
        For i = 1 To n - 1
            Dim hl = highs(i) - lows(i)
            Dim hc = Math.Abs(highs(i) - closes(i - 1))
            Dim lc = Math.Abs(lows(i) - closes(i - 1))
            tr(i) = Math.Max(hl, Math.Max(hc, lc))
        Next
        Dim atrSum As Double = 0
        For i = 0 To Math.Min(atrPeriod - 1, n - 1) : atrSum += tr(i) : Next
        If atrPeriod <= n Then atr(atrPeriod - 1) = atrSum / atrPeriod
        For i = atrPeriod To n - 1
            atr(i) = (atr(i - 1) * (atrPeriod - 1) + tr(i)) / atrPeriod
        Next

        ' SuperTrend
        Dim basicUp(n - 1) As Double, basicDn(n - 1) As Double
        Dim finalUp(n - 1) As Double, finalDn(n - 1) As Double

        For i = 0 To n - 1
            Dim mid = (highs(i) + lows(i)) / 2
            basicUp(i) = mid - multiplier * atr(i)
            basicDn(i) = mid + multiplier * atr(i)
        Next

        finalUp(0) = basicUp(0) : finalDn(0) = basicDn(0) : direction(0) = 1
        upperBand(0) = finalUp(0) : lowerBand(0) = finalDn(0)

        For i = 1 To n - 1
            ' upper band
            finalUp(i) = basicUp(i)
            If basicUp(i) > finalUp(i - 1) OrElse closes(i - 1) > finalUp(i - 1) Then
                finalUp(i) = Math.Max(basicUp(i), finalUp(i - 1))
            End If

            ' lower band
            finalDn(i) = basicDn(i)
            If basicDn(i) < finalDn(i - 1) OrElse closes(i - 1) < finalDn(i - 1) Then
                finalDn(i) = Math.Min(basicDn(i), finalDn(i - 1))
            End If

            ' direction
            If direction(i - 1) = 1 Then
                If closes(i) < finalUp(i) Then direction(i) = -1 Else direction(i) = 1
            Else
                If closes(i) > finalDn(i) Then direction(i) = 1 Else direction(i) = -1
            End If

            upperBand(i) = finalUp(i)
            lowerBand(i) = finalDn(i)
        Next
    End Sub

    ''' <summary>JMA - Jurik Moving Average 근사 구현 (length, phase, power)</summary>
    Private Shared Function CalcJMA(closes() As Double, length As Integer,
                                     phase As Integer, power As Integer) As Double()
        Dim n = closes.Length
        Dim result(n - 1) As Double

        ' phase → beta 변환
        Dim phaseRatio As Double
        If phase < -100 Then
            phaseRatio = 0.5
        ElseIf phase > 100 Then
            phaseRatio = 2.5
        Else
            phaseRatio = CDbl(phase) / 100.0 + 1.5
        End If

        ' length → alpha 변환
        Dim beta = 0.45 * (length - 1) / (0.45 * (length - 1) + 2.0)
        Dim alpha = Math.Pow(beta, CDbl(power))

        ' 초기화
        Dim e0 As Double = closes(0)
        Dim e1 As Double = 0
        Dim e2 As Double = 0
        Dim jmaVal As Double = closes(0)
        result(0) = closes(0)

        For i = 1 To n - 1
            e0 = (1.0 - alpha) * closes(i) + alpha * e0
            e1 = (closes(i) - e0) * (1.0 - beta) + beta * e1
            e2 = (e0 + phaseRatio * e1 - jmaVal) * Math.Pow(1.0 - alpha, 2) + Math.Pow(alpha, 2) * e2
            jmaVal = jmaVal + e2
            result(i) = jmaVal
        Next

        Return result
    End Function

End Class
