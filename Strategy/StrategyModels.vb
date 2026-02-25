' ===== Strategy/StrategyModels.vb =====
' 전략 정의 모델 + 신호 결과 + HoldConditions + Grade/Version

Imports System.Drawing

''' <summary>매매 조건 하나</summary>
Public Class StrategyCondition
    Public Property Indicator As String = ""
    Public Property [Operator] As String = ""
    Public Property Target As String = ""
    Public Property Value As Double = 0

    Public Overrides Function ToString() As String
        Return $"{Indicator} {[Operator]} {If(String.IsNullOrEmpty(Target), Value.ToString("N2"), Target)}"
    End Function
End Class

''' <summary>전략 정의</summary>
Public Class TradingStrategy
    Public Property Id As Integer = 0
    Public Property Name As String = ""
    Public Property Description As String = ""
    Public Property BuyConditions As New List(Of StrategyCondition)
    Public Property SellConditions As New List(Of StrategyCondition)
    Public Property HoldConditions As New List(Of StrategyCondition)   ' 매도 유예 조건
    Public Property StopLossPct As Double = -3.0
    Public Property TakeProfitPct As Double = 10.0
    Public Property MaxHoldBars As Integer = 30
    Public Property CreatedDate As DateTime = DateTime.Now
    Public Property IsActive As Boolean = True

    ' ── 버전 관리 ──
    Public Property Grade As String = ""           ' BASE / DRAFT / CANDIDATE / PROMOTED
    Public Property Version As String = "v1"
    Public Property ParentName As String = ""
    Public Property IsLocked As Boolean = False
    Public Property ChangeLog As String = ""

    ' ── 성과 기준 ──
    Public Property BaselineWinRate As Double = 0
    Public Property BaselineAvgReturn As Double = 0
    Public Property BaselineSampleCount As Integer = 0
    Public Property TestWinRate As Double = 0
    Public Property TestAvgReturn As Double = 0
    Public Property TestSampleCount As Integer = 0
    Public Property TestPeriod As String = ""

    ' === TradingStrategy 클래스에 추가할 속성들 ===

    ''' 홀드(매도유예) 조건
    'Public Property HoldConditions As New List(Of StrategyCondition)

    '''' 전략 등급 (BASE, DRAFT, PROMOTED)
    'Public Property Grade As String = "DRAFT"

    '''' 전략 버전
    'Public Property Version As String = "v1"

    '''' 잠금 여부 (BASE 전략은 True)
    'Public Property IsLocked As Boolean = False

    ''' VI 직전 매도 활성화
    Public Property UseViPreSell As Boolean = False

    ''' VI 직전 매도 임계값 (시가대비 상승률 %)
    Public Property ViPreSellPct As Double = 8.0

    ''' VI 후 재진입 허용
    Public Property AllowViReentry As Boolean = False

    ''' VI 후 재진입 대기 봉수
    Public Property ViReentryWaitBars As Integer = 5

    ''' 2차 VI 매도 임계값 (%)
    Public Property ViSecondSellPct As Double = 18.0

    ''' 일일 수익 달성 후 매매 금지
    Public Property DailyProfitLock As Boolean = False


    Public Overrides Function ToString() As String
        Dim g = If(String.IsNullOrEmpty(Grade), "", $"[{Grade}] ")
        Return $"{g}{Name} (매수:{BuyConditions.Count}조건, 매도:{SellConditions.Count}조건)"
    End Function
End Class

''' <summary>매매 신호</summary>
Public Enum SignalType
    Buy = 1
    Sell = -1
    StopLoss = -2
    TakeProfit = -3
    VIExit = -4          ' VI 직전 매도
    DailyDone = -5       ' 일일 매매 완료
End Enum

''' <summary>차트에 표시할 매매 신호</summary>
Public Class TradeSignal
    Public Property BarIndex As Integer
    Public Property Type As SignalType
    Public Property Price As Double
    Public Property Time As String = ""
    Public Property Reason As String = ""
End Class

''' <summary>전략 성과 요약</summary>
Public Class StrategyPerformance
    Public Property StrategyName As String = ""
    Public Property TotalTrades As Integer = 0
    Public Property WinTrades As Integer = 0
    Public Property LossTrades As Integer = 0
    Public Property WinRate As Double = 0
    Public Property TotalReturnPct As Double = 0
    Public Property AvgReturnPct As Double = 0
    Public Property MaxDrawdownPct As Double = 0
    Public Property AvgHoldBars As Double = 0
    Public Property ProfitFactor As Double = 0
    Public Property Signals As New List(Of TradeSignal)
    Public Property TradeDetails As New List(Of TradeDetail)

    ' ── 10%+ 적중 통계 ──
    Public Property TotalSignals As Integer = 0        ' 전체 포착 수
    Public Property Rise10Count As Integer = 0         ' 10%+ 상승 수
    Public Property Rise10Rate As Double = 0           ' 10%+ 적중률
End Class

''' <summary>개별 거래 상세</summary>
Public Class TradeDetail
    Public Property EntryIndex As Integer
    Public Property ExitIndex As Integer
    Public Property EntryPrice As Double
    Public Property ExitPrice As Double
    Public Property ReturnPct As Double
    Public Property EntryTime As String = ""
    Public Property ExitTime As String = ""
    Public Property ExitReason As String = ""
    Public Property HoldBars As Integer
    Public Property MaxRisePct As Double = 0           ' 보유 중 최대 상승률
End Class