' ===== Strategy/StrategyModels.vb =====
' 전략 정의 모델 + 신호 결과

Imports System.Drawing

''' <summary>매매 조건 하나</summary>
Public Class StrategyCondition
    Public Property Indicator As String = ""      ' MA5, MA20, RSI14, SuperTrend, JMA, TickIntensity, Price 등
    Public Property [Operator] As String = ""       ' >, <, >=, <=, CrossUp, CrossDown, =
    Public Property Target As String = ""         ' 숫자값 또는 다른 지표명
    Public Property Value As Double = 0           ' Target이 숫자일 때

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
    Public Property StopLossPct As Double = -3.0       ' 손절 %
    Public Property TakeProfitPct As Double = 10.0     ' 익절 %
    Public Property MaxHoldBars As Integer = 30        ' 최대 보유 봉수
    Public Property CreatedDate As DateTime = DateTime.Now
    Public Property IsActive As Boolean = True

    Public Overrides Function ToString() As String
        Return $"{Name} (매수:{BuyConditions.Count}조건, 매도:{SellConditions.Count}조건)"
    End Function
End Class

''' <summary>매매 신호</summary>
Public Enum SignalType
    Buy = 1
    Sell = -1
    StopLoss = -2
    TakeProfit = -3
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
End Class
