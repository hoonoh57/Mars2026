' ===== Core/ApiClient.vb =====
' server32 REST API 통신 — 수정 불필요

Public Class ApiClient

    Private Shared Function GetBaseUrl() As String
        Return DbHelper.GetIniValue("Server32", "BaseUrl", "http://localhost:8082")
    End Function

    ' ── JSON 다운로드 ──
    Public Shared Function DownloadJson(url As String) As String
        Using client As New Net.WebClient()
            client.Encoding = System.Text.Encoding.UTF8
            Return client.DownloadString(url)
        End Using
    End Function

    ' ── 분봉 URL ──
    Public Shared Function MinuteCandleUrl(code As String, tick As String,
                                            count As Integer, stopTime As String) As String
        Return $"{GetBaseUrl()}/api/market/candles/minute?code={code}&tick={tick}&count={count}&stopTime={stopTime}"
    End Function

    ' ── 틱봉 URL ──
    Public Shared Function TickCandleUrl(code As String, tick As String,
                                          count As Integer, stopTime As String) As String
        Return $"{GetBaseUrl()}/api/market/candles/tick?code={code}&tick={tick}&count={count}&stopTime={stopTime}"
    End Function

End Class
