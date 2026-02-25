' ===== Strategy/StrategyStore.vb =====
' 전략 저장/로드/삭제 (JSON 파일 기반)

Imports System.IO

Public Class StrategyStore

    Private Shared ReadOnly Property FolderPath As String
        Get
            Dim folder = Path.Combine(Application.StartupPath, "strategies")
            If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
            Return folder
        End Get
    End Property

    Public Shared Function LoadAll() As List(Of TradingStrategy)
        Dim list As New List(Of TradingStrategy)
        For Each fp In Directory.GetFiles(FolderPath, "*.json").OrderBy(Function(f) f)
            Try
                Dim json = File.ReadAllText(fp, Text.Encoding.UTF8)
                Dim s = DeserializeStrategy(json)
                If s IsNot Nothing Then list.Add(s)
            Catch
            End Try
        Next
        Return list
    End Function

    Public Shared Sub Save(strategy As TradingStrategy)
        Dim safeName = strategy.Name
        For Each c In Path.GetInvalidFileNameChars() : safeName = safeName.Replace(c, "_"c) : Next
        Dim fp = Path.Combine(FolderPath, safeName & ".json")
        Dim json = SerializeStrategy(strategy)
        File.WriteAllText(fp, json, Text.Encoding.UTF8)
    End Sub

    Public Shared Sub Delete(strategy As TradingStrategy)
        Dim safeName = strategy.Name
        For Each c In Path.GetInvalidFileNameChars() : safeName = safeName.Replace(c, "_"c) : Next
        Dim fp = Path.Combine(FolderPath, safeName & ".json")
        If File.Exists(fp) Then File.Delete(fp)
    End Sub

    ' ── 간단한 JSON 직렬화 (Newtonsoft 없이) ──
    Private Shared Function SerializeStrategy(s As TradingStrategy) As String
        Dim sb As New Text.StringBuilder()
        sb.AppendLine("{")
        sb.AppendLine($"  ""Name"": ""{Esc(s.Name)}"",")
        sb.AppendLine($"  ""Description"": ""{Esc(s.Description)}"",")
        sb.AppendLine($"  ""StopLossPct"": {s.StopLossPct},")
        sb.AppendLine($"  ""TakeProfitPct"": {s.TakeProfitPct},")
        sb.AppendLine($"  ""MaxHoldBars"": {s.MaxHoldBars},")
        sb.AppendLine($"  ""IsActive"": {s.IsActive.ToString().ToLower()},")
        sb.AppendLine($"  ""BuyConditions"": [")
        For i = 0 To s.BuyConditions.Count - 1
            Dim c = s.BuyConditions(i)
            sb.Append($"    {{""Indicator"":""{Esc(c.Indicator)}"",""Operator"":""{Esc(c.Operator)}"",""Target"":""{Esc(c.Target)}"",""Value"":{c.Value}}}")
            If i < s.BuyConditions.Count - 1 Then sb.Append(",")
            sb.AppendLine()
        Next
        sb.AppendLine("  ],")
        sb.AppendLine($"  ""SellConditions"": [")
        For i = 0 To s.SellConditions.Count - 1
            Dim c = s.SellConditions(i)
            sb.Append($"    {{""Indicator"":""{Esc(c.Indicator)}"",""Operator"":""{Esc(c.Operator)}"",""Target"":""{Esc(c.Target)}"",""Value"":{c.Value}}}")
            If i < s.SellConditions.Count - 1 Then sb.Append(",")
            sb.AppendLine()
        Next
        sb.AppendLine("  ]")
        sb.AppendLine("}")
        Return sb.ToString()
    End Function

    Private Shared Function Esc(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""")
    End Function

    Private Shared Function DeserializeStrategy(json As String) As TradingStrategy
        Dim s As New TradingStrategy()
        s.Name = ExtractStr(json, "Name")
        s.Description = ExtractStr(json, "Description")
        s.StopLossPct = ExtractDbl(json, "StopLossPct", -3)
        s.TakeProfitPct = ExtractDbl(json, "TakeProfitPct", 10)
        s.MaxHoldBars = CInt(ExtractDbl(json, "MaxHoldBars", 30))
        Dim isAct = ExtractStr(json, "IsActive")
        s.IsActive = If(isAct = "false", False, True)
        s.BuyConditions = ExtractConditions(json, "BuyConditions")
        s.SellConditions = ExtractConditions(json, "SellConditions")
        Return s
    End Function

    Private Shared Function ExtractStr(json As String, key As String) As String
        Dim pattern = $"""{key}"":"
        Dim idx = json.IndexOf(pattern)
        If idx < 0 Then pattern = $"""{key}"" :" : idx = json.IndexOf(pattern)
        If idx < 0 Then Return ""
        Dim start = json.IndexOf(""""c, idx + pattern.Length)
        If start < 0 Then Return ""
        Dim endIdx = json.IndexOf(""""c, start + 1)
        If endIdx < 0 Then Return ""
        Return json.Substring(start + 1, endIdx - start - 1).Replace("\""", """").Replace("\\", "\")
    End Function

    Private Shared Function ExtractDbl(json As String, key As String, defaultVal As Double) As Double
        Dim pattern = $"""{key}"":"
        Dim idx = json.IndexOf(pattern)
        If idx < 0 Then pattern = $"""{key}"" :" : idx = json.IndexOf(pattern)
        If idx < 0 Then Return defaultVal
        Dim start = idx + pattern.Length
        Dim sb As New Text.StringBuilder()
        For i = start To Math.Min(start + 20, json.Length - 1)
            Dim c = json(i)
            If Char.IsDigit(c) OrElse c = "."c OrElse c = "-"c Then
                sb.Append(c)
            ElseIf sb.Length > 0 Then
                Exit For
            End If
        Next
        Dim result As Double
        If Double.TryParse(sb.ToString(), result) Then Return result
        Return defaultVal
    End Function

    Private Shared Function ExtractConditions(json As String, key As String) As List(Of StrategyCondition)
        Dim list As New List(Of StrategyCondition)
        Dim arrStart = json.IndexOf($"""{key}""")
        If arrStart < 0 Then Return list
        Dim bracket = json.IndexOf("["c, arrStart)
        If bracket < 0 Then Return list
        Dim bracketEnd = json.IndexOf("]"c, bracket)
        If bracketEnd < 0 Then Return list
        Dim arrJson = json.Substring(bracket + 1, bracketEnd - bracket - 1)
        Dim objStart = 0
        While True
            Dim os = arrJson.IndexOf("{"c, objStart)
            If os < 0 Then Exit While
            Dim oe = arrJson.IndexOf("}"c, os)
            If oe < 0 Then Exit While
            Dim obj = arrJson.Substring(os, oe - os + 1)
            Dim c As New StrategyCondition()
            c.Indicator = ExtractStr(obj, "Indicator")
            c.Operator = ExtractStr(obj, "Operator")
            c.Target = ExtractStr(obj, "Target")
            c.Value = ExtractDbl(obj, "Value", 0)
            list.Add(c)
            objStart = oe + 1
        End While
        Return list
    End Function

End Class
