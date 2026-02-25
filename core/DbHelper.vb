' ===== Core/DbHelper.vb =====
' DB 연결, config.ini 읽기 — 수정 불필요

Imports MySql.Data.MySqlClient
Imports System.IO

Public Class DbHelper

    ' ── config.ini 경로 탐색 ──
    Private Shared Function FindIniPath() As String
        Dim candidates = {
            Path.Combine(Application.StartupPath, "config.ini"),
            Path.Combine(Application.StartupPath, "..", "..", "config.ini"),
            Path.Combine(Application.StartupPath, "..", "..", "..", "config.ini")
        }
        For Each p In candidates
            If File.Exists(p) Then Return Path.GetFullPath(p)
        Next
        Throw New FileNotFoundException(
            "config.ini를 찾을 수 없습니다." & vbCrLf &
            "탐색 경로: " & Application.StartupPath)
    End Function

    ' ── ini 값 읽기 ──
    Public Shared Function GetIniValue(section As String, key As String,
                                       Optional defaultValue As String = "") As String
        Dim iniPath = FindIniPath()
        Dim currentSection = ""
        For Each line In File.ReadAllLines(iniPath, System.Text.Encoding.UTF8)
            Dim trimmed = line.Trim()
            If String.IsNullOrEmpty(trimmed) OrElse trimmed.StartsWith(";") OrElse
               trimmed.StartsWith("#") Then Continue For
            If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                currentSection = trimmed.Substring(1, trimmed.Length - 2).Trim()
                Continue For
            End If
            If currentSection.Equals(section, StringComparison.OrdinalIgnoreCase) Then
                Dim eqIdx = trimmed.IndexOf("="c)
                If eqIdx > 0 Then
                    Dim k = trimmed.Substring(0, eqIdx).Trim()
                    Dim v = trimmed.Substring(eqIdx + 1).Trim()
                    If k.Equals(key, StringComparison.OrdinalIgnoreCase) Then Return v
                End If
            End If
        Next
        Return defaultValue
    End Function

    ' ── DB 연결 문자열 ──
    Public Shared Function GetConnectionString() As String
        Return $"Server={GetIniValue("Database", "Server", "localhost")};" &
               $"Port={GetIniValue("Database", "Port", "3306")};" &
               $"Database={GetIniValue("Database", "Database", "stock_information")};" &
               $"Uid={GetIniValue("Database", "Uid", "root")};" &
               $"Pwd={GetIniValue("Database", "Pwd", "")};" &
               $"CharSet={GetIniValue("Database", "CharSet", "utf8mb4")};"
    End Function

    ' ── 커넥션 생성 ──
    Public Shared Function CreateConnection() As MySqlConnection
        Return New MySqlConnection(GetConnectionString())
    End Function

    ' ── SELECT 실행 → DataTable ──
    Public Shared Function ExecuteQuery(sql As String) As DataTable
        Using conn = CreateConnection()
            conn.Open()
            Using cmd As New MySqlCommand(sql, conn)
                cmd.CommandTimeout = 60
                Using adapter As New MySqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' ── INSERT/UPDATE 실행 ──
    Public Shared Function ExecuteNonQuery(sql As String,
                                           ParamArray params() As MySqlParameter) As Integer
        Using conn = CreateConnection()
            conn.Open()
            Using cmd As New MySqlCommand(sql, conn)
                If params IsNot Nothing Then cmd.Parameters.AddRange(params)
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    ' ── 스칼라 조회 ──
    Public Shared Function ExecuteScalar(sql As String,
                                         ParamArray params() As MySqlParameter) As Object
        Using conn = CreateConnection()
            conn.Open()
            Using cmd As New MySqlCommand(sql, conn)
                If params IsNot Nothing Then cmd.Parameters.AddRange(params)
                Return cmd.ExecuteScalar()
            End Using
        End Using
    End Function

    ' ── 종목코드 → 이름 ──
    Public Shared Function GetStockName(code As String) As String
        Dim p As New MySqlParameter("@code", code)
        Dim result = ExecuteScalar("SELECT name FROM stock_base_info WHERE code=@code LIMIT 1", p)
        If result IsNot Nothing AndAlso Not IsDBNull(result) Then Return result.ToString()
        Return ""
    End Function

    ' ── 종목명 → 전체정보 맵 ──
    Public Shared Function GetStockNameMap() As Dictionary(Of String, Dictionary(Of String, Object))
        Dim nameMap As New Dictionary(Of String, Dictionary(Of String, Object))
        Using conn = CreateConnection()
            conn.Open()
            Using cmd As New MySqlCommand(
                "SELECT code, name, market, market_cap, sector FROM stock_base_info WHERE name IS NOT NULL AND name<>''", conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim n = reader("name").ToString().Trim()
                        If Not nameMap.ContainsKey(n) Then
                            Dim info As New Dictionary(Of String, Object)
                            info("code") = reader("code").ToString()
                            info("market") = If(IsDBNull(reader("market")), "", reader("market").ToString())
                            info("market_cap") = If(IsDBNull(reader("market_cap")), DBNull.Value, reader("market_cap"))
                            info("sector") = If(IsDBNull(reader("sector")), "", reader("sector").ToString())
                            nameMap(n) = info
                        End If
                    End While
                End Using
            End Using
        End Using
        Return nameMap
    End Function

    ' ── 숫자 파싱 헬퍼 ──
    Public Shared Function ToDecimalOrNull(value As Object) As Object
        If value Is Nothing OrElse value Is DBNull.Value Then Return DBNull.Value
        Dim s = value.ToString().Trim()
        If String.IsNullOrEmpty(s) Then Return DBNull.Value
        Dim result As Decimal
        If Decimal.TryParse(s, result) Then Return result
        Return DBNull.Value
    End Function

    Public Shared Function ToLongOrNull(value As Object) As Object
        If value Is Nothing OrElse value Is DBNull.Value Then Return DBNull.Value
        Dim s = value.ToString().Replace(",", "").Trim()
        If String.IsNullOrEmpty(s) Then Return DBNull.Value
        Dim result As Long
        If Long.TryParse(s, result) Then Return result
        Return DBNull.Value
    End Function

End Class

