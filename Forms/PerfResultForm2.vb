' ===== Forms/PerfResultForm2.vb =====
' UI 레이아웃 + 이벤트 연결
' Core/*, Strategy/* 에 비즈니스 로직 위임

Imports MySql.Data.MySqlClient
Imports System.IO

Public Class PerfResultForm2
    Inherits Form

    ' ── 공통 ──
    Private tabControl As New TabControl()

    ' ── 탭1: 성과검증 입력 ──
    Private tabInput As New TabPage("성과검증 입력")
    Private WithEvents btnParseAndFill As New Button()
    Private WithEvents btnSave As New Button()
    Private dtpDate As New DateTimePicker()
    Private dtpTime As New DateTimePicker()
    Private txtClipboard As New TextBox()
    Private lblStatus As New Label()
    Private dgvResult As New DataGridView()

    ' ── 탭2: SQL 실행기 ──
    Private tabSql As New TabPage("SQL 분석")
    Private txtSql As New TextBox()
    Private WithEvents btnRunSql As New Button()
    Private WithEvents btnSaveSql As New Button()
    Private WithEvents btnDeleteSql As New Button()
    Private cboSavedQueries As New ComboBox()
    Private WithEvents btnLoadSql As New Button()
    Private lblSqlStatus As New Label()
    Private dgvSqlResult As New DataGridView()

    ' ── 탭3: 캔들 분석 (개편) ──
    Private tabCandle As New TabPage("캔들 분석")
    Private dtpCandleDate As New DateTimePicker()       ' 일자 선택
    Private WithEvents btnLoadPerfData As New Button()  ' 성과 데이터 불러오기
    Private dgvPerfStocks As New DataGridView()         ' 성과검증 종목 그리드
    Private lblCandleCode As New Label()
    Private txtCandleCode As New TextBox()
    Private lblCandleName As New Label()
    Private cboMinuteType As New ComboBox()
    Private dtpMinuteStop As New DateTimePicker()
    Private nudMinuteCount As New NumericUpDown()
    Private btnDownload As New Button()
    Private cboTickType As New ComboBox()
    Private dgvMinute As New DataGridView()
    Private dgvTick As New DataGridView()
    Private lblCandleStatus As New Label()
    Private pnlChart As Panel
    Private chartData As DataTable = Nothing
    Private crosshairMousePt As Point = Point.Empty

    ' ── 쿼리 저장 폴더 ──
    Private ReadOnly Property QueryFolderPath As String
        Get
            Dim folder = Path.Combine(Application.StartupPath, "saved_queries")
            If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
            Return folder
        End Get
    End Property

    ' ══════════════════════════════════════════
    '  폼 로드
    ' ══════════════════════════════════════════
    Private Sub PerfResultForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "MARS2026 성과검증 / SQL 분석 / 캔들"
        Me.Size = New Size(1400, 900)
        Me.StartPosition = FormStartPosition.CenterScreen

        tabControl.Dock = DockStyle.Fill
        Me.Controls.Add(tabControl)

        SetupTabInput()
        SetupTabSql()
        SetupTabCandle()

        tabControl.TabPages.Add(tabInput)
        tabControl.TabPages.Add(tabSql)
        tabControl.TabPages.Add(tabCandle)
    End Sub

    ' ══════════════════════════════════════════
    '  탭1: 성과검증 입력 UI (기존 유지)
    ' ══════════════════════════════════════════
    Private Sub SetupTabInput()
        Dim pnlTop As New FlowLayoutPanel() With {.Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(5)}
        Dim lblDate As New Label() With {.Text = "검색일자:", .AutoSize = True, .Margin = New Padding(0, 6, 0, 0)}
        dtpDate.Value = DateTime.Today : dtpDate.Width = 120
        Dim lblTime As New Label() With {.Text = "검색시각:", .AutoSize = True, .Margin = New Padding(10, 6, 0, 0)}
        dtpTime.Value = DateTime.Today.AddHours(9)
        dtpTime.Format = DateTimePickerFormat.Custom : dtpTime.CustomFormat = "HH:mm"
        dtpTime.ShowUpDown = True : dtpTime.Width = 80
        pnlTop.Controls.AddRange({lblDate, dtpDate, lblTime, dtpTime})

        txtClipboard.Multiline = True : txtClipboard.ScrollBars = ScrollBars.Both
        txtClipboard.Dock = DockStyle.Top : txtClipboard.Height = 160
        txtClipboard.Font = New Font("Consolas", 9)

        Dim pnlBtn As New FlowLayoutPanel() With {.Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(5)}
        btnParseAndFill.Text = "복사 및 파싱" : btnParseAndFill.Size = New Size(120, 30)
        btnSave.Text = "테이블에 저장" : btnSave.Size = New Size(120, 30)
        lblStatus.AutoSize = True : lblStatus.Margin = New Padding(20, 6, 0, 0) : lblStatus.Text = "대기 중..."
        pnlBtn.Controls.AddRange({btnParseAndFill, btnSave, lblStatus})

        dgvResult.Dock = DockStyle.Fill : dgvResult.AllowUserToAddRows = False
        dgvResult.Font = New Font("맑은 고딕", 9)
        dgvResult.AutoGenerateColumns = False : dgvResult.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        SetupInputGrid()

        tabInput.Controls.Add(dgvResult)
        tabInput.Controls.Add(pnlBtn)
        tabInput.Controls.Add(txtClipboard)
        tabInput.Controls.Add(pnlTop)
    End Sub

    Private Sub SetupInputGrid()
        dgvResult.Columns.Clear()
        For Each def In {"code:종목코드:80", "name:종목명:130", "market:시장:60",
                         "ret_1m:1분%:65", "ret_3m:3분%:65", "ret_7h:7시간%:70",
                         "max_ret:최고%:65", "search_volume:거래량:90", "extra:기타:55",
                         "market_cap:시총(억):75", "sector:업종:120", "is_winner:10%+:45"}
            Dim p = def.Split(":"c)
            dgvResult.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = p(0), .HeaderText = p(1), .Width = CInt(p(2))})
        Next
    End Sub

    Private Sub btnParseAndFill_Click(sender As Object, e As EventArgs) Handles btnParseAndFill.Click
        Dim rawText = txtClipboard.Text
        If String.IsNullOrWhiteSpace(rawText) Then
            If Clipboard.ContainsText() Then
                rawText = Clipboard.GetText()
                txtClipboard.Text = rawText
            Else
                MessageBox.Show("텍스트가 없습니다.", "알림")
                Return
            End If
        End If

        Dim rows = ParseClipboard(rawText)
        If rows.Count = 0 Then MessageBox.Show("파싱 결과가 없습니다.", "알림") : Return

        Dim nameMap = DbHelper.GetStockNameMap()
        For Each row In rows
            Dim name = row("name").ToString().Trim()
            If nameMap.ContainsKey(name) Then
                Dim info = nameMap(name)
                row("code") = info("code") : row("market") = info("market")
                row("market_cap") = info("market_cap") : row("sector") = info("sector")
            Else
                For Each kv In nameMap
                    If kv.Key.Contains(name) OrElse name.Contains(kv.Key) Then
                        row("code") = kv.Value("code") : row("market") = kv.Value("market")
                        row("market_cap") = kv.Value("market_cap") : row("sector") = kv.Value("sector")
                        Exit For
                    End If
                Next
            End If
        Next

        dgvResult.Rows.Clear()
        Dim winnerCount = 0
        For Each row In rows
            Dim idx = dgvResult.Rows.Add()
            Dim r = dgvResult.Rows(idx)
            For Each key In {"code", "name", "market", "ret_1m", "ret_3m", "ret_7h",
                             "max_ret", "search_volume", "extra", "market_cap", "sector"}
                r.Cells(key).Value = row(key)
            Next
            Dim mr As Decimal = 0
            If row("max_ret") IsNot Nothing AndAlso Not IsDBNull(row("max_ret")) AndAlso
               Decimal.TryParse(row("max_ret").ToString(), mr) AndAlso mr >= 10 Then
                r.Cells("is_winner").Value = "★" : r.DefaultCellStyle.BackColor = Color.LightYellow
                winnerCount += 1
            End If
            If row("code")?.ToString() = "??????" Then r.DefaultCellStyle.BackColor = Color.LightPink
        Next
        lblStatus.Text = $"파싱 완료: {rows.Count}종목 (10%+: {winnerCount}건)"
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim targetDate = dtpDate.Value.Date
        Dim searchTime = dtpTime.Value.TimeOfDay
        Dim savedCount = 0, skippedCount = 0

        Using conn = DbHelper.CreateConnection()
            conn.Open()
            For Each row As DataGridViewRow In dgvResult.Rows
                If row.IsNewRow Then Continue For
                Dim code = row.Cells("code").Value?.ToString()
                If String.IsNullOrEmpty(code) OrElse code = "??????" Then skippedCount += 1 : Continue For
                Dim sql = "INSERT INTO perf_result(target_date,search_time,code,name,market," &
                    "ret_1m,ret_3m,ret_7h,max_ret,search_volume,extra,market_cap,sector) " &
                    "VALUES(@d,@t,@c,@n,@m,@r1,@r3,@r7,@mx,@v,@ex,@mc,@sc) " &
                    "ON DUPLICATE KEY UPDATE name=VALUES(name),market=VALUES(market)," &
                    "ret_1m=VALUES(ret_1m),ret_3m=VALUES(ret_3m),ret_7h=VALUES(ret_7h)," &
                    "max_ret=VALUES(max_ret),search_volume=VALUES(search_volume)," &
                    "extra=VALUES(extra),market_cap=VALUES(market_cap),sector=VALUES(sector)"
                Using cmd As New MySqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@d", targetDate)
                    cmd.Parameters.AddWithValue("@t", searchTime)
                    cmd.Parameters.AddWithValue("@c", code)
                    cmd.Parameters.AddWithValue("@n", If(row.Cells("name").Value, ""))
                    cmd.Parameters.AddWithValue("@m", If(row.Cells("market").Value, ""))
                    cmd.Parameters.AddWithValue("@r1", DbHelper.ToDecimalOrNull(row.Cells("ret_1m").Value))
                    cmd.Parameters.AddWithValue("@r3", DbHelper.ToDecimalOrNull(row.Cells("ret_3m").Value))
                    cmd.Parameters.AddWithValue("@r7", DbHelper.ToDecimalOrNull(row.Cells("ret_7h").Value))
                    cmd.Parameters.AddWithValue("@mx", DbHelper.ToDecimalOrNull(row.Cells("max_ret").Value))
                    cmd.Parameters.AddWithValue("@v", DbHelper.ToLongOrNull(row.Cells("search_volume").Value))
                    cmd.Parameters.AddWithValue("@ex", DbHelper.ToDecimalOrNull(row.Cells("extra").Value))
                    cmd.Parameters.AddWithValue("@mc", DbHelper.ToLongOrNull(row.Cells("market_cap").Value))
                    cmd.Parameters.AddWithValue("@sc", If(row.Cells("sector").Value, ""))
                    cmd.ExecuteNonQuery() : savedCount += 1
                End Using
            Next
        End Using
        lblStatus.Text = $"저장 완료: {savedCount}건, 건너뜀: {skippedCount}건"
        MessageBox.Show($"{targetDate:yyyy-MM-dd} {searchTime:hh\:mm}" & vbCrLf &
                        $"저장: {savedCount}건, 건너뜀: {skippedCount}건", "저장 완료")
    End Sub

    Private Function ParseClipboard(rawText As String) As List(Of Dictionary(Of String, Object))
        Dim rows As New List(Of Dictionary(Of String, Object))
        Dim lines = rawText.Split({vbCrLf, vbLf}, StringSplitOptions.None)
        Dim headerSkipped = 0
        For Each line In lines
            Dim trimmed = line.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For
            If headerSkipped < 2 AndAlso (trimmed.Contains("종목명") OrElse trimmed.Contains("1분간")) Then
                headerSkipped += 1 : Continue For
            End If
            Dim cols = line.Split(CChar(vbTab))
            Dim cleaned As New List(Of String)
            Dim started = False
            For Each c In cols
                Dim v = c.Trim()
                If Not started AndAlso v = "" Then Continue For
                started = True : cleaned.Add(v)
            Next
            If cleaned.Count < 6 Then Continue For
            Dim row As New Dictionary(Of String, Object)
            row("name") = cleaned(0)
            row("ret_1m") = JsonParser.ParsePct(cleaned(1))
            row("ret_3m") = JsonParser.ParsePct(cleaned(2))
            row("ret_7h") = JsonParser.ParsePct(cleaned(3))
            row("max_ret") = JsonParser.ParsePct(cleaned(4))
            row("search_volume") = JsonParser.ParseVolume(cleaned(5))
            row("extra") = If(cleaned.Count > 6, JsonParser.ParsePct(cleaned(6)), DBNull.Value)
            row("code") = "??????" : row("market") = "" : row("market_cap") = DBNull.Value : row("sector") = ""
            rows.Add(row)
        Next
        Return rows
    End Function

    ' ══════════════════════════════════════════
    '  탭2: SQL 실행기 (기존 유지)
    ' ══════════════════════════════════════════
    Private Sub SetupTabSql()
        Dim pnlQ As New FlowLayoutPanel() With {.Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(5)}
        Dim lblS As New Label() With {.Text = "저장된 쿼리:", .AutoSize = True, .Margin = New Padding(0, 6, 0, 0)}
        cboSavedQueries.Width = 350 : cboSavedQueries.DropDownStyle = ComboBoxStyle.DropDownList
        btnLoadSql.Text = "불러오기" : btnLoadSql.Size = New Size(80, 28)
        btnSaveSql.Text = "쿼리 저장" : btnSaveSql.Size = New Size(80, 28)
        btnDeleteSql.Text = "삭제" : btnDeleteSql.Size = New Size(60, 28)
        pnlQ.Controls.AddRange({lblS, cboSavedQueries, btnLoadSql, btnSaveSql, btnDeleteSql})

        txtSql.Multiline = True : txtSql.ScrollBars = ScrollBars.Both
        txtSql.Dock = DockStyle.Top : txtSql.Height = 200
        txtSql.Font = New Font("Consolas", 10) : txtSql.AcceptsReturn = True
        txtSql.AcceptsTab = True : txtSql.WordWrap = False

        Dim pnlR As New FlowLayoutPanel() With {.Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(5)}
        btnRunSql.Text = "▶ SQL 실행 (F5)" : btnRunSql.Size = New Size(140, 30)
        btnRunSql.BackColor = Color.FromArgb(0, 120, 215) : btnRunSql.ForeColor = Color.White
        btnRunSql.FlatStyle = FlatStyle.Flat
        lblSqlStatus.AutoSize = True : lblSqlStatus.Margin = New Padding(20, 6, 0, 0)
        pnlR.Controls.AddRange({btnRunSql, lblSqlStatus})

        dgvSqlResult.Dock = DockStyle.Fill : dgvSqlResult.AllowUserToAddRows = False
        dgvSqlResult.ReadOnly = True : dgvSqlResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvSqlResult.Font = New Font("맑은 고딕", 9) : dgvSqlResult.AutoGenerateColumns = True
        dgvSqlResult.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 250)

        tabSql.Controls.Add(dgvSqlResult)
        tabSql.Controls.Add(pnlR)
        tabSql.Controls.Add(txtSql)
        tabSql.Controls.Add(pnlQ)
        RefreshQueryList()
        AddHandler txtSql.KeyDown, Sub(s, ev)
                                       If ev.KeyCode = Keys.F5 Then ev.Handled = True : ev.SuppressKeyPress = True : RunSql()
                                   End Sub
    End Sub

    Private Sub btnRunSql_Click(sender As Object, e As EventArgs) Handles btnRunSql.Click
        RunSql()
    End Sub

    Private Sub RunSql()
        Dim sql = txtSql.Text.Trim()
        If String.IsNullOrEmpty(sql) Then MessageBox.Show("SQL을 입력하세요.", "알림") : Return
        Dim sw As New Diagnostics.Stopwatch() : sw.Start()
        Try
            Dim dt = DbHelper.ExecuteQuery(sql) : sw.Stop()
            dgvSqlResult.DataSource = dt
            For Each col As DataGridViewColumn In dgvSqlResult.Columns
                Dim ct = dt.Columns(col.DataPropertyName).DataType
                If ct Is GetType(Decimal) OrElse ct Is GetType(Double) OrElse ct Is GetType(Long) OrElse
                   ct Is GetType(Integer) OrElse ct Is GetType(Single) Then
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    col.DefaultCellStyle.Format = "N2"
                End If
            Next
            lblSqlStatus.Text = $"결과: {dt.Rows.Count}행 × {dt.Columns.Count}열 | {sw.ElapsedMilliseconds}ms"
            lblSqlStatus.ForeColor = Color.DarkGreen
        Catch ex As Exception
            sw.Stop()
            lblSqlStatus.Text = $"오류: {ex.Message}" : lblSqlStatus.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub btnSaveSql_Click(sender As Object, e As EventArgs) Handles btnSaveSql.Click
        Dim sql = txtSql.Text.Trim()
        If String.IsNullOrEmpty(sql) Then Return
        Dim qName = InputBox("쿼리 이름:", "쿼리 저장", "")
        If String.IsNullOrEmpty(qName) Then Return
        Dim safeName = qName
        For Each c In Path.GetInvalidFileNameChars() : safeName = safeName.Replace(c, "_"c) : Next
        Dim fp = Path.Combine(QueryFolderPath, safeName & ".sql")
        If File.Exists(fp) AndAlso MessageBox.Show("덮어쓰시겠습니까?", "확인", MessageBoxButtons.YesNo) = DialogResult.No Then Return
        File.WriteAllText(fp, sql, System.Text.Encoding.UTF8)
        RefreshQueryList()
        lblSqlStatus.Text = $"저장 완료: {qName}" : lblSqlStatus.ForeColor = Color.DarkBlue
    End Sub

    Private Sub btnLoadSql_Click(sender As Object, e As EventArgs) Handles btnLoadSql.Click
        If cboSavedQueries.SelectedIndex < 0 Then Return
        Dim fp = Path.Combine(QueryFolderPath, cboSavedQueries.SelectedItem.ToString() & ".sql")
        If File.Exists(fp) Then txtSql.Text = File.ReadAllText(fp, System.Text.Encoding.UTF8)
    End Sub

    Private Sub btnDeleteSql_Click(sender As Object, e As EventArgs) Handles btnDeleteSql.Click
        If cboSavedQueries.SelectedIndex < 0 Then Return
        Dim qName = cboSavedQueries.SelectedItem.ToString()
        If MessageBox.Show($"'{qName}' 삭제?", "확인", MessageBoxButtons.YesNo) = DialogResult.No Then Return
        Dim fp = Path.Combine(QueryFolderPath, qName & ".sql")
        If File.Exists(fp) Then File.Delete(fp)
        RefreshQueryList() : txtSql.Clear()
    End Sub

    Private Sub RefreshQueryList()
        cboSavedQueries.Items.Clear()
        If Not Directory.Exists(QueryFolderPath) Then Return
        For Each fp In Directory.GetFiles(QueryFolderPath, "*.sql").OrderBy(Function(f) f)
            cboSavedQueries.Items.Add(Path.GetFileNameWithoutExtension(fp))
        Next
        If cboSavedQueries.Items.Count > 0 Then cboSavedQueries.SelectedIndex = 0
    End Sub

    ' ══════════════════════════════════════════
    '  탭3: 캔들 분석 (개편)
    ' ══════════════════════════════════════════
    Private Sub SetupTabCandle()

        ' ── 1행: 일자 선택 + 성과 데이터 불러오기 ──
        Dim pnlDateRow As New FlowLayoutPanel() With {
            .Dock = DockStyle.Top, .Height = 42, .Padding = New Padding(5)
        }
        Dim lblSelDate As New Label() With {
            .Text = "분석일자:", .AutoSize = True, .Margin = New Padding(0, 8, 0, 0),
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold)
        }
        dtpCandleDate = New DateTimePicker() With {
            .Format = DateTimePickerFormat.Short, .Value = DateTime.Today, .Width = 110
        }
        btnLoadPerfData = New Button() With {
            .Text = "성과 데이터 불러오기", .Size = New Size(140, 28),
            .BackColor = Color.FromArgb(60, 60, 60), .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat, .Margin = New Padding(10, 3, 0, 0)
        }
        AddHandler btnLoadPerfData.Click, AddressOf BtnLoadPerfData_Click
        pnlDateRow.Controls.AddRange({lblSelDate, dtpCandleDate, btnLoadPerfData})

        ' ── 2행: 종목코드 / 종목명 / 분봉·틱봉 설정 / 다운로드 ──
        Dim pnlSettingRow As New FlowLayoutPanel() With {
            .Dock = DockStyle.Top, .Height = 42, .Padding = New Padding(5)
        }
        lblCandleCode = New Label() With {
            .Text = "종목코드:", .AutoSize = True, .Margin = New Padding(0, 8, 0, 0),
            .Font = New Font("맑은 고딕", 9)
        }
        txtCandleCode = New TextBox() With {
            .Width = 75, .Text = "006010", .Font = New Font("Consolas", 10)
        }
        AddHandler txtCandleCode.Leave, Sub(s As Object, ev As EventArgs)
                                            LookupStockName()
                                        End Sub
        lblCandleName = New Label() With {
            .Text = "(종목명)", .AutoSize = True,
            .Font = New Font("맑은 고딕", 10, FontStyle.Bold),
            .Margin = New Padding(5, 8, 0, 0), .ForeColor = Color.DarkBlue
        }

        Dim lblM As New Label() With {.Text = "분봉:", .AutoSize = True, .Margin = New Padding(15, 8, 0, 0)}
        cboMinuteType = New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 48}
        cboMinuteType.Items.AddRange({"1", "3", "5", "10", "15", "30", "60"})
        cboMinuteType.SelectedIndex = 0

        Dim lblT As New Label() With {.Text = "틱봉:", .AutoSize = True, .Margin = New Padding(10, 8, 0, 0)}
        cboTickType = New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 48}
        cboTickType.Items.AddRange({"1", "3", "5", "10", "15", "30", "60", "120"})
        cboTickType.SelectedIndex = 6

        Dim lblSt As New Label() With {.Text = "Stop:", .AutoSize = True, .Margin = New Padding(10, 8, 0, 0)}
        dtpMinuteStop = New DateTimePicker() With {
            .Format = DateTimePickerFormat.Custom,
            .CustomFormat = "yyyy-MM-dd HH:mm",
            .Value = DateTime.Today.AddHours(9).AddMinutes(10),
            .Width = 135
        }

        Dim lblC As New Label() With {.Text = "수량:", .AutoSize = True, .Margin = New Padding(10, 8, 0, 0)}
        nudMinuteCount = New NumericUpDown() With {.Minimum = 1, .Maximum = 500, .Value = 30, .Width = 52}

        btnDownload = New Button() With {
            .Text = "▼ 다운로드", .Size = New Size(100, 28),
            .BackColor = Color.FromArgb(0, 120, 215), .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat, .Margin = New Padding(10, 3, 0, 0)
        }
        AddHandler btnDownload.Click, AddressOf BtnDownloadCandle_Click

        pnlSettingRow.Controls.AddRange({lblCandleCode, txtCandleCode, lblCandleName,
                                          lblM, cboMinuteType, lblT, cboTickType,
                                          lblSt, dtpMinuteStop, lblC, nudMinuteCount, btnDownload})

        ' ── 상태 라벨 ──
        lblCandleStatus = New Label() With {
            .Text = "일자를 선택하고 [성과 데이터 불러오기]를 클릭하세요.",
            .AutoSize = True, .Dock = DockStyle.Top,
            .Padding = New Padding(5, 2, 0, 2),
            .Font = New Font("맑은 고딕", 9), .ForeColor = Color.DimGray
        }

        ' ── 성과검증 종목 그리드 ──
        dgvPerfStocks = New DataGridView() With {
            .Dock = DockStyle.Fill, .ReadOnly = True,
            .AllowUserToAddRows = False, .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .RowHeadersVisible = False,
            .Font = New Font("맑은 고딕", 8.5),
            .BackgroundColor = Color.White, .BorderStyle = BorderStyle.FixedSingle,
            .AutoGenerateColumns = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        }
        dgvPerfStocks.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 248, 255)
        AddHandler dgvPerfStocks.CellClick, AddressOf DgvPerfStocks_CellClick
        RemoveHandler dgvPerfStocks.CellFormatting, AddressOf DgvPerfStocks_CellFormatting
        AddHandler dgvPerfStocks.CellFormatting, AddressOf DgvPerfStocks_CellFormatting

        Dim lblPerfTitle As New Label() With {
            .Text = "■ 성과검증 종목 (클릭 → 캔들 다운로드)", .Dock = DockStyle.Top, .Height = 22,
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(0, 80, 0), .BackColor = Color.FromArgb(230, 255, 230),
            .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(5, 0, 0, 0)
        }

        Dim pnlPerfGrid As New Panel() With {.Dock = DockStyle.Top, .Height = 160}
        pnlPerfGrid.Controls.Add(dgvPerfStocks)
        pnlPerfGrid.Controls.Add(lblPerfTitle)

        ' ── 분봉 그리드 ──
        dgvMinute = New DataGridView() With {
            .Dock = DockStyle.Fill, .ReadOnly = True,
            .AllowUserToAddRows = False, .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .RowHeadersVisible = False, .Font = New Font("맑은 고딕", 8),
            .BackgroundColor = Color.White, .BorderStyle = BorderStyle.FixedSingle
        }
        dgvMinute.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 255)

        ' ── 틱봉 그리드 ──
        dgvTick = New DataGridView() With {
            .Dock = DockStyle.Fill, .ReadOnly = True,
            .AllowUserToAddRows = False, .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .RowHeadersVisible = False, .Font = New Font("맑은 고딕", 8),
            .BackgroundColor = Color.White, .BorderStyle = BorderStyle.FixedSingle
        }
        dgvTick.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 245)

        ' ── 분봉/틱봉 SplitContainer ──
        Dim splitGrids As New SplitContainer() With {
            .Dock = DockStyle.Fill, .Orientation = Orientation.Vertical,
            .SplitterDistance = 420, .SplitterWidth = 4,
            .BorderStyle = BorderStyle.FixedSingle
        }
        Dim lblMinTitle As New Label() With {
            .Text = "■ 분봉 + 틱강도", .Dock = DockStyle.Top, .Height = 20,
            .Font = New Font("맑은 고딕", 8, FontStyle.Bold),
            .ForeColor = Color.FromArgb(0, 100, 180), .BackColor = Color.FromArgb(230, 240, 255),
            .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(5, 0, 0, 0)
        }
        splitGrids.Panel1.Controls.Add(dgvMinute)
        splitGrids.Panel1.Controls.Add(lblMinTitle)

        Dim lblTickTitle As New Label() With {
            .Text = "■ 틱봉 원본", .Dock = DockStyle.Top, .Height = 20,
            .Font = New Font("맑은 고딕", 8, FontStyle.Bold),
            .ForeColor = Color.FromArgb(180, 80, 0), .BackColor = Color.FromArgb(255, 240, 230),
            .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(5, 0, 0, 0)
        }
        splitGrids.Panel2.Controls.Add(dgvTick)
        splitGrids.Panel2.Controls.Add(lblTickTitle)

        Dim pnlGrids As New Panel() With {.Dock = DockStyle.Top, .Height = 140}
        pnlGrids.Controls.Add(splitGrids)

        ' ── 차트 패널 ──
        pnlChart = New Panel() With {
            .Dock = DockStyle.Fill, .BackColor = Color.FromArgb(20, 20, 30)
        }
        Dim pi = GetType(Panel).GetProperty("DoubleBuffered",
                    Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        If pi IsNot Nothing Then pi.SetValue(pnlChart, True)

        AddHandler pnlChart.Paint, Sub(s As Object, pe As PaintEventArgs)
                                       If chartData IsNot Nothing AndAlso chartData.Rows.Count > 0 Then
                                           ChartRenderer.Draw(pe.Graphics, pnlChart.ClientRectangle,
                                                              chartData, crosshairMousePt)
                                       End If
                                   End Sub
        AddHandler pnlChart.MouseMove, Sub(s As Object, me2 As MouseEventArgs)
                                           crosshairMousePt = me2.Location
                                           pnlChart.Invalidate()
                                       End Sub
        AddHandler pnlChart.MouseLeave, Sub(s As Object, ev As EventArgs)
                                            crosshairMousePt = Point.Empty
                                            pnlChart.Invalidate()
                                        End Sub

        Dim lblChartTitle As New Label() With {
            .Text = "■ 캔들차트 + 틱강도(MA5/MA20) + 이동평균 + 크로스헤어",
            .Dock = DockStyle.Top, .Height = 22,
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold),
            .ForeColor = Color.White, .BackColor = Color.FromArgb(40, 40, 60),
            .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(5, 0, 0, 0)
        }
        Dim pnlChartWrap As New Panel() With {.Dock = DockStyle.Fill}
        pnlChartWrap.Controls.Add(pnlChart)
        pnlChartWrap.Controls.Add(lblChartTitle)

        ' ── 차트 컨텍스트 메뉴 (지표 토글) ──
        Dim ctxChart As New ContextMenuStrip()

        Dim indicatorDefs() = {
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("MA5 (5일)", ChartRenderer.IndicatorFlags.MA5),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("MA20 (20일)", ChartRenderer.IndicatorFlags.MA20),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("MA120 (120일)", ChartRenderer.IndicatorFlags.MA120),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("SuperTrend (14,2)", ChartRenderer.IndicatorFlags.SuperTrend),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("JMA (7,50,2)", ChartRenderer.IndicatorFlags.JMA),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("RSI(14)", ChartRenderer.IndicatorFlags.RSI14),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("RSI(5)", ChartRenderer.IndicatorFlags.RSI5),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("RSI(50)", ChartRenderer.IndicatorFlags.RSI50),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("틱강도 히스토그램", ChartRenderer.IndicatorFlags.TickIntensity),
            New Tuple(Of String, ChartRenderer.IndicatorFlags)("틱강도 MA", ChartRenderer.IndicatorFlags.TickMA)
        }

        '//////////////////////////////////////
        ' 전략관련 메뉴
        '//////////////////////////////////////
        ' ── 전략 관련 메뉴 ──
        ctxChart.Items.Add(New ToolStripSeparator())

        Dim itemStratMgr As New ToolStripMenuItem("전략 관리자...")
        AddHandler itemStratMgr.Click, Sub()
                                           Using dlg As New StrategyManagerDialog()
                                               dlg.ShowDialog(Me)
                                           End Using
                                       End Sub
        ctxChart.Items.Add(itemStratMgr)

        Dim itemApplyStrat As New ToolStripMenuItem("전략 적용 (백테스트)")
        AddHandler itemApplyStrat.Click, Sub()
                                             If chartData Is Nothing OrElse chartData.Rows.Count = 0 Then
                                                 MessageBox.Show("먼저 캔들 데이터를 다운로드하세요.", "알림")
                                                 Return
                                             End If
                                             Dim allStrategies = StrategyStore.LoadAll()
                                             If allStrategies.Count = 0 Then
                                                 MessageBox.Show("저장된 전략이 없습니다. 전략 관리자에서 먼저 생성하세요.", "알림")
                                                 Return
                                             End If
                                             ' 전략 선택 메뉴
                                             Dim selectMenu As New ContextMenuStrip()
                                             For Each strat In allStrategies
                                                 Dim s = strat  ' 클로저 캡처
                                                 Dim mi As New ToolStripMenuItem(s.ToString())
                                                 AddHandler mi.Click, Sub()
                                                                          Dim perf = StrategyEngine.Evaluate(s, chartData)
                                                                          ChartRenderer.CurrentPerformance = perf
                                                                          pnlChart.Invalidate()
                                                                          lblCandleStatus.Text = $"전략 '{s.Name}' 적용 | " &
                                                                              $"거래:{perf.TotalTrades} 승률:{perf.WinRate:N1}% " &
                                                                              $"총수익:{perf.TotalReturnPct:N2}% PF:{perf.ProfitFactor:N2}"
                                                                          lblCandleStatus.ForeColor = If(perf.TotalReturnPct >= 0, Color.DarkGreen, Color.Red)
                                                                      End Sub
                                                 selectMenu.Items.Add(mi)
                                             Next
                                             selectMenu.Show(pnlChart, pnlChart.PointToClient(Cursor.Position))
                                         End Sub
        ctxChart.Items.Add(itemApplyStrat)

        Dim itemClearStrat As New ToolStripMenuItem("전략 신호 제거")
        AddHandler itemClearStrat.Click, Sub()
                                             ChartRenderer.CurrentPerformance = Nothing
                                             pnlChart.Invalidate()
                                             lblCandleStatus.Text = "전략 신호 제거됨"
                                         End Sub
        ctxChart.Items.Add(itemClearStrat)




        For Each def In indicatorDefs
            Dim flag = def.Item2
            Dim item As New ToolStripMenuItem(def.Item1)
            item.CheckOnClick = True
            item.Checked = ChartRenderer.IsActive(flag)
            AddHandler item.Click, Sub(s2 As Object, e2 As EventArgs)
                                       ChartRenderer.Toggle(flag)
                                       DirectCast(s2, ToolStripMenuItem).Checked = ChartRenderer.IsActive(flag)
                                       pnlChart.Invalidate()
                                   End Sub
            ctxChart.Items.Add(item)
        Next

        ' 구분선 + 전체 선택/해제
        ctxChart.Items.Add(New ToolStripSeparator())
        Dim itemAll As New ToolStripMenuItem("전체 선택")
        AddHandler itemAll.Click, Sub()
                                      For Each f In indicatorDefs
                                          ChartRenderer.ActiveIndicators = ChartRenderer.ActiveIndicators Or f.Item2
                                      Next
                                      For Each mi As ToolStripMenuItem In ctxChart.Items.OfType(Of ToolStripMenuItem)()
                                          mi.Checked = True
                                      Next
                                      pnlChart.Invalidate()
                                  End Sub
        ctxChart.Items.Add(itemAll)

        Dim itemNone As New ToolStripMenuItem("전체 해제")
        AddHandler itemNone.Click, Sub()
                                       ChartRenderer.ActiveIndicators = ChartRenderer.IndicatorFlags.None
                                       For Each mi As ToolStripMenuItem In ctxChart.Items.OfType(Of ToolStripMenuItem)()
                                           mi.Checked = False
                                       Next
                                       pnlChart.Invalidate()
                                   End Sub
        ctxChart.Items.Add(itemNone)

        pnlChart.ContextMenuStrip = ctxChart




        ' ── 탭에 추가 (역순 Dock.Top) ──
        tabCandle.Controls.Add(pnlChartWrap)     ' Fill
        tabCandle.Controls.Add(pnlGrids)         ' Top - 분봉/틱봉 그리드
        tabCandle.Controls.Add(pnlPerfGrid)      ' Top - 성과검증 종목
        tabCandle.Controls.Add(lblCandleStatus)   ' Top
        tabCandle.Controls.Add(pnlSettingRow)     ' Top
        tabCandle.Controls.Add(pnlDateRow)        ' Top (최상단)

    End Sub

    ' ── 종목명 조회 ──
    Private Sub LookupStockName()
        Dim code = txtCandleCode.Text.Trim()
        If code.Length < 6 Then
            lblCandleName.Text = "(코드 6자리)" : lblCandleName.ForeColor = Color.Gray
            Return
        End If
        Try
            Dim nm = DbHelper.GetStockName(code)
            If String.IsNullOrEmpty(nm) Then
                lblCandleName.Text = "(종목 없음)" : lblCandleName.ForeColor = Color.Red
            Else
                lblCandleName.Text = nm : lblCandleName.ForeColor = Color.DarkBlue
            End If
        Catch
            lblCandleName.Text = "(DB오류)" : lblCandleName.ForeColor = Color.Red
        End Try
    End Sub

    ' ── 성과 데이터 불러오기 ──
    Private Sub BtnLoadPerfData_Click(sender As Object, e As EventArgs)
        Dim targetDate = dtpCandleDate.Value.Date
        Try
            Dim sql = "SELECT code AS 종목코드, name AS 종목명, ret_1m AS `1분%`, ret_3m AS `3분%`, " &
                      "ret_7h AS `7시간%`, max_ret AS `최고%`, search_volume AS 거래량, " &
                      "market AS 시장, sector AS 업종, search_time AS 검색시각 " &
                      "FROM perf_result WHERE target_date=@d ORDER BY max_ret DESC"
            Dim dt As New DataTable()
            Using conn = DbHelper.CreateConnection()
                conn.Open()
                Using cmd As New MySqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@d", targetDate)
                    Using adapter As New MySqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using
            dgvPerfStocks.DataSource = dt
            FormatPerfGrid()

            ' StopTime을 해당 일자 09:10으로 설정
            dtpMinuteStop.Value = targetDate.AddHours(9).AddMinutes(10)

            lblCandleStatus.Text = $"{targetDate:yyyy-MM-dd} 성과 데이터: {dt.Rows.Count}종목"
            lblCandleStatus.ForeColor = Color.DarkGreen
        Catch ex As Exception
            lblCandleStatus.Text = $"오류: {ex.Message}" : lblCandleStatus.ForeColor = Color.Red
        End Try
    End Sub

    ' ── 성과 그리드 서식 ──
    Private Sub FormatPerfGrid()
        For Each col As DataGridViewColumn In dgvPerfStocks.Columns
            Select Case col.HeaderText
                Case "1분%", "3분%", "7시간%", "최고%"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    col.DefaultCellStyle.Format = "N2"
                    col.Width = 60
                Case "거래량"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    col.DefaultCellStyle.Format = "N0"
                    col.Width = 75
                Case "종목코드"
                    col.Width = 70
                Case "종목명"
                    col.Width = 100
                Case "시장"
                    col.Width = 50
                Case "업종"
                    col.Width = 110
                Case "검색시각"
                    col.Width = 70
            End Select
        Next
    End Sub

    ' ── 성과 그리드 셀 서식 ──
    Private Sub DgvPerfStocks_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 Then Return
        Dim dgv = DirectCast(sender, DataGridView)
        Dim row = dgv.Rows(e.RowIndex)

        ' 최고% >= 10 → 노란색 배경
        If dgv.Columns.Contains("최고%") Then
            Dim mrVal As Decimal = 0
            Dim mrObj = row.Cells("최고%").Value
            If mrObj IsNot Nothing AndAlso Not IsDBNull(mrObj) AndAlso Decimal.TryParse(mrObj.ToString(), mrVal) Then
                If mrVal >= 10 Then
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 210)
                    row.DefaultCellStyle.Font = New Font("맑은 고딕", 8.5F, FontStyle.Bold)
                End If
            End If
        End If

        ' 3분% 색상
        If dgv.Columns.Contains("3분%") Then
            Dim v = row.Cells("3분%").Value
            If v IsNot Nothing AndAlso Not IsDBNull(v) Then
                Dim p As Decimal = 0
                If Decimal.TryParse(v.ToString(), p) Then
                    If p >= 5 Then row.Cells("3분%").Style.ForeColor = Color.Red : row.Cells("3분%").Style.Font = New Font("맑은 고딕", 8.5F, FontStyle.Bold)
                End If
            End If
        End If
    End Sub

    ' ── 성과 그리드 종목 클릭 → 캔들 다운로드 ──
    Private Sub DgvPerfStocks_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim row = dgvPerfStocks.Rows(e.RowIndex)

        ' 종목코드 가져오기
        Dim code = ""
        If dgvPerfStocks.Columns.Contains("종목코드") Then
            code = row.Cells("종목코드").Value?.ToString()?.Trim()
        End If
        If String.IsNullOrEmpty(code) OrElse code.Length < 6 Then Return

        ' UI에 반영
        txtCandleCode.Text = code
        LookupStockName()

        ' 검색시각으로 StopTime 설정
        If dgvPerfStocks.Columns.Contains("검색시각") Then
            Dim stObj = row.Cells("검색시각").Value
            If stObj IsNot Nothing AndAlso Not IsDBNull(stObj) Then
                Dim targetDate = dtpCandleDate.Value.Date
                Try
                    Dim ts = TimeSpan.Parse(stObj.ToString())
                    ' 검색시각 + 7시간을 StopTime으로 (장중 추이 확인용)
                    dtpMinuteStop.Value = targetDate.Add(ts).AddHours(7)
                Catch
                    dtpMinuteStop.Value = targetDate.AddHours(15).AddMinutes(30)
                End Try
            End If
        End If

        ' 자동 다운로드
        BtnDownloadCandle_Click(Nothing, Nothing)
    End Sub

    ' ── 캔들 다운로드 ──
    Private Sub BtnDownloadCandle_Click(sender As Object, e As EventArgs)
        Dim code = txtCandleCode.Text.Trim()
        If String.IsNullOrEmpty(code) OrElse code.Length < 6 Then
            MessageBox.Show("종목코드를 입력하세요.", "알림") : Return
        End If

        Dim mType = cboMinuteType.SelectedItem.ToString()
        Dim tType = cboTickType.SelectedItem.ToString()
        Dim stopTime = dtpMinuteStop.Value.ToString("yyyyMMddHHmmss")
        Dim mCount = CInt(nudMinuteCount.Value)
        Dim tCount = Math.Min(500, mCount * 20)

        lblCandleStatus.Text = "다운로드 중..." : lblCandleStatus.ForeColor = Color.Gray
        Application.DoEvents()

        Try
            Dim jsonM = ApiClient.DownloadJson(ApiClient.MinuteCandleUrl(code, mType, mCount, stopTime))
            Dim dtM = JsonParser.ParseCandles(jsonM)
            ' ★ Reverse: 최신이 우측으로
            dtM = ReverseDataTable(dtM)

            Dim jsonT = ApiClient.DownloadJson(ApiClient.TickCandleUrl(code, tType, tCount, stopTime))
            Dim dtT = JsonParser.ParseCandles(jsonT)
            dtT = ReverseDataTable(dtT)

            ' 틱강도 합산
            TickIntensity.Calculate(dtM, dtT, CInt(mType))

            dgvMinute.DataSource = Nothing : dgvMinute.DataSource = dtM
            FormatMinuteGrid()
            dgvTick.DataSource = Nothing : dgvTick.DataSource = dtT
            FormatTickGrid()

            chartData = dtM : pnlChart.Invalidate()

            Dim summary = TickIntensity.GetSummary(dtM)
            lblCandleStatus.Text = $"분봉 {dtM.Rows.Count}건 | 틱봉 {dtT.Rows.Count}건 | {lblCandleName.Text}({code}) | {summary}"
            lblCandleStatus.ForeColor = Color.DarkGreen
        Catch ex As Exception
            lblCandleStatus.Text = $"오류: {ex.Message}" : lblCandleStatus.ForeColor = Color.Red
        End Try
    End Sub

    ' ── DataTable Reverse (Sort 대신) ──
    Private Function ReverseDataTable(dt As DataTable) As DataTable
        If dt Is Nothing OrElse dt.Rows.Count <= 1 Then Return dt
        Dim reversed As DataTable = dt.Clone()
        For i = dt.Rows.Count - 1 To 0 Step -1
            reversed.ImportRow(dt.Rows(i))
        Next
        Return reversed
    End Function

    ' ── 분봉 그리드 서식 ──
    Private Sub FormatMinuteGrid()
        For Each col As DataGridViewColumn In dgvMinute.Columns
            If col.Name <> "시간" Then col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            If col.Name = "거래량" OrElse col.Name = "틱거래량" Then col.DefaultCellStyle.Format = "N0"
            If col.Name = "등락%" Then col.DefaultCellStyle.Format = "N2"
            If col.Name = "틱강도" Then col.Width = 55 : col.DefaultCellStyle.Font = New Font("Consolas", 8.5F, FontStyle.Bold)
        Next
        RemoveHandler dgvMinute.CellFormatting, AddressOf MinuteGrid_Format
        AddHandler dgvMinute.CellFormatting, AddressOf MinuteGrid_Format
    End Sub

    Private Sub MinuteGrid_Format(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 Then Return
        Dim dgv = DirectCast(sender, DataGridView)
        Dim row = dgv.Rows(e.RowIndex)

        If dgv.Columns.Contains("등락%") Then
            Dim v = row.Cells("등락%").Value
            If v IsNot Nothing AndAlso Not IsDBNull(v) Then
                Dim p As Decimal = 0
                If Decimal.TryParse(v.ToString(), p) Then
                    row.DefaultCellStyle.ForeColor = If(p > 0, Color.Red, If(p < 0, Color.Blue, Color.Black))
                End If
            End If
        End If

        If dgv.Columns.Contains("틱강도") AndAlso e.ColumnIndex = dgv.Columns("틱강도").Index Then
            Dim ti As Integer = 0
            If e.Value IsNot Nothing AndAlso Not IsDBNull(e.Value) Then Integer.TryParse(e.Value.ToString(), ti)
            If ti >= 20 Then
                e.CellStyle.BackColor = Color.FromArgb(255, 100, 100) : e.CellStyle.ForeColor = Color.White
            ElseIf ti >= 10 Then
                e.CellStyle.BackColor = Color.FromArgb(255, 180, 130)
            ElseIf ti >= 5 Then
                e.CellStyle.BackColor = Color.FromArgb(255, 230, 180)
            End If
        End If
    End Sub

    ' ── 틱봉 그리드 서식 ──
    Private Sub FormatTickGrid()
        For Each col As DataGridViewColumn In dgvTick.Columns
            If col.Name <> "시간" Then col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            If col.Name = "거래량" Then col.DefaultCellStyle.Format = "N0"
            If col.Name = "등락%" Then col.DefaultCellStyle.Format = "N2"
        Next
        RemoveHandler dgvTick.CellFormatting, AddressOf TickGrid_Format
        AddHandler dgvTick.CellFormatting, AddressOf TickGrid_Format
    End Sub

    Private Sub TickGrid_Format(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 Then Return
        Dim dgv = DirectCast(sender, DataGridView)
        Dim row = dgv.Rows(e.RowIndex)
        If dgv.Columns.Contains("등락%") Then
            Dim v = row.Cells("등락%").Value
            If v IsNot Nothing AndAlso Not IsDBNull(v) Then
                Dim p As Decimal = 0
                If Decimal.TryParse(v.ToString(), p) Then
                    row.DefaultCellStyle.ForeColor = If(p > 0, Color.Red, If(p < 0, Color.Blue, Color.Black))
                End If
            End If
        End If
    End Sub

End Class
