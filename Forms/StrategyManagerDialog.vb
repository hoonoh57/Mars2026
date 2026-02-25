' ===== Forms/StrategyManagerDialog.vb =====
' 전략 CRUD 다이얼로그

Public Class StrategyManagerDialog
    Inherits Form

    Private lstStrategies As New ListBox()
    Private txtName As New TextBox()
    Private txtDesc As New TextBox()
    Private nudStopLoss As New NumericUpDown()
    Private nudTakeProfit As New NumericUpDown()
    Private nudMaxHold As New NumericUpDown()
    Private dgvBuy As New DataGridView()
    Private dgvSell As New DataGridView()
    Private WithEvents btnNew As New Button()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnClose As New Button()
    Private WithEvents btnAddBuy As New Button()
    Private WithEvents btnRemoveBuy As New Button()
    Private WithEvents btnAddSell As New Button()
    Private WithEvents btnRemoveSell As New Button()

    Private strategies As List(Of TradingStrategy)
    Private currentStrategy As TradingStrategy

    ' 사용 가능한 지표/연산자 목록
    Private ReadOnly IndicatorList() As String = {"Close", "MA5", "MA20", "MA120",
        "RSI14", "RSI5", "RSI50", "JMA", "SuperTrend", "STDir", "TickIntensity", "High", "Low", "Volume"}
    ' 기존 OperatorList를 교체
    Private ReadOnly OperatorList() As String = {
        ">", "<", ">=", "<=", "=",
        "CrossUp", "CrossDown",
        "TurnUp", "TurnDown",
        "Rising", "Falling"
    }

    Public Sub New()
        Me.Text = "전략 관리자"
        Me.Size = New Size(900, 650)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        SetupUI()
        LoadStrategies()
    End Sub

    Private Sub SetupUI()
        ' 좌측: 전략 목록
        Dim pnlList As New Panel() With {.Dock = DockStyle.Left, .Width = 200, .Padding = New Padding(5)}
        Dim lblList As New Label() With {.Text = "전략 목록", .Dock = DockStyle.Top, .Height = 22,
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold)}
        lstStrategies.Dock = DockStyle.Fill
        AddHandler lstStrategies.SelectedIndexChanged, AddressOf LstStrategies_Selected
        Dim pnlListBtn As New FlowLayoutPanel() With {.Dock = DockStyle.Bottom, .Height = 35}
        btnNew.Text = "새 전략" : btnNew.Size = New Size(60, 28)
        btnDelete.Text = "삭제" : btnDelete.Size = New Size(55, 28)
        btnClose.Text = "닫기" : btnClose.Size = New Size(55, 28)
        pnlListBtn.Controls.AddRange({btnNew, btnDelete, btnClose})
        pnlList.Controls.Add(lstStrategies)
        pnlList.Controls.Add(pnlListBtn)
        pnlList.Controls.Add(lblList)

        ' 우측: 편집 영역
        Dim pnlEdit As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(5)}

        ' 기본정보
        Dim pnlBasic As New TableLayoutPanel() With {.Dock = DockStyle.Top, .Height = 110, .ColumnCount = 4, .RowCount = 3}
        pnlBasic.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 70))
        pnlBasic.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        pnlBasic.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 70))
        pnlBasic.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))

        pnlBasic.Controls.Add(New Label() With {.Text = "전략명:", .Anchor = AnchorStyles.Left}, 0, 0)
        txtName.Dock = DockStyle.Fill : pnlBasic.Controls.Add(txtName, 1, 0)
        pnlBasic.Controls.Add(New Label() With {.Text = "설명:", .Anchor = AnchorStyles.Left}, 2, 0)
        txtDesc.Dock = DockStyle.Fill : pnlBasic.Controls.Add(txtDesc, 3, 0)

        pnlBasic.Controls.Add(New Label() With {.Text = "손절%:", .Anchor = AnchorStyles.Left}, 0, 1)
        nudStopLoss.Minimum = -50 : nudStopLoss.Maximum = 0 : nudStopLoss.DecimalPlaces = 1 : nudStopLoss.Value = -3 : nudStopLoss.Dock = DockStyle.Fill
        pnlBasic.Controls.Add(nudStopLoss, 1, 1)
        pnlBasic.Controls.Add(New Label() With {.Text = "익절%:", .Anchor = AnchorStyles.Left}, 2, 1)
        nudTakeProfit.Minimum = 0 : nudTakeProfit.Maximum = 100 : nudTakeProfit.DecimalPlaces = 1 : nudTakeProfit.Value = 10 : nudTakeProfit.Dock = DockStyle.Fill
        pnlBasic.Controls.Add(nudTakeProfit, 3, 1)

        pnlBasic.Controls.Add(New Label() With {.Text = "최대보유:", .Anchor = AnchorStyles.Left}, 0, 2)
        nudMaxHold.Minimum = 1 : nudMaxHold.Maximum = 500 : nudMaxHold.Value = 30 : nudMaxHold.Dock = DockStyle.Fill
        pnlBasic.Controls.Add(nudMaxHold, 1, 2)
        btnSave.Text = "저장" : btnSave.Size = New Size(80, 28) : btnSave.BackColor = Color.FromArgb(0, 120, 215) : btnSave.ForeColor = Color.White
        pnlBasic.Controls.Add(btnSave, 3, 2)

        ' 매수조건
        Dim pnlBuy As New Panel() With {.Dock = DockStyle.Top, .Height = 170}
        Dim lblBuy As New Label() With {.Text = "■ 매수 조건 (AND)", .Dock = DockStyle.Top, .Height = 20,
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold), .ForeColor = Color.Red}
        SetupConditionGrid(dgvBuy)
        dgvBuy.Dock = DockStyle.Fill
        Dim pnlBuyBtn As New FlowLayoutPanel() With {.Dock = DockStyle.Bottom, .Height = 30}
        btnAddBuy.Text = "+ 조건 추가" : btnAddBuy.Size = New Size(90, 25)
        btnRemoveBuy.Text = "- 삭제" : btnRemoveBuy.Size = New Size(70, 25)
        pnlBuyBtn.Controls.AddRange({btnAddBuy, btnRemoveBuy})
        pnlBuy.Controls.Add(dgvBuy) : pnlBuy.Controls.Add(pnlBuyBtn) : pnlBuy.Controls.Add(lblBuy)

        ' 매도조건
        Dim pnlSell As New Panel() With {.Dock = DockStyle.Fill}
        Dim lblSell As New Label() With {.Text = "■ 매도 조건 (AND)", .Dock = DockStyle.Top, .Height = 20,
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold), .ForeColor = Color.Blue}
        SetupConditionGrid(dgvSell)
        dgvSell.Dock = DockStyle.Fill
        Dim pnlSellBtn As New FlowLayoutPanel() With {.Dock = DockStyle.Bottom, .Height = 30}
        btnAddSell.Text = "+ 조건 추가" : btnAddSell.Size = New Size(90, 25)
        btnRemoveSell.Text = "- 삭제" : btnRemoveSell.Size = New Size(70, 25)
        pnlSellBtn.Controls.AddRange({btnAddSell, btnRemoveSell})
        pnlSell.Controls.Add(dgvSell) : pnlSell.Controls.Add(pnlSellBtn) : pnlSell.Controls.Add(lblSell)

        pnlEdit.Controls.Add(pnlSell)
        pnlEdit.Controls.Add(pnlBuy)
        pnlEdit.Controls.Add(pnlBasic)

        Me.Controls.Add(pnlEdit)
        Me.Controls.Add(pnlList)


        ' 도움말
        Dim lblHelp As New Label() With {
            .Dock = DockStyle.Bottom, .Height = 60, .Padding = New Padding(5),
            .Font = New Font("맑은 고딕", 8), .ForeColor = Color.DimGray,
            .Text = "■ 연산자 설명:" & vbCrLf &
                    "  >, <, >=, <=, = : 비교  |  CrossUp/Down : 교차  |  TurnUp : 하락→상승 전환  |  TurnDown : 상승→하락 전환" & vbCrLf &
                    "  Rising/Falling : N봉 연속 상승/하락 (Value에 봉수 입력)  |  JMA(1) = 1봉 전 JMA 값 참조"
        }
        pnlEdit.Controls.Add(lblHelp)


    End Sub

    Private Sub SetupConditionGrid(dgv As DataGridView)
        dgv.AllowUserToAddRows = False
        dgv.AutoGenerateColumns = False
        dgv.RowHeadersVisible = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.Font = New Font("맑은 고딕", 9)

        Dim colInd As New DataGridViewComboBoxColumn() With {
            .Name = "Indicator", .HeaderText = "지표", .Width = 130}
        For Each ind In IndicatorList : colInd.Items.Add(ind) : Next
        ' 이전 봉 참조 추가
        For Each ind In IndicatorList
            colInd.Items.Add(ind & "(1)")   ' 1봉 전
            colInd.Items.Add(ind & "(2)")   ' 2봉 전
        Next
        dgv.Columns.Add(colInd)

        Dim colOp As New DataGridViewComboBoxColumn() With {
            .Name = "Operator", .HeaderText = "연산자", .Width = 100}
        colOp.Items.AddRange(OperatorList)
        dgv.Columns.Add(colOp)

        Dim colTarget As New DataGridViewComboBoxColumn() With {
            .Name = "Target", .HeaderText = "대상/지표", .Width = 130}
        colTarget.Items.Add("")
        For Each ind In IndicatorList : colTarget.Items.Add(ind) : Next
        For Each ind In IndicatorList
            colTarget.Items.Add(ind & "(1)")
            colTarget.Items.Add(ind & "(2)")
        Next
        dgv.Columns.Add(colTarget)

        dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Value", .HeaderText = "값(숫자)", .Width = 80})
    End Sub


    Private Sub LoadStrategies()
        strategies = StrategyStore.LoadAll()
        lstStrategies.Items.Clear()
        For Each s In strategies : lstStrategies.Items.Add(s.Name) : Next
    End Sub

    Private Sub LstStrategies_Selected(sender As Object, e As EventArgs)
        If lstStrategies.SelectedIndex < 0 Then Return
        currentStrategy = strategies(lstStrategies.SelectedIndex)
        txtName.Text = currentStrategy.Name
        txtDesc.Text = currentStrategy.Description
        nudStopLoss.Value = CDec(Math.Max(-50, Math.Min(0, currentStrategy.StopLossPct)))
        nudTakeProfit.Value = CDec(Math.Max(0, Math.Min(100, currentStrategy.TakeProfitPct)))
        nudMaxHold.Value = Math.Max(1, Math.Min(500, currentStrategy.MaxHoldBars))
        FillConditionGrid(dgvBuy, currentStrategy.BuyConditions)
        FillConditionGrid(dgvSell, currentStrategy.SellConditions)
    End Sub

    Private Sub FillConditionGrid(dgv As DataGridView, conditions As List(Of StrategyCondition))
        dgv.Rows.Clear()
        For Each c In conditions
            Dim idx = dgv.Rows.Add()
            dgv.Rows(idx).Cells("Indicator").Value = c.Indicator
            dgv.Rows(idx).Cells("Operator").Value = c.Operator
            dgv.Rows(idx).Cells("Target").Value = If(String.IsNullOrEmpty(c.Target), "", c.Target)
            dgv.Rows(idx).Cells("Value").Value = c.Value
        Next
    End Sub

    Private Function ReadConditions(dgv As DataGridView) As List(Of StrategyCondition)
        Dim list As New List(Of StrategyCondition)
        For Each row As DataGridViewRow In dgv.Rows
            If row.IsNewRow Then Continue For
            Dim c As New StrategyCondition()
            c.Indicator = If(row.Cells("Indicator").Value?.ToString(), "")
            c.Operator = If(row.Cells("Operator").Value?.ToString(), "")
            c.Target = If(row.Cells("Target").Value?.ToString(), "")
            Dim v As Double = 0
            If row.Cells("Value").Value IsNot Nothing Then Double.TryParse(row.Cells("Value").Value.ToString(), v)
            c.Value = v
            If Not String.IsNullOrEmpty(c.Indicator) AndAlso Not String.IsNullOrEmpty(c.Operator) Then list.Add(c)
        Next
        Return list
    End Function

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        currentStrategy = New TradingStrategy() With {.Name = "새 전략_" & DateTime.Now.ToString("HHmmss")}
        txtName.Text = currentStrategy.Name : txtDesc.Text = ""
        nudStopLoss.Value = -3 : nudTakeProfit.Value = 10 : nudMaxHold.Value = 30
        dgvBuy.Rows.Clear() : dgvSell.Rows.Clear()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrWhiteSpace(txtName.Text) Then MessageBox.Show("전략명을 입력하세요.") : Return
        If currentStrategy Is Nothing Then currentStrategy = New TradingStrategy()
        currentStrategy.Name = txtName.Text.Trim()
        currentStrategy.Description = txtDesc.Text.Trim()
        currentStrategy.StopLossPct = CDbl(nudStopLoss.Value)
        currentStrategy.TakeProfitPct = CDbl(nudTakeProfit.Value)
        currentStrategy.MaxHoldBars = CInt(nudMaxHold.Value)
        currentStrategy.BuyConditions = ReadConditions(dgvBuy)
        currentStrategy.SellConditions = ReadConditions(dgvSell)
        StrategyStore.Save(currentStrategy)
        LoadStrategies()
        MessageBox.Show($"'{currentStrategy.Name}' 저장 완료", "전략 관리자")
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If currentStrategy Is Nothing Then Return
        If MessageBox.Show($"'{currentStrategy.Name}' 삭제?", "확인", MessageBoxButtons.YesNo) = DialogResult.No Then Return
        StrategyStore.Delete(currentStrategy) : currentStrategy = Nothing
        LoadStrategies() : txtName.Clear() : txtDesc.Clear() : dgvBuy.Rows.Clear() : dgvSell.Rows.Clear()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnAddBuy_Click(sender As Object, e As EventArgs) Handles btnAddBuy.Click
        dgvBuy.Rows.Add()
    End Sub
    Private Sub btnRemoveBuy_Click(sender As Object, e As EventArgs) Handles btnRemoveBuy.Click
        If dgvBuy.SelectedRows.Count > 0 Then dgvBuy.Rows.Remove(dgvBuy.SelectedRows(0))
    End Sub
    Private Sub btnAddSell_Click(sender As Object, e As EventArgs) Handles btnAddSell.Click
        dgvSell.Rows.Add()
    End Sub
    Private Sub btnRemoveSell_Click(sender As Object, e As EventArgs) Handles btnRemoveSell.Click
        If dgvSell.SelectedRows.Count > 0 Then dgvSell.Rows.Remove(dgvSell.SelectedRows(0))
    End Sub

End Class
