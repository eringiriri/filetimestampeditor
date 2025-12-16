# ファイルタイムスタンプエディタ
# Windows Forms GUIでファイルのタイムスタンプを編集するツール

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# フォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "ファイルタイムスタンプエディタ"
$form.Size = New-Object System.Drawing.Size(500, 350)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.AllowDrop = $true

# ファイルパス用ラベル
$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Text = "ファイル:"
$labelFile.Location = New-Object System.Drawing.Point(10, 20)
$labelFile.Size = New-Object System.Drawing.Size(50, 20)
$form.Controls.Add($labelFile)

# ファイルパス用テキストボックス
$textFilePath = New-Object System.Windows.Forms.TextBox
$textFilePath.Location = New-Object System.Drawing.Point(70, 18)
$textFilePath.Size = New-Object System.Drawing.Size(310, 20)
$textFilePath.ReadOnly = $true
$textFilePath.AllowDrop = $true
$form.Controls.Add($textFilePath)

# 参照ボタン
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "参照..."
$buttonBrowse.Location = New-Object System.Drawing.Point(390, 16)
$buttonBrowse.Size = New-Object System.Drawing.Size(80, 25)
$form.Controls.Add($buttonBrowse)

# ドロップエリアラベル
$labelDropArea = New-Object System.Windows.Forms.Label
$labelDropArea.Text = "ここにファイルをドラッグ&ドロップ"
$labelDropArea.Location = New-Object System.Drawing.Point(70, 50)
$labelDropArea.Size = New-Object System.Drawing.Size(310, 20)
$labelDropArea.ForeColor = [System.Drawing.Color]::Gray
$labelDropArea.Font = New-Object System.Drawing.Font("MS UI Gothic", 9, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($labelDropArea)

# グループボックス（タイムスタンプ編集エリア）
$groupTimestamp = New-Object System.Windows.Forms.GroupBox
$groupTimestamp.Text = "タイムスタンプ"
$groupTimestamp.Location = New-Object System.Drawing.Point(10, 80)
$groupTimestamp.Size = New-Object System.Drawing.Size(460, 170)
$form.Controls.Add($groupTimestamp)

# 作成日時ラベル
$labelCreation = New-Object System.Windows.Forms.Label
$labelCreation.Text = "作成日時:"
$labelCreation.Location = New-Object System.Drawing.Point(20, 30)
$labelCreation.Size = New-Object System.Drawing.Size(80, 20)
$groupTimestamp.Controls.Add($labelCreation)

# 作成日時 DateTimePicker
$dtpCreationDate = New-Object System.Windows.Forms.DateTimePicker
$dtpCreationDate.Location = New-Object System.Drawing.Point(110, 28)
$dtpCreationDate.Size = New-Object System.Drawing.Size(150, 20)
$dtpCreationDate.Format = "Custom"
$dtpCreationDate.CustomFormat = "yyyy/MM/dd"
$dtpCreationDate.Enabled = $false
$groupTimestamp.Controls.Add($dtpCreationDate)

$dtpCreationTime = New-Object System.Windows.Forms.DateTimePicker
$dtpCreationTime.Location = New-Object System.Drawing.Point(270, 28)
$dtpCreationTime.Size = New-Object System.Drawing.Size(100, 20)
$dtpCreationTime.Format = "Time"
$dtpCreationTime.ShowUpDown = $true
$dtpCreationTime.Enabled = $false
$groupTimestamp.Controls.Add($dtpCreationTime)

# 更新日時ラベル
$labelLastWrite = New-Object System.Windows.Forms.Label
$labelLastWrite.Text = "更新日時:"
$labelLastWrite.Location = New-Object System.Drawing.Point(20, 70)
$labelLastWrite.Size = New-Object System.Drawing.Size(80, 20)
$groupTimestamp.Controls.Add($labelLastWrite)

# 更新日時 DateTimePicker
$dtpLastWriteDate = New-Object System.Windows.Forms.DateTimePicker
$dtpLastWriteDate.Location = New-Object System.Drawing.Point(110, 68)
$dtpLastWriteDate.Size = New-Object System.Drawing.Size(150, 20)
$dtpLastWriteDate.Format = "Custom"
$dtpLastWriteDate.CustomFormat = "yyyy/MM/dd"
$dtpLastWriteDate.Enabled = $false
$groupTimestamp.Controls.Add($dtpLastWriteDate)

$dtpLastWriteTime = New-Object System.Windows.Forms.DateTimePicker
$dtpLastWriteTime.Location = New-Object System.Drawing.Point(270, 68)
$dtpLastWriteTime.Size = New-Object System.Drawing.Size(100, 20)
$dtpLastWriteTime.Format = "Time"
$dtpLastWriteTime.ShowUpDown = $true
$dtpLastWriteTime.Enabled = $false
$groupTimestamp.Controls.Add($dtpLastWriteTime)

# アクセス日時ラベル
$labelLastAccess = New-Object System.Windows.Forms.Label
$labelLastAccess.Text = "アクセス日時:"
$labelLastAccess.Location = New-Object System.Drawing.Point(20, 110)
$labelLastAccess.Size = New-Object System.Drawing.Size(80, 20)
$groupTimestamp.Controls.Add($labelLastAccess)

# アクセス日時 DateTimePicker
$dtpLastAccessDate = New-Object System.Windows.Forms.DateTimePicker
$dtpLastAccessDate.Location = New-Object System.Drawing.Point(110, 108)
$dtpLastAccessDate.Size = New-Object System.Drawing.Size(150, 20)
$dtpLastAccessDate.Format = "Custom"
$dtpLastAccessDate.CustomFormat = "yyyy/MM/dd"
$dtpLastAccessDate.Enabled = $false
$groupTimestamp.Controls.Add($dtpLastAccessDate)

$dtpLastAccessTime = New-Object System.Windows.Forms.DateTimePicker
$dtpLastAccessTime.Location = New-Object System.Drawing.Point(270, 108)
$dtpLastAccessTime.Size = New-Object System.Drawing.Size(100, 20)
$dtpLastAccessTime.Format = "Time"
$dtpLastAccessTime.ShowUpDown = $true
$dtpLastAccessTime.Enabled = $false
$groupTimestamp.Controls.Add($dtpLastAccessTime)

# 秒の説明ラベル
$labelNote = New-Object System.Windows.Forms.Label
$labelNote.Text = "※ 時刻は時:分:秒で指定できます"
$labelNote.Location = New-Object System.Drawing.Point(110, 140)
$labelNote.Size = New-Object System.Drawing.Size(250, 20)
$labelNote.ForeColor = [System.Drawing.Color]::Gray
$labelNote.Font = New-Object System.Drawing.Font("MS UI Gothic", 8)
$groupTimestamp.Controls.Add($labelNote)

# 適用ボタン
$buttonApply = New-Object System.Windows.Forms.Button
$buttonApply.Text = "適用"
$buttonApply.Location = New-Object System.Drawing.Point(300, 265)
$buttonApply.Size = New-Object System.Drawing.Size(80, 30)
$buttonApply.Enabled = $false
$form.Controls.Add($buttonApply)

# 閉じるボタン
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "閉じる"
$buttonClose.Location = New-Object System.Drawing.Point(390, 265)
$buttonClose.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($buttonClose)

# ファイル読み込み関数
function Load-FileTimestamp {
    param([string]$filePath)

    if (-not (Test-Path $filePath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show(
            "ファイルが見つかりません: $filePath",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    try {
        $fileInfo = Get-Item -LiteralPath $filePath
        $textFilePath.Text = $filePath

        # 作成日時
        $dtpCreationDate.Value = $fileInfo.CreationTime
        $dtpCreationTime.Value = $fileInfo.CreationTime

        # 更新日時
        $dtpLastWriteDate.Value = $fileInfo.LastWriteTime
        $dtpLastWriteTime.Value = $fileInfo.LastWriteTime

        # アクセス日時
        $dtpLastAccessDate.Value = $fileInfo.LastAccessTime
        $dtpLastAccessTime.Value = $fileInfo.LastAccessTime

        # コントロールを有効化
        $dtpCreationDate.Enabled = $true
        $dtpCreationTime.Enabled = $true
        $dtpLastWriteDate.Enabled = $true
        $dtpLastWriteTime.Enabled = $true
        $dtpLastAccessDate.Enabled = $true
        $dtpLastAccessTime.Enabled = $true
        $buttonApply.Enabled = $true

        $labelDropArea.Visible = $false
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "ファイルの読み込みに失敗しました: $_",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# 日付と時刻を結合する関数
function Combine-DateTime {
    param(
        [DateTime]$date,
        [DateTime]$time
    )
    return New-Object DateTime($date.Year, $date.Month, $date.Day, $time.Hour, $time.Minute, $time.Second)
}

# ドラッグ&ドロップイベント
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})

$form.Add_DragDrop({
    $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
        Load-FileTimestamp -filePath $files[0]
    }
})

# 参照ボタンクリックイベント
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "ファイルを選択"
    $openFileDialog.Filter = "すべてのファイル (*.*)|*.*"

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Load-FileTimestamp -filePath $openFileDialog.FileName
    }
})

# 適用ボタンクリックイベント
$buttonApply.Add_Click({
    $filePath = $textFilePath.Text

    if (-not (Test-Path $filePath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show(
            "ファイルが見つかりません",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    try {
        $fileInfo = Get-Item -LiteralPath $filePath

        # 作成日時を設定
        $creationTime = Combine-DateTime -date $dtpCreationDate.Value -time $dtpCreationTime.Value
        $fileInfo.CreationTime = $creationTime

        # 更新日時を設定
        $lastWriteTime = Combine-DateTime -date $dtpLastWriteDate.Value -time $dtpLastWriteTime.Value
        $fileInfo.LastWriteTime = $lastWriteTime

        # アクセス日時を設定
        $lastAccessTime = Combine-DateTime -date $dtpLastAccessDate.Value -time $dtpLastAccessTime.Value
        $fileInfo.LastAccessTime = $lastAccessTime

        [System.Windows.Forms.MessageBox]::Show(
            "タイムスタンプを更新しました",
            "完了",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "タイムスタンプの更新に失敗しました: $_",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# 閉じるボタンクリックイベント
$buttonClose.Add_Click({
    $form.Close()
})

# フォームを表示
[void]$form.ShowDialog()
