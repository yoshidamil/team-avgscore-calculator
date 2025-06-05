# -*- coding: utf-8 -*-
# チーム平均スコア計算ツール
# このスクリプトはCSVファイルとExcelファイルの両方からチームスコアを読み取り、平均点を計算します

# エンコーディング設定
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 必要なモジュールの確認とインポート
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcelモジュールをインストールしています..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# ファイルパスの設定
$csvPath = ".\src\supabase\Supabase Snippet Retrieve Team Scores by User.csv"
$excelFolder = Join-Path (Get-Location) "src\form"

# CSVデータのインポート
try {
    if (-not (Test-Path $csvPath)) {
        Write-Error "CSVファイルが見つかりません: $csvPath"
        exit 1
    }
    $csvData = Import-Csv -Path $csvPath -Encoding UTF8
    Write-Host "CSVファイルを読み込みました: $csvPath"
} catch {
    Write-Error "CSVファイルの読み込みに失敗しました: $_"
    exit 1
}

# Excelデータのインポート
try {
    $excelFiles = Get-ChildItem -Path $excelFolder -Filter "*.xlsx"
    if ($excelFiles.Count -eq 0) {
        Write-Error "$excelFolder にExcelファイルが見つかりません"
        exit 1
    }
    
    $excelPath = $excelFiles[0].FullName
    Write-Host "使用するExcelファイル: $excelPath"
    
    if (-not (Test-Path $excelPath)) {
        Write-Error "Excelファイルが見つかりません: $excelPath"
        exit 1
    }
    
    $excelData = Import-Excel -Path $excelPath
    Write-Host "Excelデータの読み込みが完了しました"
} catch {
    Write-Error "Excelデータの読み込みに失敗しました: $_"
    exit 1
}

# CSVからチーム列を取得
$csvTeamColumns = $csvData[0].PSObject.Properties.Name | Where-Object { $_ -ne "email" }
Write-Host "CSVのチーム列: $($csvTeamColumns -join ', ')"

# Excelからチーム列を取得
$excelTeamColumns = @()
$excelTeamMappings = @{}

# チーム名の列を特定（例: "Team A", "Team B"など）
$excelData[0].PSObject.Properties.Name | ForEach-Object {
    $propName = $_
    if ($propName -match "Team\s+([A-Z])") {
        $excelTeamColumns += $propName
        $teamId = $matches[1]
        $excelTeamMappings["Team $teamId"] = $propName
    }
}
Write-Host "Excelのチーム列: $($excelTeamColumns -join ', ')"

# 両方のソースからチームデータを統合
$teams = @{}

# CSVのチームを追加
$csvTeamColumns | ForEach-Object {
    $teamName = $_
    if ($teamName -match "Team\s+([A-Z])") {
        $teamId = $matches[1]
        $teams[$teamId] = @{
            Id = $teamId
            CsvColumn = $teamName
            ExcelColumn = $null
            DisplayName = "Team $teamId"
        }
    }
}

# Excelのチームを追加/マージ
foreach ($teamId in $excelTeamMappings.Keys) {
    $shortId = $teamId -replace "Team\s+", ""
    
    if ($teams.ContainsKey($shortId)) {
        $teams[$shortId].ExcelColumn = $excelTeamMappings[$teamId]
    } else {
        $teams[$shortId] = @{
            Id = $shortId
            CsvColumn = $null
            ExcelColumn = $excelTeamMappings[$teamId]
            DisplayName = $teamId
        }
    }
}

# 特定されたチームの表示
Write-Host "`n特定されたチーム:"
$teamColumns = @()
foreach ($teamId in $teams.Keys | Sort-Object) {
    $teamInfo = $teams[$teamId]
    $teamColumns += $teamInfo.DisplayName
    Write-Host "- チーム $teamId"
    Write-Host "  - CSV列: $($teamInfo.CsvColumn)"
    Write-Host "  - Excel列: $($teamInfo.ExcelColumn)"
}

# チームスコアの初期化
$teamScores = @{}
foreach ($teamId in $teams.Keys) {
    $teamScores[$teamId] = @()
}

# 重複チェック用にExcelからメールアドレスを抽出
$excelEmails = @()
$excelEmailColumn = $null

try {
    # Excel列名一覧を表示
    Write-Host "`nExcel列名一覧:"
    foreach ($propName in $excelData[0].PSObject.Properties.Name) {
        Write-Host "  - '$propName'"
    }
    
    # まず、メールアドレス関連の列を直接探す
    foreach ($propName in $excelData[0].PSObject.Properties.Name) {
        $lowerName = $propName.ToLower()
        if ($lowerName -eq "email" -or $lowerName -eq "mail" -or $lowerName -eq "e-mail") {
            $excelEmailColumn = $propName
            Write-Host "メールアドレス列を見つけました: '$propName'"
            break
        }
    }
    
    # 直接の一致がない場合、日本語のメール列を探す
    if (-not $excelEmailColumn) {
        foreach ($propName in $excelData[0].PSObject.Properties.Name) {
            if ($propName -match "メ") {
                $excelEmailColumn = $propName
                Write-Host "日本語のメールアドレス列を見つけました: '$propName'"
                break
            }
        }
    }
    
    # 最後の手段として、値がメールアドレスのような形式かチェック
    if (-not $excelEmailColumn) {
        Write-Host "内容からメールアドレス列を検出しています..."
        foreach ($propName in $excelData[0].PSObject.Properties.Name) {
            $value = $excelData[0].$propName
            if ($value -and $value -match "@" -and $value -match "\.") {
                $excelEmailColumn = $propName
                Write-Host "内容からメールアドレス列を特定しました: '$propName'"
                break
            }
        }
    }
    
    # メールアドレスを抽出
    if ($excelEmailColumn) {
        $excelEmails = $excelData | ForEach-Object { $_.$excelEmailColumn } | Where-Object { $_ -and $_ -ne "" }
        Write-Host "Excelファイルには $($excelEmails.Count) 件のメールアドレスがあります (列: $excelEmailColumn)"
        if ($excelEmails.Count -gt 0) {
            Write-Host "Excelのメールアドレス例: $($excelEmails[0])"
        }
    } else {
        Write-Host "Excelファイルにメールアドレス列が見つかりません。すべてのCSVデータを処理します。"
    }
} catch {
    Write-Host "Excelからのメールアドレス抽出に失敗しました: $_"
}

# CSVからスコアを収集
Write-Host "`nCSVデータからスコアを収集しています..."
$duplicateCount = 0
$processedCount = 0

try {
    foreach ($row in $csvData) {
        $skipThisUser = $false
        $csvEmail = $row.email
        
        if ($csvEmail -and $excelEmails -contains $csvEmail) {
            Write-Host "  メールアドレス $csvEmail はExcelファイルに存在するためスキップします"
            $skipThisUser = $true
            $duplicateCount++
        }
        
        if (-not $skipThisUser) {
            $processedCount++
            $hasScore = $false
            
            foreach ($teamId in $teams.Keys) {
                $teamInfo = $teams[$teamId]
                if ($teamInfo.CsvColumn -and $row.($teamInfo.CsvColumn) -and $row.($teamInfo.CsvColumn) -ne "null") {
                    $score = $row.($teamInfo.CsvColumn)
                    if ($score -match '^\d+$') {
                        $teamScores[$teamId] += [int]$score
                        $hasScore = $true
                        Write-Host "  チーム $teamId のスコア $score を追加しました（CSVから、メール: $csvEmail）"
                    }
                }
            }
            
            if (-not $hasScore) {
                Write-Host "  メールアドレス $csvEmail の有効なスコアが見つかりません"
            }
        }
    }
} catch {
    Write-Error "CSVデータの処理中にエラーが発生しました: $_"
    exit 1
}

Write-Host "CSV処理サマリー: $processedCount 件のメールを処理、$duplicateCount 件の重複をスキップ"

# Excelからスコアを収集
Write-Host "`nExcelデータからスコアを収集しています..."
try {
    foreach ($row in $excelData) {
        # ログ用にメールアドレスを取得
        $emailValue = "N/A"
        if ($excelEmailColumn) {
            $emailValue = $row.$excelEmailColumn
        }
        
        # チームスコアを処理
        foreach ($teamId in $teams.Keys) {
            $teamInfo = $teams[$teamId]
            if ($teamInfo.ExcelColumn -and $row.($teamInfo.ExcelColumn) -and $null -ne $row.($teamInfo.ExcelColumn)) {
                $score = $row.($teamInfo.ExcelColumn)
                if ($score -match '^\d+$') {
                    $teamScores[$teamId] += [int]$score
                    Write-Host "  チーム $teamId のスコア $score を追加しました（Excelから、メール: $emailValue）"
                }
            }
        }
    }
} catch {
    Write-Error "Excelデータの処理中にエラーが発生しました: $_"
    exit 1
}

# スコア収集のサマリーを表示
Write-Host "`nスコア収集サマリー:"
foreach ($teamId in $teams.Keys | Sort-Object) {
    $scores = $teamScores[$teamId]
    $scoreCount = $scores.Count
    if ($scoreCount -gt 0) {
        $sum = ($scores | Measure-Object -Sum).Sum
        Write-Host "- チーム $teamId : $scoreCount 件のスコアを収集、合計: $sum"
    } else {
        Write-Host "- チーム $teamId : スコアなし"
    }
}

# 平均スコアの計算と表示
Write-Host "`nチーム平均スコア:"
foreach ($teamId in $teams.Keys | Sort-Object) {
    $scores = $teamScores[$teamId]
    if ($scores.Count -gt 0) {
        $average = ($scores | Measure-Object -Average).Average
        Write-Host ("チーム {0} : {1:F2} (投票数: {2})" -f $teamId, $average, $scores.Count)
    } else {
        Write-Host "チーム $teamId : スコアなし"
    }
}

# CSV出力用に結果を準備
$results = @()
foreach ($teamId in $teams.Keys | Sort-Object) {
    $scores = $teamScores[$teamId]
    if ($scores.Count -gt 0) {
        $average = ($scores | Measure-Object -Average).Average
        $obj = [PSCustomObject] @{
            Team = "Team $teamId"
            AverageScore = [math]::Round($average, 2)
            VoteCount = $scores.Count
        }
        $results += $obj
    }
}

# 平均スコアで結果をソート（降順）
$sortedResults = $results | Sort-Object -Property AverageScore -Descending

# CSVに出力
$outputPath = ".\TeamScoreResults.csv"
$sortedResults | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`n結果を保存しました: $outputPath"
