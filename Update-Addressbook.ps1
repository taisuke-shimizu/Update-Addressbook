# PowerShell スクリプト：Update-Addressbook.ps1

<#
.SYNOPSIS
    アドレス帳CSVを最新のメールアドレス情報で更新し、出力ファイルを生成するスクリプト。

.DESCRIPTION
    - address.csv（ベースファイル）
    - メールアドレス取得者（all）最新.csv
    - メーリングアドレス(all).csv

    既存アドレス帳でメールアドレスが重複するレコードを削除します。
    これらのメールアドレスをキーに、退職者マーク付与、新規アドレスの追加を行い、new-address.csv を生成します。
    
    処理パターン
    address.csv     メールアドレス取得者（all）最新.csv     処理
    -----------------------------------------------------------------------------
    有              有                                    変更無し
    有              無                                    氏名の末に(退職済)を付与
    無　　　　　　   有                                    新規追加

    address.csv     メーリングアドレス(all).csv            処理
    -----------------------------------------------------------------------------
    有              有                                    変更無し
    有              無                                    氏名の末に(削除済)を付与
    無　　　　　　   有                                    新規追加

.PARAMETER address.csv
    ベースのアドレス帳ファイル。Shift-JIS形式。
    webmailからエクスポートします。

.PARAMETER メールアドレス取得者（all）最新.csv
    メールアドレス更新対象リスト。
    desknetsからダウンロードします。
    https://fmget.on.arena.ne.jp/cgi-def/dneo/zdoc.cgi?cmd=docindex&log=on&cginame=zdoc.cgi#folder=646&cmd=docrefer&id=33467

.PARAMETER メーリングアドレス(all).csv
    メーリングリスト参加者リスト。
    desknetsからダウンロードします。
    https://fmget.on.arena.ne.jp/cgi-def/dneo/zdoc.cgi?cmd=docindex&log=on&cginame=zdoc.cgi#folder=646&cmd=docrefer&id=33466

.OUTPUTS
    new-address.csv（Shift-JIS形式）

.NOTES
    実行には PowerShell 7 を推奨。
    スクリプトと同じフォルダに3つのCSVファイルを置き、以下のいずれかの方法で実行します。

.EXAMPLE
    PowerShell 7 コンソールから実行：

        pwsh ./Update-Addressbook.ps1

    または、右クリックからプログラム（powershell7）を指定して実行：

        pwsh.exe
    
.VERSION
    0.0.1
#>

# スクリプトディレクトリを取得
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# 入出力ファイルパスを絶対パスで定義
$baseFile       = Join-Path $scriptDir "address.csv"
$newFile        = Join-Path $scriptDir "メールアドレス取得者（all）最新.csv"
$mailListFile   = Join-Path $scriptDir "メーリングアドレス(all).csv"
$outputFile     = Join-Path $scriptDir "new-address.csv"
$tempFile       = "$outputFile.tmp"

# Shift-JISエンコーディング
$enc = [System.Text.Encoding]::GetEncoding("Shift-JIS")

# CSV読み込み（列名あり）
$baseData = Get-Content -Path $baseFile -Encoding $enc | ConvertFrom-Csv

# 列名を正規化
$baseData = $baseData | ForEach-Object {
    $cleanObj = @{}
    $_.PSObject.Properties | ForEach-Object {
        $key = $_.Name.Trim().Replace("　", "").Replace(" ", "")
        $cleanObj[$key] = $_.Value
    }
    [PSCustomObject]$cleanObj
}

# メールアドレス取得者（all）最新.csv 読み込み
$newRaw = Get-Content -Path $newFile -Encoding $enc | ConvertFrom-Csv -Header "ID", "氏名", "氏名ふりがな", "メールアドレス"

# メーリングアドレス(all).csv 読み込み
$mailList = Get-Content -Path $mailListFile -Encoding $enc | ConvertFrom-Csv -Header "ID", "氏名", "氏名ふりがな", "メールアドレス"

# 比較用キー（メールアドレスのみ）を作成
$newMails  = $newRaw  | ForEach-Object { $_.メールアドレス }
$mailMails = $mailList | ForEach-Object { $_.メールアドレス }

# 重複削除（@fmget.co.jp）
$exampleData = $baseData | Where-Object { $_.メールアドレス -match "@fmget.co.jp" } | Sort-Object 氏名, メールアドレス -Unique
$nonExampleData = $baseData | Where-Object { $_.メールアドレス -notmatch "@fmget.co.jp" }
$baseData = $exampleData + $nonExampleData

# 更新リスト
$updatedList = @()

foreach ($row in $baseData) {
    $email = $row.メールアドレス

    # 追加要件: @fmget.co.jp 以外のアドレス（社外アドレスなど）は処理対象外としてそのまま追加
    if ($email -notmatch "@fmget.co.jp") {
        $updatedList += $row
        continue
    }

    $existsInNew = $newMails -contains $email
    $existsInMail = $mailMails -contains $email

    if ($existsInNew -or $existsInMail) {
        $updatedList += $row
    } else {
        # どちらにも存在しない場合（無効になった社内アドレス）
        # MLかどうかの判別は完全にはできないため、"ML" や "メーリング" が名前に含まれていれば削除済、それ以外は退職済とする
        if ($row.氏名 -notmatch "\(退職済\)|\(削除済\)") {
            if ($row.氏名 -match "ML|メーリング") {
                $row.氏名 = $row.氏名 + "(削除済)"
            } else {
                $row.氏名 = $row.氏名 + "(退職済)"
            }
        }
        $updatedList += $row
    }
}

# 新規追加（メールアドレス取得者）
foreach ($newRow in $newRaw) {
    $email = $newRow.メールアドレス
    $exists = $updatedList | Where-Object { $_.メールアドレス -eq $email }
    if (-not $exists) {
        $newEntry = [PSCustomObject]@{
            'ＩＤ（システムＩＤ：自動発番）' = $newRow.ID
            '氏名'                           = $newRow.氏名
            '氏名ふりがな'                   = $newRow.氏名ふりがな
            'メールアドレス'                 = $newRow.メールアドレス
            '携帯電話'                       = ""
            '自宅-国名'                     = ""
            '自宅-郵便番号'                 = ""
            '自宅-都道府県'                 = ""
            '自宅-市区町村・番地'           = ""
            '自宅-ビル・マンション名'       = ""
            '自宅-TEL'                      = ""
            '自宅-FAX'                      = ""
            '勤務先-国名'                   = ""
            '勤務先-郵便番号'               = ""
            '勤務先-都道府県'               = ""
            '勤務先-市区町村・番地'         = ""
            '勤務先-ビル・マンション名'     = ""
            '勤務先-会社名'                 = ""
            '勤務先-会社名ふりがな'         = ""
            '勤務先-代表TEL'                = ""
            '勤務先-代表FAX'                = ""
            '勤務先-役職'                   = ""
            '勤務先-事業所'                 = ""
            '勤務先-部署名１'               = ""
            '勤務先-部署名２'               = ""
            '勤務先-部署TEL'                = ""
            '勤務先-部署FAX'                = ""
            '個人-メールアドレス（個人）'   = ""
            '個人-メールアドレス（携帯）'   = ""
            '個人-生年月日'                 = ""
            '個人-備考'                     = ""
        }
        $updatedList += $newEntry
    }
}

# 新規追加（メーリングアドレス）
foreach ($mailRow in $mailList) {
    $email = $mailRow.メールアドレス
    $exists = $updatedList | Where-Object { $_.メールアドレス -eq $email }
    if (-not $exists) {
        $newEntry = [PSCustomObject]@{
            'ＩＤ（システムＩＤ：自動発番）' = $mailRow.ID
            '氏名'                           = $mailRow.氏名
            '氏名ふりがな'                   = $mailRow.氏名ふりがな
            'メールアドレス'                 = $mailRow.メールアドレス
            '携帯電話'                       = ""
            '自宅-国名'                     = ""
            '自宅-郵便番号'                 = ""
            '自宅-都道府県'                 = ""
            '自宅-市区町村・番地'           = ""
            '自宅-ビル・マンション名'       = ""
            '自宅-TEL'                      = ""
            '自宅-FAX'                      = ""
            '勤務先-国名'                   = ""
            '勤務先-郵便番号'               = ""
            '勤務先-都道府県'               = ""
            '勤務先-市区町村・番地'         = ""
            '勤務先-ビル・マンション名'     = ""
            '勤務先-会社名'                 = ""
            '勤務先-会社名ふりがな'         = ""
            '勤務先-代表TEL'                = ""
            '勤務先-代表FAX'                = ""
            '勤務先-役職'                   = ""
            '勤務先-事業所'                 = ""
            '勤務先-部署名１'               = ""
            '勤務先-部署名２'               = ""
            '勤務先-部署TEL'                = ""
            '勤務先-部署FAX'                = ""
            '個人-メールアドレス（個人）'   = ""
            '個人-メールアドレス（携帯）'   = ""
            '個人-生年月日'                 = ""
            '個人-備考'                     = ""
        }
        $updatedList += $newEntry
    }
}

# 必要な列だけを持つオブジェクトに変換
$outputData = $updatedList | Select-Object `
    @{Name='Col1'; Expression={$_.'ＩＤ（システムＩＤ：自動発番）'}},
    @{Name='Col2'; Expression={$_.氏名}},
    @{Name='Col3'; Expression={$_.氏名ふりがな}},
    @{Name='Col4'; Expression={$_.メールアドレス}}

# ConvertTo-Csv で変換（PowerShell 7以降）
# -NoTypeInformation: 型情報を出力しない
# -UseQuotes AsNeeded: カンマを含むデータがある場合のみ " " で囲む（安全）
# Select-Object -Skip 1: 1行目のヘッダー（Col1,Col2...）を取り除く
$outputLines = $outputData | ConvertTo-Csv -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1

# 一時UTF-8ファイル（列なしCSV用）
$tempFile = "$outputFile.tmp"

# UTF-8で一時ファイル出力
[System.IO.File]::WriteAllLines($tempFile, $outputLines, [System.Text.Encoding]::UTF8)

# UTF-8 → Shift_JIS(CP932) 変換して最終ファイルに保存
$content = Get-Content -Path $tempFile
[System.IO.File]::WriteAllLines($outputFile, $content, [System.Text.Encoding]::GetEncoding("shift-jis"))

# 一時ファイル削除
Remove-Item $tempFile

Write-Host "✅ 完了：$outputFile に出力されました"

Pause
