# AD Account Unlock Tool

Active Directory のアカウントロックを一括で解除する、モダンなWPF GUIを搭載したPowerShellツールです。

## 概要

本ツールは、CSVファイルからユーザーリストを読み込み、対象のドメインコントローラー(DC)に対してアカウントロック解除を一括実行します。
実行環境やセキュリティ要件に合わせて、**「RSATモード」** と **「WinRMモード」** の2つの実行方式を使い分けることができるハイブリッド仕様です。
実行に使用する資格情報は、スクリプト・バッチファイルへの事前記載、コマンドライン引数指定、または実行時のダイアログ入力の3通りから選択できます。

## 主要機能

* **モダンで直感的なGUI**: PowerShellとWPF(XAML)を組み合わせた、ダークテーマの洗練されたUI。
* **2つの実行モード**:
  * **RSATモード (Local Client)**: 実行端末の RSAT (Active Directory モジュール) を使用し、ローカルからLDAP経由で安全に実行します。
  * **WinRMモード (Remote Execute)**: `Invoke-Command` を使用し、対象DCにリモート接続して実行します（実行端末にRSATは不要です）。
* **柔軟な資格情報の指定**: スクリプト設定・バッチファイル・コマンドライン引数・実行時ダイアログの4つの方法で資格情報を指定できます。
* **リアルタイムログ出力**: 画面下部に処理の進捗をリアルタイムで表示し、同時に `logs` フォルダへファイルとして保存します。
* **RSAT状態チェック**: 実行端末のRSATインストール状態を確認し、未インストールの場合はWindowsの設定画面（オプション機能）へ誘導します。
* **管理者権限の自動昇格**: スクリプト起動時に権限をチェックし、不足している場合は自動的にUACプロンプトを表示して昇格します。
* **フェイルセーフ設計**: CSVの読み込みエラーや、個別のユーザー処理エラーが発生しても、全体を停止させずに処理を継続します。

## 動作要件

* **OS**: Windows 10 / 11, または Windows Server 2016 以降
* **PowerShell**: バージョン 5.1 以上
* **必須要件**:
  * **RSATモード利用時**: 実行端末に `RSAT: Active Directory Domain Services および Lightweight Directory Services ツール` がインストールされていること。
  * **WinRMモード利用時**: 対象のドメインコントローラーで PowerShell Remoting (`Enable-PSRemoting`) が有効になっていること。

## 入力データの要件 (CSVファイル)

処理対象のユーザーリストとして、以下の要件を満たすCSVファイルを用意してください。

* **文字コード**: UTF-8
* **必須列**: `SamAccountName` (デフォルト設定の場合)
  * ※ スクリプト上部のユーザー設定エリア (`$script:TargetColumnName`) を変更することで、他の属性（`UserPrincipalName` など）で識別することも可能です。

## 使い方

1. 本ツールをダウンロードまたはクローンし、任意のフォルダに配置します。
2. 同階層にある `users.csv` を編集し、ロック解除したいユーザーの `SamAccountName` を入力します。
3. **資格情報を事前に設定する場合（任意）**:
   * `UserUnlock.ps1` のユーザー設定エリアに `$script:DefaultUsername` / `$script:DefaultPlainPassword` を記載するか、
   * `Run_UserUnlock.bat` の設定エリアに `USERNAME` / `PLAINPASSWORD` を記載します。
   * 省略した場合は手順6で入力ダイアログが表示されます。
4. `Run_UserUnlock.bat` をダブルクリックして起動します。（コマンドプロンプト画面を隠し、スクリプトを安全なモードで起動します）
5. 画面が表示されたら、環境に合わせてタブ（RSATモード / WinRMモード）を選択します。
6. 「対象ドメインコントローラ」と「対象CSVファイル」を指定し、「▶ 〇〇モードで実行」ボタンをクリックします。
   資格情報が未設定の場合は認証ダイアログが表示されるので、ロック解除権限を持つアカウントの資格情報（ID/パスワード）を入力します。
7. 処理が完了すると、画面下部のログエリアおよびポップアップで結果のサマリーが表示されます。

## 資格情報の指定方法

資格情報は以下の優先順位で決定されます。いずれも未指定の場合は実行時にダイアログで入力を促します。

| 優先度 | 方法 | 設定箇所 |
|:---:|---|---|
| 1 | コマンドライン引数 | `-Username` / `-PlainPassword` 引数で起動 |
| 2 | バッチファイル | `Run_UserUnlock.bat` の `USERNAME` / `PLAINPASSWORD` |
| 3 | スクリプト設定 | `UserUnlock.ps1` の `$script:DefaultUsername` / `$script:DefaultPlainPassword` |
| 4 | 実行時ダイアログ | 上記すべて未指定の場合、実行ボタン押下時に表示 |

**コマンドライン引数での起動例:**

```powershell
powershell -File UserUnlock.ps1 -Username "DOMAIN\admin" -PlainPassword "P@ssw0rd"
```

## 設定の変更

`UserUnlock.ps1` の最上部にあるユーザー設定エリアを編集することで、デフォルトの挙動を変更できます。

```powershell
# デフォルトの接続先DC
$script:DefaultTargetDC    = "dc01.example.com"
# CSVの対象列名
$script:TargetColumnName   = "SamAccountName"
# ログフォルダ名
$script:LogDirectoryName   = "logs"
# 実行に使用するユーザー名（省略時はダイアログで入力）
$script:DefaultUsername    = ""   # 例: "DOMAIN\admin"
# 実行に使用するパスワード（省略時はダイアログで入力）
$script:DefaultPlainPassword = "" # 例: "P@ssw0rd"
```

`Run_UserUnlock.bat` の設定エリアでも同様に資格情報を指定できます。

```bat
set USERNAME=
set PLAINPASSWORD=
```
