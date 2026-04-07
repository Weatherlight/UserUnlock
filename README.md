# AD Account Unlock Tool

Active Directory のアカウントロックを一括で解除する、モダンなWPF GUIを搭載したPowerShellツールです。

## 概要

本ツールは、CSVファイルからユーザーリストを読み込み、対象のドメインコントローラー(DC)に対してアカウントロック解除を一括実行します。
実行環境やセキュリティ要件に合わせて、**「RSATモード」** と **「WinRMモード」** の2つの実行方式を使い分けることができるハイブリッド仕様です。

## 主要機能

* **モダンで直感的なGUI**: PowerShellとWPF(XAML)を組み合わせた、ダークテーマの洗練されたUI。
* **2つの実行モード**:
  * **RSATモード (Local Client)**: 実行端末の RSAT (Active Directory モジュール) を使用し、ローカルからLDAP経由で安全に実行します。
  * **WinRMモード (Remote Execute)**: `Invoke-Command` を使用し、対象DCにリモート接続して実行します（実行端末にRSATは不要です）。
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
3. `Run_UserUnlock.bat` をダブルクリックして起動します。（コマンドプロンプト画面を隠し、スクリプトを安全なモードで起動します）
4. 画面が表示されたら、環境に合わせてタブ（RSATモード / WinRMモード）を選択します。
5. 「対象ドメインコントローラ」と「対象CSVファイル」を指定します。
6. 「▶ 〇〇モードで実行」ボタンをクリックし、ロック解除権限を持つアカウントの資格情報（ID/パスワード）を入力します。
7. 処理が完了すると、画面下部のログエリアおよびポップアップで結果のサマリーが表示されます。

## 設定の変更

スクリプト (`UserUnlock.ps1`) の最上部にあるユーザー設定エリアを編集することで、デフォルトの挙動を変更できます。

```powershell
# デフォルトの接続先DC
$script:DefaultTargetDC = "dc01.example.local"
# CSVの対象列名
$script:TargetColumnName = "SamAccountName"
# ログフォルダ名
$script:LogDirectoryName = "logs"