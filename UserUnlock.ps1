<#
.SYNOPSIS
    ADアカウント一括ロック解除ツール

.DESCRIPTION
    【概要】
    CSVファイルから対象ユーザーのリスト（SamAccountName等）を読み込み、
    Active Directory のアカウントロックを一括で解除するGUIツールです。
    
    【実行方式（2つのモード）】
    本ツールは環境に合わせて以下の2つのモードをタブで切り替えて使用できます。
    
    1. RSATモード (Local Client)
       - 実行端末の「Active Directory モジュール (RSAT)」を使用して、ローカルからDCへ処理を要求します。
       - DC側のWinRM設定が不要で、よりセキュアな運用（専用の委任ユーザーによる実行など）に適しています。
    
    2. WinRMモード (Remote Execute)
       - 実行端末から対象DCへ `Invoke-Command` でリモート接続し、DC上で直接処理を実行します。
       - 実行端末にRSATをインストールする必要がありません（DC側でWinRMが有効である必要があります）。

    【使用手順】
    1. 事前準備: 
       本スクリプトと同じフォルダに `users.csv` を配置します。
       ※CSVには必ず `$script:TargetColumnName` で指定した列（デフォルトは SamAccountName）を含めてください。
    2. ツールの起動:
       バッチファイル（.bat）から起動するか、PowerShellで直接実行します。自動的に管理者権限に昇格します。
    3. モードの選択:
       環境に合わせてタブ（RSATモード または WinRMモード）を選択します。
    4. 実行:
       対象DCのFQDN、CSVパスを確認し、「▶ 実行」ボタンをクリックします。
       認証ダイアログが表示されるので、ロック解除権限を持つユーザーの資格情報を入力してください。
    5. 結果の確認:
       画面下部のリアルタイムログ、および同階層の `logs` フォルダ内のログファイルに処理結果が出力されます。

.NOTES
    Last Updated: 2026/04/07
#>

# ==========================================
# 管理者権限のチェックと自動昇格 (STAモード、非表示設定)
# ==========================================
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $argList = "-Sta -WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$($MyInvocation.MyCommand.Path)`""
    Start-Process powershell -ArgumentList $argList -Verb RunAs
    exit
}

# ==========================================
# --- ユーザー設定エリア ---
# ==========================================
$script:DefaultTargetDC = "dc01.example.com"
$script:TargetColumnName = "SamAccountName"
$script:LogDirectoryName = "logs"
# ==========================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

$script:AppDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($script:AppDir)) {
    $script:AppDir = if ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path } else { $PWD.Path }
}

# ==========================================
# 1. モダンGUI XAMLの定義
# ==========================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ADアカウント・アンロックツール" Height="900" Width="950" 
        Background="#1E1E1E" WindowStartupLocation="CenterScreen" FontFamily="Segoe UI, Meiryo, sans-serif">
    <Window.Resources>
        
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#454545"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#007ACC"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="FlatButton" TargetType="Button">
            <Setter Property="Background" Value="#333333"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="RenderTransformOrigin" Value="0.5,0.5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#444444"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#007ACC"/>
                    <Setter Property="BorderBrush" Value="#007ACC"/>
                    <Setter Property="RenderTransform">
                        <Setter.Value>
                            <ScaleTransform ScaleX="0.97" ScaleY="0.97"/>
                        </Setter.Value>
                    </Setter>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Foreground" Value="#777777"/>
                    <Setter Property="Background" Value="#252526"/>
                    <Setter Property="BorderBrush" Value="#333333"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource FlatButton}">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="BorderBrush" Value="#007ACC"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="16"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#0098FF"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#2D7D46"/>
            <Setter Property="BorderBrush" Value="#2D7D46"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#3BA35B"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="Border" Padding="25,12" Margin="0" BorderThickness="0,0,0,3" BorderBrush="Transparent" Background="Transparent" CornerRadius="5,5,0,0">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#007ACC" />
                                <Setter TargetName="Border" Property="Background" Value="#252526" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter Property="Foreground" Value="#CCCCCC" />
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#333333" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="250"/>
        </Grid.RowDefinitions>

        <TabControl Name="MainTab" Background="#252526" BorderThickness="0" Margin="0,0,0,20">
            
            <TabItem Header="RSATモード（ローカルクライアント）">
                <Grid Margin="30">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="25"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="25"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="230"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Text="対象ドメインコントローラ（DC）：" Foreground="White" VerticalAlignment="Center" FontSize="15" FontWeight="SemiBold" Grid.Row="0" Grid.Column="0"/>
                    <TextBox Name="txtDcLocal" Grid.Row="0" Grid.Column="1" Grid.ColumnSpan="2" Height="36" FontSize="15"/>
                    
                    <TextBlock Text="対象CSVファイル：" Foreground="White" VerticalAlignment="Center" FontSize="15" FontWeight="SemiBold" Grid.Row="2" Grid.Column="0"/>
                    <TextBox Name="txtCsvLocal" Grid.Row="2" Grid.Column="1" Height="36" FontSize="15" Margin="0,0,12,0"/>
                    <Button Name="btnBrowseLocal" Content="参照..." Grid.Row="2" Grid.Column="2" Style="{StaticResource FlatButton}"/>
                    
                    <TextBlock Text="RSAT状態：" Foreground="White" VerticalAlignment="Center" FontSize="15" FontWeight="SemiBold" Grid.Row="4" Grid.Column="0"/>
                    <StackPanel Orientation="Horizontal" Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2">
                        <TextBlock Name="txtRsatStatus" Text="確認をクリックしてください ->" Foreground="#AAAAAA" VerticalAlignment="Center" Width="200" FontSize="15"/>
                        <Button Name="btnCheckRsat" Content="確認" Style="{StaticResource FlatButton}" Width="100" Margin="0,0,12,0"/>
                        <Button Name="btnInstallRsat" Content="RSATを手動でインストール" Style="{StaticResource FlatButton}" IsEnabled="False" Width="Auto"/>
                    </StackPanel>

                    <Border Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" Background="#1C2F3E" BorderBrush="#005A9E" BorderThickness="1" CornerRadius="6" Padding="18" Margin="0,30,0,30" VerticalAlignment="Top">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="30"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Viewbox Grid.Column="0" VerticalAlignment="Top" HorizontalAlignment="Center" Width="20" Height="20" Margin="0,3,10,0">
                                <Path Fill="#9CDCFE" Data="M11 15h2v2h-2zm0-8h2v6h-2zm1-5C6.477 2 2 6.477 2 12s4.477 10 10 10 10-4.477 10-10S17.523 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z"/>
                            </Viewbox>
                            <StackPanel Grid.Column="1" Orientation="Vertical">
                                <TextBlock Text="必須要件と注意事項" Foreground="#9CDCFE" FontSize="15" FontWeight="Bold" Margin="0,0,0,10"/>
                                <TextBlock Text="・実行端末に RSAT (Active Directory モジュール) がインストールされている必要があります。" Foreground="#84C0F3" Margin="5,0,0,6"/>
                                <TextBlock Text="・指定する CSVファイル には「$($script:TargetColumnName)」列が必ず存在している必要があります。" Foreground="#84C0F3" Margin="5,0,0,0"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <Button Name="btnExecuteLocal" Content="▶ RSATモードで実行" Grid.Row="6" Grid.Column="1" Grid.ColumnSpan="2" HorizontalAlignment="Right" Style="{StaticResource SuccessButton}" Width="250" Height="50"/>
                </Grid>
            </TabItem>

            <TabItem Header="WinRMモード（リモート実行）">
                <Grid Margin="30">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="25"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="230"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Text="対象ドメインコントローラ（DC）：" Foreground="White" VerticalAlignment="Center" FontSize="15" FontWeight="SemiBold" Grid.Row="0" Grid.Column="0"/>
                    <TextBox Name="txtDcRemote" Grid.Row="0" Grid.Column="1" Grid.ColumnSpan="2" Height="36" FontSize="15"/>
                    
                    <TextBlock Text="対象CSVファイル：" Foreground="White" VerticalAlignment="Center" FontSize="15" FontWeight="SemiBold" Grid.Row="2" Grid.Column="0"/>
                    <TextBox Name="txtCsvRemote" Grid.Row="2" Grid.Column="1" Height="36" FontSize="15" Margin="0,0,12,0"/>
                    <Button Name="btnBrowseRemote" Content="参照..." Grid.Row="2" Grid.Column="2" Style="{StaticResource FlatButton}"/>
                    
                    <Border Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3" Background="#2B2D26" BorderBrush="#7D7A2D" BorderThickness="1" CornerRadius="6" Padding="18" Margin="0,30,0,30" VerticalAlignment="Top">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="30"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Viewbox Grid.Column="0" VerticalAlignment="Top" HorizontalAlignment="Center" Width="20" Height="20" Margin="0,3,10,0">
                                <Path Fill="#E0DCA8" Data="M11 15h2v2h-2zm0-8h2v6h-2zm1-5C6.477 2 2 6.477 2 12s4.477 10 10 10 10-4.477 10-10S17.523 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z"/>
                            </Viewbox>
                            <StackPanel Grid.Column="1" Orientation="Vertical">
                                <TextBlock Text="必須要件と注意事項" Foreground="#E0DCA8" FontSize="15" FontWeight="Bold" Margin="0,0,0,10"/>
                                <TextBlock Text="・対象のDC側で WinRM (PSRemoting) が有効化されている必要があります。自端末のRSATは不要です。" Foreground="#C7C281" Margin="5,0,0,6"/>
                                <TextBlock Text="・指定する CSVファイル には「$($script:TargetColumnName)」列が必ず存在している必要があります。" Foreground="#C7C281" Margin="5,0,0,0"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <Button Name="btnExecuteRemote" Content="▶ WinRMモードで実行" Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" HorizontalAlignment="Right" Style="{StaticResource PrimaryButton}" Width="250" Height="50"/>
                </Grid>
            </TabItem>
        </TabControl>

        <TextBlock Text="リアルタイム・ログ出力：" Foreground="#CCCCCC" FontSize="14" FontWeight="SemiBold" Grid.Row="1" Margin="6,0,0,6"/>
        <Border Grid.Row="2" Background="#0C0C0C" BorderBrush="#3E3E42" BorderThickness="1" CornerRadius="6" Padding="6">
            <TextBox Name="txtLog" VerticalContentAlignment="Top" Background="Transparent" Foreground="#10E860" FontFamily="Consolas" FontSize="14" BorderThickness="0" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" Margin="0"/>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# コントロールの割り当て
$txtDcLocal    = $window.FindName("txtDcLocal")
$txtCsvLocal   = $window.FindName("txtCsvLocal")
$btnBrowseLoc  = $window.FindName("btnBrowseLocal")
$btnExecLoc    = $window.FindName("btnExecuteLocal")

$txtRsatStatus = $window.FindName("txtRsatStatus")
$btnCheckRsat  = $window.FindName("btnCheckRsat")
$btnInstallRsat= $window.FindName("btnInstallRsat")

$txtDcRemote   = $window.FindName("txtDcRemote")
$txtCsvRemote  = $window.FindName("txtCsvRemote")
$btnBrowseRem  = $window.FindName("btnBrowseRemote")
$btnExecRem    = $window.FindName("btnExecuteRemote")

$txtLog        = $window.FindName("txtLog")
$dispatcher    = [System.Windows.Threading.Dispatcher]::CurrentDispatcher

# 初期値
$txtDcLocal.Text = $txtDcRemote.Text = $script:DefaultTargetDC
$defaultCsv = Join-Path -Path $script:AppDir -ChildPath "users.csv"
$txtCsvLocal.Text = $txtCsvRemote.Text = $defaultCsv

# ==========================================
# 共通関数 (ログ出力)
# ==========================================
function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $logMsg = "[$timestamp] $Message"
    $window.Dispatcher.Invoke([Action]{
        $txtLog.AppendText($logMsg + "`r`n")
        $txtLog.ScrollToEnd()
    })
    try {
        $logDir = Join-Path -Path $script:AppDir -ChildPath $script:LogDirectoryName
        if (-not (Test-Path -Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
        $logFile = Join-Path -Path $logDir -ChildPath "ADUnlock_$(Get-Date -Format 'yyyyMMdd').log"
        $logMsg | Out-File -FilePath $logFile -Append -Encoding UTF8
    } catch {}
    $dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

# ==========================================
# RSAT チェック & 手動インストール誘導
# ==========================================
$btnCheckRsat.Add_Click({
    Write-Log "RSATの有無を確認しています..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    $btnCheckRsat.IsEnabled = $false
    $dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    try {
        if (Get-Module -ListAvailable -Name ActiveDirectory) {
            $txtRsatStatus.Text = "インストール済み"
            $txtRsatStatus.Foreground = "#10E860"
            $btnInstallRsat.IsEnabled = $false
            Write-Log "RSAT (Active Directory モジュール) はインストールされています。"
        } else {
            $txtRsatStatus.Text = "未インストール"
            $txtRsatStatus.Foreground = "#FF6B6B"
            $btnInstallRsat.IsEnabled = $true
            Write-Log "RSAT がインストールされていません。「RSATを手動でインストール」ボタンから手順を確認してください。"
        }
    } catch {
        Write-Log "状態確認失敗: $($_.Exception.Message)"
    } finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $btnCheckRsat.IsEnabled = $true
    }
})

$btnInstallRsat.Add_Click({
    Write-Log "RSATの手動インストール手順を表示し、設定画面を開きます。"
    try {
        Start-Process "ms-settings:optionalfeatures"
        Start-Sleep -Seconds 1.5
    } catch { }

    $msg = "背後にWindowsの「オプション機能」画面を開きました。`n`n"
    $msg += "【インストール手順】`n"
    $msg += "1. 「機能を表示」または「機能の追加」ボタンをクリックします。`n"
    $msg += "2. 検索ボックスに「RSAT」または「Active Directory」と入力します。`n"
    $msg += "3. 『RSAT: Active Directory Domain Services および Lightweight Directory Services ツール』にチェックを入れます。`n"
    $msg += "4. 「次へ」 > 「インストール」の順にクリックします。`n`n"
    $msg += "※ 画面上でインストールの進捗状況（プログレスバー）が確認できます。`n"
    $msg += "※ インストールが完了したら、このツールの「確認」ボタンを再度押してください。"

    $window.Topmost = $true
    [System.Windows.MessageBox]::Show($window, $msg, "RSATの手動インストール手順", 0, 64) | Out-Null
    $window.Topmost = $false
})

# ==========================================
# メイン実行処理 (RSAT / WinRM)
# ==========================================
$BrowseAction = {
    param($targetTextBox)
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "CSV files (*.csv)|*.csv"
    $fd.InitialDirectory = $script:AppDir
    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $targetTextBox.Text = $fd.FileName }
}
$btnBrowseLoc.Add_Click({ &$BrowseAction $txtCsvLocal })
$btnBrowseRem.Add_Click({ &$BrowseAction $txtCsvRemote })

$RemoteLogic = {
    param($usersArray, $colName)
    Import-Module ActiveDirectory -ErrorAction Stop
    $results = foreach ($user in $usersArray) {
        $id = $user.$colName
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $obj = [pscustomobject]@{ User=$id; Unlocked=$false; Error="" }
        try {
            $ad = Get-ADUser -Identity $id -Properties LockedOut -ErrorAction Stop
            if ($ad.LockedOut) {
                Unlock-ADAccount -Identity $id -ErrorAction Stop
                Start-Sleep -Milliseconds 500
                if (-not (Get-ADUser -Identity $id -Properties LockedOut).LockedOut) { $obj.Unlocked = $true }
                else { $obj.Error = "State not changed" }
            } else { $obj.Error = "Unchanged (Not Locked)" }
        } catch { $obj.Error = $_.Exception.Message }
        $obj
    }
    return $results
}

function Start-UnlockProcess {
    param([string]$Mode, [string]$DC, [string]$CSV)

    if (-not (Test-Path $CSV -PathType Leaf)) { [System.Windows.MessageBox]::Show("指定されたCSVファイルが見つかりません。", "エラー", 0, 16); return }
    $users = @(Import-Csv $CSV -Encoding UTF8)
    if ($users.Count -eq 0) { [System.Windows.MessageBox]::Show("CSVファイルにデータが含まれていません。", "エラー", 0, 16); return }
    if (-not $users[0].PSObject.Properties.Match($script:TargetColumnName)) { [System.Windows.MessageBox]::Show("列名 '$($script:TargetColumnName)' がありません。", "エラー", 0, 16); return }

    if ([System.Windows.MessageBox]::Show("[$Mode] $DC に対して $($users.Count) 件のロック解除処理を実行しますか？", "確認", 1) -ne "OK") { return }
    
    $cred = Get-Credential -Message "ロック解除を実行するための資格情報を入力してください`r`n`r`nRSAT  ：アカウントロック解除を委任されたユーザー`r`nWinRM：Domain Admins権限ユーザー"
    if (-not $cred) { 
        Write-Log "資格情報の入力がキャンセルされました。処理を中止します。"
        return 
    }

    Write-Log "--------------------------------------------------"
    Write-Log "開始: モード=$Mode / 接続先=$DC"
    
    try {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $results = $null

        if ($Mode -eq "RSAT") {
            try { Import-Module ActiveDirectory -ErrorAction Stop } catch { throw "RSAT(Active Directoryモジュール)が見つかりません。" }
            $results = foreach ($u in $users) {
                $id = $u.$($script:TargetColumnName)
                if ([string]::IsNullOrWhiteSpace($id)) { continue }
                $obj = [pscustomobject]@{ User=$id; Unlocked=$false; Error="" }
                try {
                    $ad = Get-ADUser -Identity $id -Properties LockedOut -Server $DC -Credential $cred -ErrorAction Stop
                    if ($ad.LockedOut) {
                        Unlock-ADAccount -Identity $id -Server $DC -Credential $cred -ErrorAction Stop
                        Start-Sleep -Milliseconds 500
                        if (-not (Get-ADUser -Identity $id -Properties LockedOut -Server $DC -Credential $cred).LockedOut) { $obj.Unlocked = $true }
                    } else { $obj.Error = "Unchanged (Not Locked)" }
                } catch { $obj.Error = $_.Exception.Message }
                $obj
            }
        } else {
            $results = Invoke-Command -ComputerName $DC -Credential $cred -ScriptBlock $RemoteLogic -ArgumentList (,$users), $script:TargetColumnName -ErrorAction Stop
        }

        if ($results) { Write-Log ($results | Select User, Unlocked, Error | Format-Table -AutoSize | Out-String) }
        $ok = @($results | Where { $_.Unlocked }).Count
        $ng = @($results | Where { $_.Error -and $_.Error -ne "Unchanged (Not Locked)" }).Count
        Write-Log "完了: 成功=$ok / 失敗=$ng"
        [System.Windows.MessageBox]::Show("完了しました。`n成功: $ok 件 / 失敗: $ng 件", "処理完了", 0, 64)
    } catch {
        Write-Log "ERROR: $($_.Exception.Message)"
    } finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        Write-Log "--------------------------------------------------"
    }
}

$btnExecLoc.Add_Click({ Start-UnlockProcess -Mode "RSAT" -DC $txtDcLocal.Text -CSV $txtCsvLocal.Text })
$btnExecRem.Add_Click({ Start-UnlockProcess -Mode "WinRM" -DC $txtDcRemote.Text -CSV $txtCsvRemote.Text })

$window.ShowDialog() | Out-Null