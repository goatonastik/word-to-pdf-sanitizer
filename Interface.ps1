param([switch]$Check)
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'LogTail.ps1')
. (Join-Path $PSScriptRoot 'ProcessLauncher.ps1')
$versionPath = Join-Path $PSScriptRoot 'VERSION.txt'
$appVersion = if (Test-Path -LiteralPath $versionPath) { (Get-Content -LiteralPath $versionPath -Raw).Trim() } else { 'unknown' }
if (!$Check) {
 $script:appMutex = New-Object Threading.Mutex($false, 'Local\WordPdfConverterUI')
 if (!$script:appMutex.WaitOne(0)) { [Windows.MessageBox]::Show('Word to PDF Sanitizer is already open.') | Out-Null; exit }
}
[xml]$layout = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Word to PDF Sanitizer" Width="1120" Height="900" MinWidth="980" MinHeight="760" WindowStartupLocation="CenterScreen" Background="#F4F6FA">
 <Grid Margin="24"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="100"/><RowDefinition Height="Auto"/><RowDefinition Height="150"/></Grid.RowDefinitions>
 <TextBlock Margin="0,0,0,16" Text="Direct Word export with background verification. Keep your originals organized." Foreground="#455468"/>
 <StackPanel Grid.Row="1">
  <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="24"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <StackPanel><TextBlock Text="Starting directory" FontWeight="SemiBold"/><DockPanel Margin="0,6,0,8"><Button Name="BrowseRoot" Content="Browse..." DockPanel.Dock="Right" Width="80" Margin="8,0,0,0"/><TextBox Name="Root" Height="34" VerticalContentAlignment="Center" AllowDrop="True" ToolTip="Drag a folder here, paste its path, or browse."/></DockPanel></StackPanel>
   <StackPanel Grid.Column="2"><TextBlock Text="Originals directory" FontWeight="SemiBold"/><DockPanel Margin="0,6,0,8"><Button Name="BrowseArchive" Content="Browse..." DockPanel.Dock="Right" Width="80" Margin="8,0,0,0"/><TextBox Name="Archive" Height="34" VerticalContentAlignment="Center" AllowDrop="True" ToolTip="A separate folder for unchanged Word originals."/></DockPanel></StackPanel>
  </Grid>
  <TextBlock Name="Destination" TextWrapping="Wrap" Foreground="#455468" Margin="0,0,0,12"/>
 </StackPanel>
 <ScrollViewer Name="OptionsScroll" Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Margin="0,0,0,12">
  <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="28"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <StackPanel>
    <TextBlock Text="Privacy and export" FontSize="17" FontWeight="SemiBold" Margin="0,0,0,10"/>
    <CheckBox Name="PreserveVisibleRevisions" IsChecked="True" Margin="0,4" ToolTip="Enabled: preserve saved visible redlines and accept hidden revisions. Disabled: accept all edits and discard all revision history in the PDF export copy. Originals retain their history."><TextBlock Text="Preserve tracked changes only when visible by default" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="HideText" IsChecked="True" Margin="0,4" ToolTip="Remove hidden-font ranges from the export copy. Off preserves their saved hidden formatting; it does not reveal text."><TextBlock Text="Exclude hidden text (remove from the in-memory export copy)" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="HideProperties" IsChecked="True" Margin="0,4"><TextBlock Text="Exclude Word document properties from PDF export" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="RemoveAllComments" IsChecked="True" Margin="0,4"><TextBlock Text="Remove ALL comments (including visible comments)" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="StripMetadata" IsChecked="True" Margin="0,4"><TextBlock Text="Remove PDF properties and added document information, including XMP" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="RemoveLinks" IsChecked="True" Margin="0,4"><TextBlock Text="Remove hyperlinks and their destinations (visible text stays)" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="ExcludeTags" IsChecked="True" Margin="0,4"><TextBlock Text="Exclude accessibility tags and alternative text (reduces screen-reader support)" TextWrapping="Wrap"/></CheckBox>
   </StackPanel>
   <StackPanel Grid.Column="2">
    <TextBlock Text="Processing and review" FontSize="17" FontWeight="SemiBold" Margin="0,0,0,10"/>
    <DockPanel Margin="0,4"><TextBlock Text="Verification queue limit:" VerticalAlignment="Center"/><TextBox Name="QueueLimit" Text="0" Width="65" Margin="10,0,0,0" HorizontalAlignment="Left"/></DockPanel>
    <TextBlock Text="0 = unlimited; positive number = maximum pending PDFs" TextWrapping="Wrap" Foreground="#455468" Margin="0,0,0,8"/>
    <DockPanel Margin="0,4"><TextBlock Text="Verification workers:" VerticalAlignment="Center"/><TextBox Name="VerifierWorkers" Text="0" Width="65" Margin="10,0,0,0" HorizontalAlignment="Left"/></DockPanel>
    <TextBlock Text="0 = automatic; 1–16 = fixed parallel PDF workers" TextWrapping="Wrap" Foreground="#455468" Margin="0,0,0,8"/>
    <CheckBox Name="ShowReviewPopup" IsChecked="True" Margin="0,4"><TextBlock Text="Show pop up dialogue box if files flagged for review" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="WarnExistingWord" IsChecked="True" Margin="0,4"><TextBlock Text="Warn before converting when Microsoft Word is already open" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="DebugLogging" IsChecked="True" Margin="0,4" ToolTip="Records operation stages, paths, lengths, worker activity and exceptions. Does not deliberately log document text or property values."><TextBlock Text="Enable debug logging (for troubleshooting)" TextWrapping="Wrap"/></CheckBox>
    <CheckBox Name="Sound" Content="Completion ding" IsChecked="True" Margin="0,4"/>
    <WrapPanel Margin="0,10,0,8"><Button Name="OpenReview" Content="Open #REVIEW folder" Padding="10,5" Margin="0,0,8,6"/><Button Name="OpenLogs" Content="Open logs" Padding="10,5" Margin="0,0,8,6"/><Button Name="PrivacyReport" Content="Privacy Excel" Padding="10,5" Margin="0,0,0,6" IsEnabled="False"/></WrapPanel>
    <TextBlock TextWrapping="Wrap" Foreground="#455468" Text="Original Word files remain unchanged. Every PDF is sanitized and verified before its original is archived. Unexpected attachments, scripts and interactive content are rejected."/>
   </StackPanel>
  </Grid>
 </ScrollViewer>
 <StackPanel Grid.Row="3" Name="RunPanel">
  <WrapPanel Margin="0,0,0,10"><Button Name="Preview" Content="Preview files" Padding="16,8" Margin="0,0,10,0"/><Button Name="Start" Content="Convert and archive" Padding="16,8" Background="#215AC8" Foreground="White"/></WrapPanel>
  <ProgressBar Name="Progress" Height="16" Minimum="0" Maximum="100" Margin="0,0,0,6"/>
  <TextBlock Name="Status" Text="Ready" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis" ClipToBounds="True" ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}" Margin="0,0,0,4"/>
  <TextBlock Name="Counts" Text="Preview or start to count documents." TextWrapping="NoWrap" TextTrimming="CharacterEllipsis" ClipToBounds="True" ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}" Margin="0,0,0,8"/>
  <WrapPanel Margin="0,0,0,6"><Button Name="CopyLog" Content="Copy displayed log to clipboard" Padding="10,6" Margin="0,0,8,4"/><Button Name="RestoreOriginals" Content="Replace converted documents with originals" Padding="10,6" Margin="0,0,8,4"/><Button Name="DeleteLogs" Content="Delete logs..." Padding="10,6" Margin="0,0,0,4"/></WrapPanel>
  <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,10"><TextBlock Text="Reports:" VerticalAlignment="Center" Margin="0,0,8,0"/><ComboBox Name="Reports" Width="360" Height="30" VerticalContentAlignment="Center" Margin="0,0,8,0" ToolTip="Generated Excel reports; newest is selected."/><Button Name="OpenReport" Content="Open report" Padding="12,5" Margin="0,0,8,0" IsEnabled="False"/><Button Name="OpenDetailedReport" Content="Open detailed report" Padding="12,5" IsEnabled="False"/></StackPanel>
 </StackPanel>
 <TextBox Grid.Row="4" Name="Log" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Background="White" Padding="10"/>
 </Grid>
</Window>
'@
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $layout))
$window.Title = 'Word to PDF Sanitizer v' + $appVersion
$controls = @{}
foreach ($name in 'Root','Archive','BrowseRoot','BrowseArchive','Destination','HideText','HideProperties','PreserveVisibleRevisions','RemoveAllComments','StripMetadata','Preview','Start','Status','Log','Progress','Counts','Sound','OpenLogs','RemoveLinks','ExcludeTags','QueueLimit','VerifierWorkers','PrivacyReport','ShowReviewPopup','OpenReview','DebugLogging') { $controls[$name] = $window.FindName($name) }
foreach($name in 'CopyLog','RestoreOriginals','DeleteLogs','Reports','OpenReport','OpenDetailedReport','WarnExistingWord') { $controls[$name]=$window.FindName($name) }
$settingsPath = Join-Path $PSScriptRoot 'preferences.json'
$script:running = $null
function Save-Preferences {
 [int]$validQueue = 0
 if (![int]::TryParse($controls.QueueLimit.Text, [ref]$validQueue) -or $validQueue -lt 0) { $controls.QueueLimit.Text='0' }
 [int]$validWorkers = 0
 if (![int]::TryParse($controls.VerifierWorkers.Text,[ref]$validWorkers) -or $validWorkers -lt 0 -or $validWorkers -gt 16) { $controls.VerifierWorkers.Text='0' }
 $values = @{root=$controls.Root.Text; archive=$controls.Archive.Text; hideText=[bool]$controls.HideText.IsChecked; preserveVisibleRevisions=[bool]$controls.PreserveVisibleRevisions.IsChecked; removeAllComments=[bool]$controls.RemoveAllComments.IsChecked; stripMetadata=[bool]$controls.StripMetadata.IsChecked; excludeProperties=[bool]$controls.HideProperties.IsChecked; removeLinks=[bool]$controls.RemoveLinks.IsChecked; excludeTags=[bool]$controls.ExcludeTags.IsChecked; queueLimit=[int]$controls.QueueLimit.Text; verifierWorkers=[int]$controls.VerifierWorkers.Text; showReviewPopup=[bool]$controls.ShowReviewPopup.IsChecked; debugLogging=[bool]$controls.DebugLogging.IsChecked; schema=3; sound=[bool]$controls.Sound.IsChecked}
 $values.warnExistingWord=[bool]$controls.WarnExistingWord.IsChecked
 $values | ConvertTo-Json | Set-Content -LiteralPath $settingsPath -Encoding UTF8
}
function Update-Destination {
 try {
  $rootText=$controls.Root.Text.Trim().Trim('"');$archiveText=$controls.Archive.Text.Trim().Trim('"')
  $leaf=[IO.Path]::GetFileName($rootText.TrimEnd('\'))
  $controls.Destination.Text='Archive layout: '+$archiveText.TrimEnd('\')+'\'+$leaf+'\[same subfolders]\document.docx'
  if($archiveText) { $controls.Destination.Text+="`nReview: "+(Join-Path ([IO.Directory]::GetParent([IO.Path]::GetFullPath($archiveText.TrimEnd('\'))).FullName) ('#REVIEW\'+$leaf+'\[same subfolders]')) }
 } catch { $controls.Destination.Text='Enter valid Starting and Originals folder paths.' }
}
function Get-QueueRootLabel($name) {
 if($name.Length -le 100) { return $name }
 $prefix=$name.Substring(0,80)
 if([char]::IsHighSurrogate($prefix[$prefix.Length-1])) { $prefix=$prefix.Substring(0,$prefix.Length-1) }
 $hash=[Security.Cryptography.SHA256]::Create()
 try { $suffix=([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($name)))).Replace('-','').Substring(0,10).ToLowerInvariant() } finally { $hash.Dispose() }
 return $prefix+' ['+$suffix+']'
}
if (Test-Path -LiteralPath $settingsPath) {
 try {
  $saved = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
  $controls.Root.Text = $saved.root; $controls.Archive.Text = $saved.archive
  if ($saved.schema -eq 3) {
   foreach ($name in 'HideText','PreserveVisibleRevisions','RemoveAllComments','StripMetadata','RemoveLinks','ExcludeTags') { if ($null -ne $saved.$name) { $controls[$name].IsChecked=[bool]$saved.$name } }
   $controls.HideProperties.IsChecked=[bool]$saved.excludeProperties
   if ($null -ne $saved.queueLimit) { $controls.QueueLimit.Text=[string]$saved.queueLimit }
   if ($null -ne $saved.verifierWorkers) { $controls.VerifierWorkers.Text=[string]$saved.verifierWorkers }
  }
  foreach ($name in 'Sound','ShowReviewPopup','DebugLogging','WarnExistingWord') { if ($null -ne $saved.$name) { $controls[$name].IsChecked = [bool]$saved.$name } }
 } catch {}
}
foreach ($name in 'Root','Archive') {
 $controls[$name].Add_PreviewDragOver({ param($sender,$e) $e.Effects = 'None'; if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) { $e.Effects = 'Copy' }; $e.Handled = $true })
 $controls[$name].Add_PreviewDrop({ param($sender,$e) $paths = $e.Data.GetData([Windows.DataFormats]::FileDrop); if ($paths.Count -eq 1 -and [IO.Directory]::Exists($paths[0])) { $sender.Text = $paths[0] }; $e.Handled = $true })
 $controls[$name].Add_TextChanged({ param($sender,$e)
  $value=$sender.Text.Trim()
  if($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) { $sender.Text=$value.Substring(1,$value.Length-2);$sender.CaretIndex=$sender.Text.Length }
  Update-Destination
 })
}
function Browse-Folder($box) {
 $dialog=$null
 try { $dialog=New-Object Windows.Forms.FolderBrowserDialog;if($dialog.ShowDialog() -eq 'OK') { $box.Text=$dialog.SelectedPath } } catch { [Windows.MessageBox]::Show($_.Exception.Message,'Browse folder') | Out-Null } finally { if($dialog) { $dialog.Dispose() } }
}
$controls.BrowseRoot.Add_Click({ Browse-Folder $controls.Root })
$controls.BrowseArchive.Add_Click({ Browse-Folder $controls.Archive })
$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(400)
$script:logsDir = Join-Path $PSScriptRoot 'Logs'
function Refresh-Reports {
 $found = New-Object 'System.Collections.Generic.List[IO.FileInfo]'
 $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
 function Add-ReportFile([IO.FileInfo]$file) {
  if($file -and $file.Extension -eq '.xlsx' -and $seen.Add($file.FullName)) { $found.Add($file) }
 }
 try {
  if(Test-Path -LiteralPath $script:logsDir) {
   foreach($file in Get-ChildItem -LiteralPath $script:logsDir -Recurse -File -Filter '*.xlsx' -ErrorAction SilentlyContinue) { Add-ReportFile $file }
  }
  $archiveText=$controls.Archive.Text.Trim().Trim('"')
  if($archiveText -and (Test-Path -LiteralPath $archiveText -PathType Container)) {
   $reportPattern=' - (conversion|review|privacy|scan)-report(?: FAILED)?(?: ERRORS)?\.xlsx$'
   $folders=New-Object 'System.Collections.Generic.List[string]';$folders.Add([IO.Path]::GetFullPath($archiveText))
   foreach($folder in Get-ChildItem -LiteralPath $archiveText -Directory -ErrorAction SilentlyContinue) { $folders.Add($folder.FullName) }
   $parent=[IO.Directory]::GetParent([IO.Path]::GetFullPath($archiveText.TrimEnd('\')))
   if($parent) {
    $reviewBase=Join-Path $parent.FullName '#REVIEW'
    if(Test-Path -LiteralPath $reviewBase -PathType Container) {
     $folders.Add($reviewBase)
     foreach($folder in Get-ChildItem -LiteralPath $reviewBase -Directory -ErrorAction SilentlyContinue) { $folders.Add($folder.FullName) }
    }
   }
   foreach($folder in $folders) {
    foreach($file in Get-ChildItem -LiteralPath $folder -File -Filter '*.xlsx' -ErrorAction SilentlyContinue) {
     if($file.Name -match $reportPattern) { Add-ReportFile $file }
    }
   }
  }
 } catch {
  try { [void][IO.Directory]::CreateDirectory($script:logsDir);Add-Content -LiteralPath (Join-Path $script:logsDir 'interface-unhandled-errors.log') -Value ('Report list refresh: '+$_.Exception.ToString()) } catch {}
 }
 $controls.Reports.Items.Clear()
 foreach($file in $found | Sort-Object LastWriteTimeUtc -Descending) {
  $item=New-Object Windows.Controls.ComboBoxItem
  $item.Content=$file.Name+'  —  '+$file.Directory.Name
  $item.Tag=$file.FullName;$item.ToolTip=$file.FullName
  [void]$controls.Reports.Items.Add($item)
 }
 if($controls.Reports.Items.Count -gt 0) { $controls.Reports.SelectedIndex=0 }
 $controls.OpenReport.IsEnabled=($controls.Reports.SelectedIndex -ge 0)
 $controls.OpenDetailedReport.IsEnabled=($controls.Reports.SelectedIndex -ge 0)
}
$controls.Reports.Add_SelectionChanged({ $available=($controls.Reports.SelectedIndex -ge 0);$controls.OpenReport.IsEnabled=$available;$controls.OpenDetailedReport.IsEnabled=$available })
$controls.OpenReport.Add_Click({
 try {
  $item=$controls.Reports.SelectedItem
  if(!$item -or !(Test-Path -LiteralPath ([string]$item.Tag) -PathType Leaf)) { throw 'Select an available report first.' }
  Start-Process -FilePath ([string]$item.Tag)
 } catch { [Windows.MessageBox]::Show($_.Exception.Message,'Open report') | Out-Null;Refresh-Reports }
})
function Invoke-ReviewManager([string[]]$arguments) {
 $python=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
 $manager=Join-Path $PSScriptRoot 'review_manager.py'
 $output=& $python -X utf8 $manager @arguments 2>&1 | Out-String
 if($LASTEXITCODE -ne 0) { throw $output.Trim() }
 return $output.Trim()
}
function Show-ReportLockFailure($owner,$path,$message) {
 [xml]$busyXml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Report is in use" Width="560" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner">
 <StackPanel Margin="24"><TextBlock Name="Message" TextWrapping="Wrap" Margin="0,0,0,18"/><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button Name="Abort" Content="Abort" Padding="20,7" Margin="0,0,10,0"/><Button Name="Retry" Content="Retry" Padding="20,7" IsDefault="True"/></StackPanel></StackPanel>
</Window>
'@
 $dialog=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $busyXml));$dialog.Owner=$owner;$dialog.Tag=$false
 $dialog.FindName('Message').Text="The selected report could not be locked. Close it in Excel or any other program, then retry.`n`n$path`n`n$message"
 $dialog.FindName('Abort').Add_Click({$dialog.Tag=$false;$dialog.Close()});$dialog.FindName('Retry').Add_Click({$dialog.Tag=$true;$dialog.Close()})
 [void]$dialog.ShowDialog();return [bool]$dialog.Tag
}
function Lock-ReportWithRetry($owner,$path) {
 while($true) {
  try { return [IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None) }
  catch { if(!(Show-ReportLockFailure $owner $path $_.Exception.Message)) { return $null } }
 }
}
function New-DetailedReviewDialog {
 [xml]$detailXml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Detailed document review" Width="1360" Height="650" MinWidth="1050" MinHeight="420" WindowStartupLocation="CenterOwner" Background="#F4F6FA">
 <Grid Margin="20"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <TextBlock Name="Heading" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis" Margin="0,0,0,12" FontSize="15" FontWeight="SemiBold"/>
  <DataGrid Grid.Row="1" Name="ReviewGrid" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" HeadersVisibility="Column" SelectionMode="Single" Background="White" RowHeight="38">
   <DataGrid.Columns>
    <DataGridTemplateColumn Header="Original document" Width="2.2*"><DataGridTemplateColumn.CellTemplate><DataTemplate><Button Content="{Binding originalLabel}" Tag="{Binding original}" CommandParameter="open" ToolTip="{Binding original}" HorizontalContentAlignment="Left" Background="Transparent" BorderThickness="0" Padding="6"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
    <DataGridTemplateColumn Header="Converted version" Width="2.2*"><DataGridTemplateColumn.CellTemplate><DataTemplate><Button Content="{Binding convertedLabel}" Tag="{Binding pdf}" CommandParameter="open" ToolTip="{Binding pdf}" HorizontalContentAlignment="Left" Background="Transparent" BorderThickness="0" Padding="6"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
    <DataGridTemplateColumn Header="Why?" Width="52"><DataGridTemplateColumn.CellTemplate><DataTemplate><TextBlock Text="?" ToolTip="{Binding reason}" FontSize="20" FontWeight="Bold" Foreground="#215AC8" HorizontalAlignment="Center" VerticalAlignment="Center"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
    <DataGridTemplateColumn Header="Accept" Width="105"><DataGridTemplateColumn.CellTemplate><DataTemplate><Button Content="✓ Accept" Foreground="#147A35" Tag="{Binding id}" CommandParameter="accept" IsEnabled="{Binding canDecide}" Padding="7,4" Margin="4"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
    <DataGridTemplateColumn Header="Revert" Width="105"><DataGridTemplateColumn.CellTemplate><DataTemplate><Button Content="✕ Revert" Foreground="#B42318" Tag="{Binding id}" CommandParameter="revert" IsEnabled="{Binding canDecide}" Padding="7,4" Margin="4"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
    <DataGridTemplateColumn Header="Undo" Width="86"><DataGridTemplateColumn.CellTemplate><DataTemplate><Button Content="Undo" Tag="{Binding id}" CommandParameter="undo" IsEnabled="{Binding canUndo}" Padding="7,4" Margin="4"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
    <DataGridTextColumn Header="Decision" Binding="{Binding decisionDisplay}" Width="2.4*"/>
   </DataGrid.Columns>
  </DataGrid>
  <Grid Grid.Row="2" Margin="0,12,0,0"><TextBlock Name="DetailStatus" Text="Choose Accept or Revert for each document. Double-clicking is not required; click a filename to open it." VerticalAlignment="Center" TextWrapping="Wrap" Margin="0,0,110,0"/><Button Name="Close" Content="Close" Width="90" Padding="14,7" HorizontalAlignment="Right" IsDefault="True"/></Grid>
 </Grid>
</Window>
'@
 return [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $detailXml))
}
function New-ExistingWordDialog($count) {
 [xml]$noticeXml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Microsoft Word is already open" Width="570" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner">
 <StackPanel Margin="24"><TextBlock Name="Message" TextWrapping="Wrap" Margin="0,0,0,18"/><CheckBox Name="DontWarnAgain" Content="Don't warn me again" Margin="0,0,0,18"/><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button Name="Cancel" Content="Cancel" Padding="20,7" Margin="0,0,10,0" IsCancel="True"/><Button Name="Continue" Content="Continue" Padding="20,7"/></StackPanel></StackPanel>
</Window>
'@
 $dialog=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $noticeXml));$dialog.Tag=$false
 $dialog.FindName('Message').Text="Detected $count existing Word process(es). Conversion will use a separate Word instance and leave your open documents alone.`n`nWord preferences are saved for your Windows account. Which Word instance closes last can affect future preferences, particularly if conversion is interrupted. The converter restores its settings on normal shutdown.`n`nAvoid opening or editing files in the Starting directory while they are being converted. Continue?"
 return $dialog
}
function Confirm-ExistingWord($owner=$window) {
 if(!$controls.WarnExistingWord.IsChecked){return $true}
 $existing=@(Get-Process WINWORD -ErrorAction SilentlyContinue)
 if(!$existing.Count){return $true}
 $notice=New-ExistingWordDialog $existing.Count;$notice.Owner=$owner
 $notice.FindName('Cancel').Add_Click({$notice.Tag=$false;$notice.Close()})
 $notice.FindName('Continue').Add_Click({$notice.Tag=$true;$notice.Close()})
 [void]$notice.ShowDialog()
 if($notice.Tag -and $notice.FindName('DontWarnAgain').IsChecked){$controls.WarnExistingWord.IsChecked=$false;Save-Preferences}
 return [bool]$notice.Tag
}
$controls.WarnExistingWord.Add_Click({try {Save-Preferences} catch {[Windows.MessageBox]::Show($_.Exception.Message,'Save preferences')|Out-Null}})
function Show-DetailedReview($reportPath) {
 if($script:running -and !$script:running.HasExited) { [Windows.MessageBox]::Show('Wait for the current conversion or maintenance operation to finish.','Detailed review') | Out-Null;return }
 if(!$reportPath -or !(Test-Path -LiteralPath $reportPath -PathType Leaf)) { [Windows.MessageBox]::Show('Select an available Excel report first.','Detailed review') | Out-Null;return }
 $lock=Lock-ReportWithRetry $window $reportPath;if(!$lock){return}
 $snapshot=Join-Path ([IO.Path]::GetTempPath()) ('word-pdf-review-'+[Guid]::NewGuid().ToString('N')+'.xlsx')
 try {
  $target=[IO.File]::Open($snapshot,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
  try {$lock.Position=0;$lock.CopyTo($target);$target.Flush()} finally {$target.Dispose()}
  $json=Invoke-ReviewManager @('list','--report',$snapshot,'--identity-report',$reportPath)
  $entries=@($json | ConvertFrom-Json)
  if($entries.Count -eq 0) { [Windows.MessageBox]::Show('The selected report contains no documents flagged for review. Choose its review, conversion, or privacy report if another report type is selected.','Detailed review') | Out-Null;return }
  $detail=New-DetailedReviewDialog;$detail.Owner=$window;$grid=$detail.FindName('ReviewGrid');$status=$detail.FindName('DetailStatus')
  $detail.FindName('Heading').Text=[IO.Path]::GetFileName($reportPath);$grid.ItemsSource=$entries
  $grid.AddHandler([Windows.Controls.Button]::ClickEvent,[Windows.RoutedEventHandler]{param($sender,$event)
   $button=$event.OriginalSource
   while($button -and !($button -is [Windows.Controls.Button])){$button=[Windows.Media.VisualTreeHelper]::GetParent($button)}
   if(!$button -or !$button.CommandParameter){return}
   $action=[string]$button.CommandParameter
   if($action -eq 'open') { try { if(Test-Path -LiteralPath ([string]$button.Tag)){Start-Process -FilePath ([string]$button.Tag)}else{throw 'The selected file is not currently available.'} } catch {[Windows.MessageBox]::Show($_.Exception.Message,'Open document')|Out-Null};return }
   $row=@($entries|Where-Object {$_.id -eq [string]$button.Tag})[0];if(!$row){return}
   if($action -eq 'undo' -and $row.decision -eq 'Reverted' -and !(Confirm-ExistingWord $detail)){return}
   try {
    $detail.Cursor=[Windows.Input.Cursors]::Wait;$status.Text= if($action -eq 'undo' -and $row.decision -eq 'Reverted'){'Converting this document again and verifying it...'}else{'Applying '+$action+'...'}
    [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Render)
    $answer=Invoke-ReviewManager @('action','--report',$reportPath,'--id',$row.id,'--action',$action,'--document',$row.document,'--pdf',$row.pdf,'--root',$controls.Root.Text.Trim().Trim('"'),'--originals',$controls.Archive.Text.Trim().Trim('"'),'--settings',$settingsPath) | ConvertFrom-Json
    foreach($name in 'decision','decisionTime','decisionResult','decisionDisplay','canDecide','canUndo'){$row.$name=$answer.$name}
    $grid.ItemsSource=$null;$grid.ItemsSource=$entries;$status.Text=$answer.decisionResult
   } catch {[Windows.MessageBox]::Show($_.Exception.Message,'Detailed review action')|Out-Null;$status.Text='Action did not complete. Review the error and file state before retrying.'}
   finally {$detail.Cursor=[Windows.Input.Cursors]::Arrow}
  })
  $detail.FindName('Close').Add_Click({$detail.Close()})
  [void]$detail.ShowDialog()
 } finally {
  if($lock){$lock.Dispose()}
  if(Test-Path -LiteralPath $snapshot){Remove-Item -LiteralPath $snapshot -Force}
 }
 while($true) {
  try { [void](Invoke-ReviewManager @('apply','--report',$reportPath));break }
  catch { if(!(Show-ReportLockFailure $window $reportPath $_.Exception.Message)){break} }
 }
 Refresh-Reports
}
$controls.OpenDetailedReport.Add_Click({$item=$controls.Reports.SelectedItem;if($item){Show-DetailedReview ([string]$item.Tag)}})
$controls.Archive.Add_LostKeyboardFocus({ Refresh-Reports })
function Read-Shared($path) {
 if (!(Test-Path -LiteralPath $path)) { return '' }
 $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
 $reader = New-Object IO.StreamReader($stream)
 try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function New-ReviewDialog {
 [xml]$dialogXml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Documents need review" Width="540" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner">
 <StackPanel Margin="24"><TextBlock Name="Message" TextWrapping="Wrap" FontSize="15" Margin="0,0,0,14"/><CheckBox Name="DontShow" Content="Do not show this again" Margin="0,0,0,18"/><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button Name="Open" Content="Open #REVIEW folder" Padding="14,7" Margin="0,0,10,0"/><Button Name="Detailed" Content="Detailed review" Padding="14,7" Margin="0,0,10,0"/><Button Name="OK" Content="OK" Padding="22,7" IsDefault="True"/></StackPanel></StackPanel>
</Window>
'@
 $dialog=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $dialogXml))
 return $dialog
}
function Show-ReviewNotification($count,$folder,$report) {
 $dialog=New-ReviewDialog
 $dialog.Owner=$window
 $dialog.FindName('Message').Text=('{0} document(s) were added to REVIEW during this batch and need manual review.' -f $count)+"`n`n"+$folder+"`n`nReasons and copy locations are recorded in the Excel reports."
 $dialog.FindName('Open').Add_Click({ if(Test-Path -LiteralPath $folder) { Start-Process explorer.exe -ArgumentList ('"'+$folder+'"') } })
 $dialog.FindName('Detailed').Add_Click({$dialog.Close();Show-DetailedReview $report})
 $dialog.FindName('OK').Add_Click({ $dialog.Close() })
 [void]$dialog.ShowDialog()
 if($dialog.FindName('DontShow').IsChecked) { $controls.ShowReviewPopup.IsChecked=$false; Save-Preferences }
}
$controls.OpenReview.Add_Click({
 try {
  if([string]::IsNullOrWhiteSpace($controls.Archive.Text)) { throw 'Choose an Originals directory first.' }
  $folder=Join-Path ([IO.Directory]::GetParent([IO.Path]::GetFullPath($controls.Archive.Text.Trim().Trim('"').TrimEnd('\'))).FullName) '#REVIEW'
  if(Test-Path -LiteralPath $folder) { Start-Process explorer.exe -ArgumentList ('"'+$folder+'"') }
  else { [Windows.MessageBox]::Show('No #REVIEW folder has been created for this Originals directory yet.','REVIEW') | Out-Null }
 } catch { [Windows.MessageBox]::Show($_.Exception.Message,'Open review folder') | Out-Null }
})
$controls.OpenLogs.Add_Click({ try { [void][IO.Directory]::CreateDirectory($script:logsDir);Start-Process explorer.exe -ArgumentList ('"'+$script:logsDir+'"') } catch { [Windows.MessageBox]::Show($_.Exception.Message,'Open logs') | Out-Null } })
$controls.PrivacyReport.Add_Click({
 try { if ($script:lastPrivacyReport -and (Test-Path -LiteralPath $script:lastPrivacyReport)) { Start-Process -FilePath $script:lastPrivacyReport } }
 catch { [Windows.MessageBox]::Show($_.Exception.Message, 'Open privacy report') | Out-Null }
})
function Run-Conversion($preview) {
 try {
  if (!$controls.Root.Text.Trim() -or !$controls.Archive.Text.Trim()) { throw 'Choose both folders first.' }
  [int]$queue = 0; if (![int]::TryParse($controls.QueueLimit.Text, [ref]$queue) -or $queue -lt 0) { throw 'Queue limit must be 0 or a positive whole number.' }
  [int]$workers = 0; if (![int]::TryParse($controls.VerifierWorkers.Text,[ref]$workers) -or $workers -lt 0 -or $workers -gt 16) { throw 'Verification workers must be 0 (automatic) or a whole number from 1 to 16.' }; Save-Preferences
  if(!$preview -and !(Confirm-ExistingWord)){ $controls.Status.Text='Conversion cancelled before starting.';return }
  $python = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
  if (!(Test-Path -LiteralPath $python)) { throw 'The bundled Python runtime is missing.' }
  New-Item -ItemType Directory -Path $script:logsDir -Force | Out-Null
  $rootName=[IO.Path]::GetFileName([IO.Path]::GetFullPath($controls.Root.Text.Trim()).TrimEnd('\'))
  $phoenix=[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow,'US Mountain Standard Time')
  $script:runId=(Get-QueueRootLabel $rootName)+' ['+$phoenix.ToString('yyyy-MM-dd HH-mm-ss.fff')+']'
  $script:outLog = Join-Path $script:logsDir ($script:runId + ' - output FAILED.log')
  $script:errLog = Join-Path $script:logsDir ($script:runId + ' - console-errors FAILED.log')
  $script:progressPath = Join-Path $script:logsDir ($script:runId + ' - progress FAILED.json')
  $script:interfaceLog = Join-Path $script:logsDir ($script:runId + ' - interface-errors FAILED.log')
  $script:uiHadErrors=$false
  $script:logTails = @((New-Object WordPdfLogTail($script:outLog)), (New-Object WordPdfLogTail($script:errLog)))
  $script:isPreview = $preview
  $script:latestProgress=$null
  $argsList = @('-X','utf8','-u', ('"' + (Join-Path $PSScriptRoot 'convert.py') + '"'), '--root', ('"' + $controls.Root.Text.Trim().TrimEnd('\') + '\."'), '--archive', ('"' + $controls.Archive.Text.Trim().TrimEnd('\') + '\."'), '--settings', ('"' + $settingsPath + '"'), '--preserve-root')
  if ($controls.Root.Text.Contains('"') -or $controls.Archive.Text.Contains('"')) { throw 'Enter folder paths without quotation marks.' }
  if ($preview) { $argsList += '--preview' }
  $argsList += @('--progress', ('"' + $script:progressPath + '"'))
  $argsList += @('--queue-name', ('"' + $script:runId + '"'))
  $script:running = New-Object WordPdfChild($python,($argsList -join ' '),$PSScriptRoot,$script:outLog,$script:errLog)
  $script:maintenanceMode=$false
  $script:lastPrivacyReport=$null
  $controls.RestoreOriginals.IsEnabled=$false; $controls.DeleteLogs.IsEnabled=$false
  foreach ($name in 'Start','Preview','Root','Archive','BrowseRoot','BrowseArchive','HideText','HideProperties','PreserveVisibleRevisions','RemoveAllComments','StripMetadata','RemoveLinks','ExcludeTags','QueueLimit','VerifierWorkers','PrivacyReport','DebugLogging') { $controls[$name].IsEnabled = $false }
  $controls.Progress.Value = 0; $controls.Progress.IsIndeterminate = $true
  $controls.Log.Clear()
  $controls.Counts.Text = 'Counting documents...'
  $controls.Status.Text = 'Scanning folders...'; $timer.Start()
 } catch {
  $details=$_.Exception.ToString()
  $startupLog=Join-Path $script:logsDir ('startup-failure-'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff')+' FAILED ERRORS.log')
  try { New-Item -ItemType Directory -Path $script:logsDir -Force | Out-Null; $details | Set-Content -LiteralPath $startupLog -Encoding UTF8 } catch {}
  [Windows.MessageBox]::Show($_.Exception.Message+"`n`nStartup details: "+$startupLog, 'Word to PDF Sanitizer') | Out-Null
 }
}
$controls.Preview.Add_Click({ Run-Conversion $true })
$controls.Start.Add_Click({ Run-Conversion $false })
$controls.CopyLog.Add_Click({
 try { [Windows.Clipboard]::SetDataObject([string]$controls.Log.Text,$true); $controls.Status.Text='Displayed log copied to clipboard.' }
 catch { [Windows.MessageBox]::Show($_.Exception.Message,'Copy log') | Out-Null }
})
function Set-MaintenanceBusy($busy) {
 foreach($name in 'Start','Preview','RestoreOriginals','DeleteLogs','Root','Archive','BrowseRoot','BrowseArchive') { $controls[$name].IsEnabled=!$busy }
 if(!$busy) { $timer.Stop();$controls.Progress.IsIndeterminate=$false;$script:running=$null;$script:maintenanceMode=$false;$script:maintenancePlanning=$false }
}
function Run-Maintenance($action) {
 if($script:running -and !$script:running.HasExited) { [Windows.MessageBox]::Show('Wait for the current queue to finish.','Maintenance') | Out-Null;return }
 try {
  if(!$controls.Root.Text.Trim() -or !$controls.Archive.Text.Trim()) { throw 'Select Starting and Originals directories first.' }
  $python=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
  $scriptPath=Join-Path $PSScriptRoot 'maintenance.py';$maintenanceDir=Join-Path $script:logsDir 'Maintenance'
  [void][IO.Directory]::CreateDirectory($maintenanceDir)
  $stamp=Get-Date -Format 'yyyyMMdd-HHmmss-fff';$planPath=Join-Path $maintenanceDir ($action+'-'+$stamp+'.plan')
  if($controls.Root.Text.Contains('"') -or $controls.Archive.Text.Contains('"')) { throw 'Enter valid folder paths.' }
  $script:maintenanceContext=@{Python=$python;ScriptPath=$scriptPath;Directory=$maintenanceDir;Stamp=$stamp;PlanPath=$planPath;Action=$action}
  $script:outLog=Join-Path $maintenanceDir ($stamp+'-plan-output.log');$script:errLog=Join-Path $maintenanceDir ($stamp+'-plan-error.log')
  $script:interfaceLog=Join-Path $maintenanceDir ($stamp+'-interface.log');$script:uiHadErrors=$false
  $arguments='-X utf8 -u "'+$scriptPath+'" plan-'+$action+' --root "'+$controls.Root.Text.Trim().TrimEnd('\')+'\." --originals "'+$controls.Archive.Text.Trim().TrimEnd('\')+'\." --plan "'+$planPath+'"'
  $script:running=New-Object WordPdfChild($python,$arguments,$PSScriptRoot,$script:outLog,$script:errLog)
  $script:maintenanceMode=$true;$script:maintenancePlanning=$true;$script:isPreview=$true;$script:latestProgress=$null
  $script:logTails=@((New-Object WordPdfLogTail($script:outLog)),(New-Object WordPdfLogTail($script:errLog)))
  Set-MaintenanceBusy $true
  $controls.Log.Clear();$controls.Progress.IsIndeterminate=$true;$controls.Status.Text='Checking files and preparing a plan...';$controls.Counts.Text='No files will change until you confirm the plan.';$timer.Start()
 } catch { Set-MaintenanceBusy $false;[Windows.MessageBox]::Show($_.Exception.Message,'Maintenance') | Out-Null }
}
function Complete-MaintenancePlan {
 try {
  $context=$script:maintenanceContext;$action=$context.Action
  if($script:running.ExitCode -ne 0) { throw (Read-Shared $script:errLog) }
  $plan=(Read-Shared $context.PlanPath) | ConvertFrom-Json
  if($action -eq 'restore') {
   if($plan.ready -eq 0) { [Windows.MessageBox]::Show('No eligible originals can be restored. Blocked: '+$plan.blocked+'; already restored: '+$plan.alreadyRestored+"`nPlan: "+$context.PlanPath,'Restore originals') | Out-Null;Set-MaintenanceBusy $false;return }
   $message='Restore '+$plan.ready+' Word original(s), remove their matching converted PDFs, and remove the verified archive copies?'+"`n`nBlocked records: "+$plan.blocked+"`nPlan: "+$context.PlanPath
   if($plan.legacy -gt 0) { $message+="`n`n"+$plan.legacy+' older conversion(s) have no PDF checksum. Their recorded matching PDFs will be deleted, including any later PDF edits.' }
  } else {
   if($plan.count -eq 0) { [Windows.MessageBox]::Show('No eligible logs found.','Delete logs') | Out-Null;Set-MaintenanceBusy $false;return }
   $message='Delete '+$plan.count+' log file(s) from application Logs and the selected originals/review folders?'+"`n`nWord documents, PDFs and restoration records will be kept.`nPlan: "+$context.PlanPath
  }
  if([Windows.MessageBox]::Show($message,'Confirm '+$action,'YesNo','Warning','No') -ne 'Yes') { Set-MaintenanceBusy $false;$controls.Status.Text='Maintenance cancelled. No files changed.';return }
  $script:outLog=Join-Path $context.Directory ($context.Stamp+'-output.log');$script:errLog=Join-Path $context.Directory ($context.Stamp+'-error.log')
  $arguments='-X utf8 -u "'+$context.ScriptPath+'" '+$action+' --plan "'+$context.PlanPath+'"'
  if($action -eq 'restore' -and $plan.legacy -gt 0) { $arguments+=' --allow-legacy' }
  $script:running=New-Object WordPdfChild($context.Python,$arguments,$PSScriptRoot,$script:outLog,$script:errLog)
  $script:maintenancePlanning=$false
  $script:logTails=@((New-Object WordPdfLogTail($script:outLog)),(New-Object WordPdfLogTail($script:errLog)))
  $controls.Status.Text='Maintenance running...'
 } catch { Set-MaintenanceBusy $false;$controls.Status.Text='Maintenance could not start.';[Windows.MessageBox]::Show($_.Exception.Message,'Maintenance') | Out-Null }
}
$controls.RestoreOriginals.Add_Click({ Run-Maintenance 'restore' })
$controls.DeleteLogs.Add_Click({ Run-Maintenance 'delete-logs' })
$timer.Add_Tick({
 try {
 $exited = $script:running.HasExited
 if($exited) { $script:running.WaitForExit() }
 if($exited -and $script:running.LoggingError) {
  $script:uiHadErrors=$true
  Add-Content -LiteralPath $script:interfaceLog -Value $script:running.LoggingError
 }
 $appended = $false; $moreLogData = $false
 foreach ($tail in $script:logTails) {
  $newText = $tail.ReadNew($exited)
  if ($newText.Length -gt 0) { $controls.Log.AppendText($newText); $appended=$true }
  if ($tail.HasMore) { $moreLogData=$true }
 }
 if ($appended) { $controls.Log.ScrollToEnd() }
 if (!$script:maintenanceMode) {
  $p=$null
  if($exited -and $script:running.QueueResultJson) { try { $p=$script:running.QueueResultJson | ConvertFrom-Json } catch {} }
  try { if(!$p -and (Test-Path -LiteralPath $script:progressPath)) { $p=(Read-Shared $script:progressPath) | ConvertFrom-Json } } catch {
   try { Add-Content -LiteralPath $script:interfaceLog -Value ('Progress read delayed: '+$_.Exception.Message) } catch {}
  }
  if($p) {
  $script:latestProgress=$p
  if ($p.privacyReport -and [IO.Path]::GetExtension($p.privacyReport) -eq '.xlsx' -and (Test-Path -LiteralPath $p.privacyReport)) { $script:lastPrivacyReport=$p.privacyReport; $controls.PrivacyReport.IsEnabled=$true }
  $controls.Progress.IsIndeterminate = ($p.phase -eq 'Scanning folders')
  if ($p.total -gt 0) { $controls.Progress.Value = 100 * $p.done / $p.total }
  $controls.Status.Text = $p.phase
  if ($p.current -and !$p.finished) { $controls.Status.Text += ': ' + [IO.Path]::GetFileName($p.current) }
  $controls.Counts.Text = '{0} / {1} processed | Converted: {2} | Skipped: {3} | Failed: {4} | PDF only: {5}' -f $p.done,$p.total,$p.converted,$p.skipped,$p.failed,$p.partial
  if ($null -ne $p.exported) { $controls.Counts.Text += ' | Exported: ' + $p.exported + ' | Pending checks: ' + $p.queued }
  if ($p.unprocessed -gt 0) { $controls.Counts.Text += ' | Not attempted: ' + $p.unprocessed }
  if ($p.reviewFlagged -gt 0) { $controls.Counts.Text += ' | Review: '+$p.reviewFlagged+' ('+$p.reviewAdded+' copies added)' }
  if ($p.reviewCopyErrors -gt 0) { $controls.Counts.Text += ' | REVIEW COPY ERRORS: '+$p.reviewCopyErrors }
 }
 }
 if ($exited -and !$moreLogData -and $script:maintenanceMode -and $script:maintenancePlanning) { Complete-MaintenancePlan;return }
 if ($exited -and !$moreLogData -and $script:maintenanceMode) {
  $timer.Stop(); $controls.Progress.IsIndeterminate=$false
  $controls.Status.Text='Maintenance finished - review the displayed results.'
  Set-MaintenanceBusy $false
  Refresh-Reports
  return
 }
 if ($exited -and !$moreLogData) {
  $timer.Stop(); $controls.Progress.IsIndeterminate = $false
  $completed=$script:running.ExitCode -eq 0 -and [bool]$script:latestProgress.queueCompleted
  $hadErrors=$script:uiHadErrors -or [bool]$script:latestProgress.queueErrors -or $script:running.ExitCode -ne 0
  $suffix=''; if(!$completed) { $suffix+=' FAILED' }; if($hadErrors) { $suffix+=' ERRORS' }
  foreach($pair in @(@($script:outLog,'output','log'),@($script:errLog,'console-errors','log'),@($script:progressPath,'progress','json'),@($script:interfaceLog,'interface-errors','log'))) {
   if(Test-Path -LiteralPath $pair[0]) {
    $newPath=Join-Path $script:logsDir ($script:runId+' - '+$pair[1]+$suffix+'.'+$pair[2])
    if($newPath -ne $pair[0]) { try { [IO.File]::Move($pair[0],$newPath) } catch { $controls.Log.AppendText("`nCould not rename log: "+$_.Exception.Message) } }
   }
  }
  $controls.Status.Text = 'Finished - review counts and report below. Ready for another batch.'
  if ($script:running.ExitCode -ne 0) { $controls.Status.Text = 'Process stopped with an error - see the log below.' }
  foreach ($name in 'Start','Preview','Root','Archive','BrowseRoot','BrowseArchive','HideText','HideProperties','PreserveVisibleRevisions','RemoveAllComments','StripMetadata','RemoveLinks','ExcludeTags','QueueLimit','VerifierWorkers','PrivacyReport','DebugLogging') { $controls[$name].IsEnabled = $true }
  $controls.PrivacyReport.IsEnabled=[bool]($script:lastPrivacyReport -and (Test-Path -LiteralPath $script:lastPrivacyReport))
  $script:running = $null
  $controls.RestoreOriginals.IsEnabled=$true; $controls.DeleteLogs.IsEnabled=$true
  Refresh-Reports
  if (!$script:isPreview -and $controls.Sound.IsChecked) { [System.Media.SystemSounds]::Asterisk.Play() }
  if (!$script:isPreview -and $controls.ShowReviewPopup.IsChecked -and $script:latestProgress.reviewAdded -gt 0) {
   Show-ReviewNotification $script:latestProgress.reviewAdded $script:latestProgress.reviewRoot $script:latestProgress.reviewReport
  }
 }
 } catch {
  # A log can be changing while read. Keep the desktop window alive and retry.
  $script:uiHadErrors=$true
  try { Add-Content -LiteralPath $script:interfaceLog -Value ((Get-Date -Format o) + ' ' + $_.Exception.ToString()) } catch {}
 }
})
$window.Add_Closing({ param($sender,$e) if ($script:running -and !$script:running.HasExited) { $e.Cancel = $true; [Windows.MessageBox]::Show('Please allow the current batch to finish before closing.', 'Conversion running') | Out-Null } else { Save-Preferences } })
Update-Destination
Refresh-Reports
$window.Dispatcher.Add_UnhandledException({ param($sender,$e)
 $e.Handled=$true
 try { [void][IO.Directory]::CreateDirectory($script:logsDir);Add-Content -LiteralPath (Join-Path $script:logsDir 'interface-unhandled-errors.log') -Value $e.Exception.ToString() } catch {}
 [Windows.MessageBox]::Show($e.Exception.Message,'Word to PDF Sanitizer') | Out-Null
})
if ($Check) {
 if($window.Title -ne ('Word to PDF Sanitizer v'+$appVersion)) { throw 'Window title does not contain the current version.' }
 if(!$controls.Reports -or !$controls.OpenReport -or !$controls.OpenDetailedReport) { throw 'Report picker controls are missing.' }
 foreach($node in $layout.SelectNodes('//*[local-name()="TextBlock"]')) { if($node.Text -eq 'Word to PDF Sanitizer') { throw 'Redundant in-window program name remains.' } }
 [void](New-ReviewDialog)
 $wordNotice=New-ExistingWordDialog 2
 if(!$wordNotice.FindName('Cancel') -or !$wordNotice.FindName('Continue') -or [bool]$wordNotice.Tag){throw 'Existing Word notice must support Continue/Cancel and default to cancellation.'}
 if(!$wordNotice.FindName('DontWarnAgain') -or !$controls.WarnExistingWord){throw 'Existing Word warning preference controls are missing.'}
 $previousWarning=$controls.WarnExistingWord.IsChecked
 $controls.WarnExistingWord.IsChecked=$false
 if(!(Confirm-ExistingWord)){throw 'Disabled warning must permit continuation without a popup.'}
 $controls.WarnExistingWord.IsChecked=$previousWarning
 [void](New-DetailedReviewDialog)
 if(!(Test-Path -LiteralPath (Join-Path $PSScriptRoot 'review_manager.py'))) { throw 'Detailed review manager is missing.' }
 $controls.Root.Text='"C:\Example Folder"'
 if($controls.Root.Text -ne 'C:\Example Folder') { throw 'Quoted path normalization failed.' }
 $controls.Root.Text='C:\Invalid|Path';Update-Destination
 $long='Client '+('x'*210)
 if((Get-QueueRootLabel $long).Length -gt 100) { throw 'Queue root label is too long.' }
 $checkFolder=Join-Path ([IO.Path]::GetTempPath()) ('word-pdf-interface-check-'+[Guid]::NewGuid())
 try {
  [void][IO.Directory]::CreateDirectory((Join-Path $checkFolder 'Source'))
  $testReportFolder=Join-Path $checkFolder 'Originals\Source';[void][IO.Directory]::CreateDirectory($testReportFolder)
  $testReport=Join-Path $testReportFolder 'Source [2099-01-01 00-00-00.000] - conversion-report.xlsx'
  [IO.File]::WriteAllBytes($testReport,[byte[]](1,2,3));[IO.File]::SetLastWriteTimeUtc($testReport,[DateTime]::UtcNow.AddDays(1))
  $controls.Root.Text=Join-Path $checkFolder 'Source';$controls.Archive.Text=Join-Path $checkFolder 'Originals'
  Refresh-Reports
  if($controls.Reports.SelectedItem.Tag -ne $testReport -or !$controls.OpenReport.IsEnabled) { throw 'Newest generated report was not selected.' }
  Run-Maintenance 'restore'
  if(!$script:maintenancePlanning -or !$script:running -or $controls.Root.IsEnabled) { throw 'Maintenance planning did not start asynchronously.' }
  $script:running.WaitForExit()
  if($script:running.ExitCode -ne 0) { throw (Read-Shared $script:errLog) }
  Set-MaintenanceBusy $false
 } finally {
  if([IO.Directory]::Exists($checkFolder)) { [IO.Directory]::Delete($checkFolder,$true) }
 }
 'PASS: versioned title, report picker, review dialog, quoted/invalid paths, bounded log names and asynchronous maintenance planning.'
 exit
}
[void]$window.ShowDialog()











