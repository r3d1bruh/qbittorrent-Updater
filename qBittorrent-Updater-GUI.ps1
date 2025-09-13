# qBittorrent-Updater-GUI.ps1  (Win11 WPF, DARK, standard build, modern picker only, fixed bullets + const)

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  powershell -STA -ExecutionPolicy Bypass -File "$PSCommandPath"; exit
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Drawing

# --- DWM (Mica + dark title) ---
$src = @"
using System;
using System.Runtime.InteropServices;
public static class DwmApi {
  [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);
}
"@
Add-Type -TypeDefinition $src -ErrorAction SilentlyContinue
function Enable-Mica($hwnd){
  try{
    int $v=1; [void][DwmApi]::DwmSetWindowAttribute($hwnd,20,[ref]$v,4); [void][DwmApi]::DwmSetWindowAttribute($hwnd,19,[ref]$v,4)
    int $mica=2; [void][DwmApi]::DwmSetWindowAttribute($hwnd,38,[ref]$mica,4)
  }catch{}
}

# --- Modern folder picker (IFileDialog) ---
$picker = @"
using System;
using System.Runtime.InteropServices;
namespace FolderDialog {
  [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")] internal class FileOpenDialog {}
  [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IFileDialog { int Show(IntPtr p); void SetFileTypes(uint a, IntPtr b); void SetFileTypeIndex(uint i); void GetFileTypeIndex(out uint i);
    void Advise(IntPtr a, out uint b); void Unadvise(uint a); void SetOptions(uint o); void GetOptions(out uint o); void SetDefaultFolder(IntPtr a);
    void SetFolder(IntPtr a); void GetFolder(out IntPtr a); void GetCurrentSelection(out IntPtr a); void SetFileName(string s); void GetFileName(out IntPtr s);
    void SetTitle(string s); void SetOkButtonLabel(string s); void SetFileNameLabel(string s); void GetResult(out IShellItem i); void AddPlace(IntPtr a,int b);
    void SetDefaultExtension(string s); void Close(int hr); void SetClientGuid(ref Guid g); void ClearClientData(); void SetFilter(IntPtr p); }
  [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IShellItem { void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv); void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName, out IntPtr ppszName); void GetAttributes(uint m, out uint a); void Compare(IShellItem psi, uint h, out int o); }
  public static class FolderPicker {
    const uint FOS_PICKFOLDERS=0x20, FOS_FORCEFILESYSTEM=0x40, FOS_PATHMUSTEXIST=0x800, FOS_NO_VALIDATE=0x100, FOS_DONTADDTORECENT=0x02000000, SIGDN_FILESYSPATH=0x80058000;
    [DllImport("ole32.dll")] static extern void CoTaskMemFree(IntPtr p);
    public static string PickFolder(IntPtr owner){
      var dlg=(IFileDialog)new FileOpenDialog(); uint opts; dlg.GetOptions(out opts);
      dlg.SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_NO_VALIDATE | FOS_DONTADDTORECENT);
      if (dlg.Show(owner)!=0) return null; // cancel -> null
      IShellItem item; dlg.GetResult(out item);
      IntPtr psz; item.GetDisplayName(SIGDN_FILESYSPATH, out psz);
      string path=Marshal.PtrToStringUni(psz); CoTaskMemFree(psz); return path;
    }
  }
}
"@
Add-Type -TypeDefinition $picker

# --- Logging ---
$LogDir="$env:ProgramData\qBittorrentUpdater"; $Log=Join-Path $LogDir "run.log"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
function Log($m){ "[$(Get-Date -Format s)] $m" | Out-File $Log -Append -Encoding utf8 }

# --- Engine (standard build only) ---
function Normalize-Version([string]$v){ if([string]::IsNullOrWhiteSpace($v)){return $null}; ($v -replace '^[vV]','').Trim() }
function Get-InstalledVersion([string]$dir){
  $exe=Join-Path $dir "qbittorrent.exe"
  if(Test-Path $exe){ return (Get-Item $exe).VersionInfo.ProductVersion }
  return $null
}
function Get-LatestStableVersion{
  $h=Invoke-WebRequest -UseBasicParsing -Uri "https://www.qbittorrent.org/"
  $m=[regex]::Match($h.Content,"Latest:\s*v?(?<ver>\d+(?:\.\d+)+)"); if($m.Success){$m.Groups['ver'].Value}else{throw "Could not detect latest version"}
}
function Get-DownloadUrl([string]$ver){
  "https://sourceforge.net/projects/qbittorrent/files/qbittorrent-win32/qbittorrent-$ver/qbittorrent_${ver}_x64_setup.exe/download"
}
function Resolve-Redirect([string]$u){
  $f=$u; for($i=0;$i -lt 6;$i++){ $r=Invoke-WebRequest -Uri $f -MaximumRedirection 0 -UseBasicParsing -ea SilentlyContinue
    if($r.StatusCode -ge 300 -and $r.Headers.Location){ $f=$r.Headers.Location } else { break } } $f
}
function Ensure-NotRunning{ $p=Get-Process -Name "qbittorrent" -ea SilentlyContinue; if($p){$p|Stop-Process -Force} }
function Run-Installer([string]$u,[string]$dir){
  $tmp=Join-Path $env:TEMP ("qbittorrent_"+[Guid]::NewGuid().ToString("N")+".exe")
  try{
    $final=Resolve-Redirect $u; Log "FinalURL=$final"
    Invoke-WebRequest -Uri $final -OutFile $tmp -UseBasicParsing -Headers @{ 'User-Agent'='Wget' }
    if((Get-Item $tmp).Length -lt 5MB){ throw "Download too small; likely HTML" }
    $args="/S /D=" + ($dir.TrimEnd('\'))
    $psi=New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName=$tmp; $psi.Arguments=$args; $psi.UseShellExecute=$true; $psi.Verb="runas"
    $p=[System.Diagnostics.Process]::Start($psi); $p.WaitForExit(); return $p.ExitCode
  } finally { Remove-Item $tmp -Force -ea SilentlyContinue }
}
function Ensure-ScheduledTask([string]$name,[string]$script,[string]$dir,[datetime]$at,[bool]$startup){
  $exe=(Get-Command powershell.exe).Source
  $args=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"'-f $script),'-InstallDir',('"{0}"'-f $dir)) -join ' '
  $action=New-ScheduledTaskAction -Execute $exe -Argument $args
  $trigs=@(New-ScheduledTaskTrigger -Daily -At $at); if($startup){$trigs+=New-ScheduledTaskTrigger -AtStartup}
  $set=New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RunOnlyIfNetworkAvailable -MultipleInstances IgnoreNew
  $pr=New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
  $ex=Get-ScheduledTask -TaskName $name -ea SilentlyContinue
  if($ex){ Set-ScheduledTask -TaskName $name -Action $action -Trigger $trigs -Settings $set -Principal $pr | Out-Null }
  else { Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigs -Settings $set -Principal $pr -Description "qBittorrent auto-update" | Out-Null }
}

# --- XAML (dark styles + bullet fix) ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="qBittorrent Updater" Height="540" Width="860" WindowStartupLocation="CenterScreen"
        Background="#141414" Foreground="#F2F2F2" FontFamily="Segoe UI Variable" FontSize="13">
  <Window.Resources>
    <SolidColorBrush x:Key="CardBg" Color="#1E1E1E"/>
    <SolidColorBrush x:Key="CardBorder" Color="#2C2C2C"/>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="16"/>
      <Setter Property="Margin" Value="12"/>
      <Setter Property="Background" Value="{StaticResource CardBg}"/>
      <Setter Property="BorderBrush" Value="{StaticResource CardBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Margin" Value="0,6,0,6"/>
      <Setter Property="Padding" Value="8"/>
      <Setter Property="Background" Value="#232323"/>
      <Setter Property="BorderBrush" Value="#3A3A3A"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Foreground" Value="#F2F2F2"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,6,0,6"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Margin" Value="0,6,8,6"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="Background" Value="#3D6AE6"/>
      <Setter Property="BorderBrush" Value="#3D6AE6"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border CornerRadius="10"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="6,2,6,2"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#5B84F1"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Background" Value="#3A3A3A"/>
          <Setter Property="BorderBrush" Value="#3A3A3A"/>
          <Setter Property="Foreground" Value="#9E9E9E"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Grid Margin="8">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Style="{StaticResource Card}">
      <StackPanel>
        <TextBlock Text="qBittorrent Updater (standard build)" 
                   FontSize="20" FontWeight="SemiBold" Margin="0,0,0,6"/>
        <TextBlock Opacity="0.85">
          <Run Text="Windows 11 dark UI"/><Run Text=" &#x2022; "/>
          <Run Text="Mica backdrop"/><Run Text=" &#x2022; "/>
          <Run Text="Daily auto-update + manual run"/>
        </TextBlock>
      </StackPanel>
    </Border>

    <!-- Body -->
    <Border Grid.Row="1" Style="{StaticResource Card}">
      <StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
          <TextBlock Text="Install directory:" Width="140" VerticalAlignment="Center"/>
          <TextBox x:Name="TxtDir" Text="E:\!Piracy\qBittorrent" Width="500" Margin="0,0,8,0"/>
          <Button x:Name="BtnBrowse" Content="Browse..." Width="90"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
          <TextBlock Text="Daily time (HH:mm):" Width="140" VerticalAlignment="Center"/>
          <TextBox x:Name="TxtTime" Text="03:30" Width="80"/>
          <CheckBox x:Name="ChkStartup" Content="Also run at startup" Margin="20,0,0,0" IsChecked="True"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
          <Button x:Name="BtnUpdate" Content="Update now"/>
          <Button x:Name="BtnSchedule" Content="Create or update daily task" Margin="8,0,0,0"/>
          <Button x:Name="BtnOpenLog" Content="Open log" Margin="8,0,0,0"/>
          <Button x:Name="BtnExit" Content="Exit" Margin="8,0,0,0"/>
        </StackPanel>
        <TextBox x:Name="TxtOut" Height="220" Background="#161616" BorderBrush="#2C2C2C"
                 Foreground="#F2F2F2" IsReadOnly="True" TextWrapping="Wrap"
                 VerticalScrollBarVisibility="Auto" AcceptsReturn="True"/>
      </StackPanel>
    </Border>

    <!-- Footer -->
    <Border Grid.Row="2" Style="{StaticResource Card}">
      <TextBlock Text="Tip: run once as admin to set the task. Logs: C:\ProgramData\qBittorrentUpdater\run.log" Opacity="0.85"/>
    </Border>
  </Grid>
</Window>
"@

# --- Load UI ---
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Controls
$TxtDir=$window.FindName('TxtDir'); $BtnBrowse=$window.FindName('BtnBrowse')
$TxtTime=$window.FindName('TxtTime'); $ChkStartup=$window.FindName('ChkStartup')
$BtnUpdate=$window.FindName('BtnUpdate'); $BtnSchedule=$window.FindName('BtnSchedule')
$BtnOpenLog=$window.FindName('BtnOpenLog'); $BtnExit=$window.FindName('BtnExit')
$TxtOut=$window.FindName('TxtOut')

$window.Add_SourceInitialized({
  $h=[System.Windows.Interop.WindowInteropHelper]::new($window).Handle
  Enable-Mica $h
})

function AppendOut($s){ $TxtOut.AppendText($s + [Environment]::NewLine); $TxtOut.ScrollToEnd() }

# Modern picker only; cancel = silent
$BtnBrowse.Add_Click({
  try {
    $h=[System.Windows.Interop.WindowInteropHelper]::new($window).Handle
    $picked = [FolderDialog.FolderPicker]::PickFolder($h)
    if ($picked) { $TxtDir.Text = $picked }
  } catch { return }
})

$BtnOpenLog.Add_Click({ if(Test-Path $Log){ Start-Process notepad.exe $Log } })
$BtnExit.Add_Click({ $window.Close() })

$BtnSchedule.Add_Click({
  try{
    $inst=$TxtDir.Text
    $time=[DateTime]::ParseExact($TxtTime.Text,'HH:mm',$null)
    AppendOut "Scheduling..."
    Ensure-ScheduledTask -name "qBittorrent Auto Update" -script $PSCommandPath -dir $inst -at $time -startup:$($ChkStartup.IsChecked)
    AppendOut "Scheduled daily at $($time.ToString('HH:mm'))  (startup=$($ChkStartup.IsChecked))"
  } catch { [System.Windows.MessageBox]::Show($_.ToString(),"Error") }
})

$BtnUpdate.Add_Click({
  $BtnUpdate.IsEnabled=$false
  try{
    $inst=$TxtDir.Text
    AppendOut "Checking versions..."
    $installedRaw=Get-InstalledVersion $inst; $latestRaw=Get-LatestStableVersion
    $installed=Normalize-Version $installedRaw; $latest=Normalize-Version $latestRaw
    AppendOut "Installed: $installedRaw"
    AppendOut "Latest:    $latest"
    if($installed -and ([version]$latest -le [version]$installed)){ AppendOut "Already up to date."; return }
    AppendOut "Stopping qBittorrent and updating..."; Ensure-NotRunning
    $dl=Get-DownloadUrl -ver $latest; AppendOut "Downloading: $dl"
    $code=Run-Installer -u $dl -dir $inst
    Start-Sleep 2
    $newRaw=Get-InstalledVersion $inst
    $new=Normalize-Version $newRaw
    if($new -and ([version]$new -eq [version]$latest)){
      AppendOut "Update complete. Installed: $new"
    } else {
      AppendOut "Update failed. ExitCode=$code InstalledAfter='$newRaw'"
    }
  } finally { $BtnUpdate.IsEnabled=$true }
})

[void]$window.ShowDialog()
