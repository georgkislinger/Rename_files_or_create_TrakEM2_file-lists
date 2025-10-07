<# 
TrakEM2 Helper GUI — 2025-10 Flexible Coordinates
----------------------------------------------------------
Features:
 - Supports ANY combination of coordinates: x, y, z, xy, xz, yz, xyz
 - Missing coordinates auto-filled with 0
 - Full manual template control (e.g., Tile_#z#-#x#___#y#-#r#)
 - Auto-detect tile dimensions from first image
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-Error($msg)  { [System.Windows.MessageBox]::Show($msg,"Error","OK","Error") | Out-Null }
function Show-Info($msg)   { [System.Windows.MessageBox]::Show($msg,"Info","OK","Information") | Out-Null }

# ---------------------------
# XAML GUI Layout
# ---------------------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="TrakEM2 Helper GUI" Height="920" Width="1100" WindowStartupLocation="CenterScreen" Background="#0F172A">
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- FOLDER -->
    <Border Grid.Row="0" Padding="12" CornerRadius="12" Background="#111827" BorderBrush="#374151" BorderThickness="1">
      <StackPanel>
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="Data Folder:" Foreground="#E5E7EB" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <TextBox Name="TxtFolder" Width="700" Background="#1F2937" Foreground="#E5E7EB" BorderBrush="#374151" BorderThickness="1" Padding="6"/>
          <Button Name="BtnBrowse" Content="Browse…" Margin="8,0,0,0" Padding="10,6"/>
          <Button Name="BtnPreview" Content="Preview Filenames" Margin="8,0,0,0" Padding="10,6"/>
        </StackPanel>
        <WrapPanel Margin="0,10,0,0">
          <TextBlock Text="Default Extension (optional):" Foreground="#9CA3AF" Margin="0,0,8,0"/>
          <TextBox Name="TxtDefaultExt" Width="100" Text="tif" Background="#1F2937" Foreground="#E5E7EB" BorderBrush="#374151" BorderThickness="1" Padding="4"/>
        </WrapPanel>
      </StackPanel>
    </Border>

    <!-- OPTIONS -->
    <Border Grid.Row="1" Margin="0,12,0,12" Padding="12" CornerRadius="12" Background="#111827" BorderBrush="#374151" BorderThickness="1">
      <StackPanel>

        <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
          <RadioButton Name="RbRename" Content="Rename files" IsChecked="True" Foreground="#E5E7EB" Margin="0,0,20,0"/>
          <RadioButton Name="RbImport" Content="Create TrakEM2 Import TSV / Copy" Foreground="#E5E7EB"/>
        </StackPanel>

        <!-- RENAME -->
        <StackPanel Name="PanelRename" Visibility="Visible">
          <WrapPanel>
            <TextBlock Text="Level:" Foreground="#E5E7EB" Margin="0,0,8,0"/>
            <ComboBox Name="CmbLevel" Width="200">
              <ComboBoxItem Content="parent" IsSelected="True"/>
              <ComboBoxItem Content="grandparent"/>
              <ComboBoxItem Content="greatgrand"/>
            </ComboBox>
            <TextBlock Text="Extension:" Foreground="#E5E7EB" Margin="16,0,8,0"/>
            <TextBox Name="TxtRenameExt" Width="100" Text="tif"/>
            <CheckBox Name="ChkDryRun" Content="Dry run (no rename)" Foreground="#E5E7EB" Margin="16,0,0,0" IsChecked="True"/>
          </WrapPanel>
          <Button Name="BtnRunRename" Content="Run Rename" Width="200" Margin="0,8,0,0"/>
        </StackPanel>

        <!-- IMPORT -->
        <StackPanel Name="PanelImport" Visibility="Collapsed" Margin="0,12,0,0">
          <UniformGrid Columns="3" Rows="5">
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="TileSizeX:" Foreground="#E5E7EB" Width="120"/>
              <TextBox Name="TxtTileX" Width="100" Text="4000"/>
              <Button Name="BtnDetectSize" Content="Auto" Width="50" Margin="4,0,0,0" ToolTip="Auto-detect from first image"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="TileSizeY:" Foreground="#E5E7EB" Width="120"/>
              <TextBox Name="TxtTileY" Width="100" Text="6000"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="Bit depth:" Foreground="#E5E7EB" Width="120"/>
              <ComboBox Name="CmbBit" Width="100">
                <ComboBoxItem Content="8bit" IsSelected="True"/>
                <ComboBoxItem Content="16bit"/>
              </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="OverlapX:" Foreground="#E5E7EB" Width="120"/>
              <TextBox Name="TxtOvX" Width="100" Text="100"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="OverlapY:" Foreground="#E5E7EB" Width="120"/>
              <TextBox Name="TxtOvY" Width="100" Text="100"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="Z offset:" Foreground="#E5E7EB" Width="120"/>
              <TextBox Name="TxtZStart" Width="100" Text="0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="Extension:" Foreground="#E5E7EB" Width="120"/>
              <TextBox Name="TxtImpExt" Width="100" Text="tif"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="Template:" Foreground="#E5E7EB" Width="120"/>
              <ComboBox Name="CmbTemplate" Width="140">
                <ComboBoxItem Content="Zeiss"/>
                <ComboBoxItem Content="Thermo"/>
                <ComboBoxItem Content="Sequential Numbers" IsSelected="True"/>
                <ComboBoxItem Content="Custom Pattern"/>
              </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,8,8">
              <TextBlock Text="Path format:" Foreground="#E5E7EB" Width="120"/>
              <ComboBox Name="CmbPathFormat" Width="220">
                <ComboBoxItem Content="Relative (.\subdir\file)" IsSelected="True"/>
                <ComboBoxItem Content="File name only"/>
                <ComboBoxItem Content="Full path"/>
              </ComboBox>
            </StackPanel>
            <CheckBox Name="ChkIncludeBase" Content="Include base folder in relative path" Foreground="#E5E7EB" IsChecked="False"/>
          </UniformGrid>

          <!-- Sequential Numbers Panel -->
          <StackPanel Name="PanelSequential" Visibility="Visible" Margin="0,8,0,0">
            <TextBlock Text="Number order (any combo: x/y/z/r). Missing coords auto-fill with 0:" Foreground="#C7D2FE"/>
            <WrapPanel>
              <TextBox Name="TxtNumberOrder" Width="200" Text="z" Background="#1F2937" Foreground="#E5E7EB" Padding="6"/>
              <Button Name="BtnGuessOrder" Content="Guess Order" Width="140" Margin="8,0,0,0"/>
              <TextBlock Text="Examples: 'z', 'xz', 'xy', 'xyz', 'xyzrr' (r=skip)" Foreground="#9CA3AF" Margin="12,4,0,0"/>
            </WrapPanel>
          </StackPanel>

          <!-- Custom Pattern Panel -->
          <StackPanel Name="PanelCustomPattern" Visibility="Collapsed" Margin="0,8,0,0">
            <TextBlock Text="Custom pattern (#x#, #y#, #z# optional, #r#=skip). Missing coords auto-fill with 0:" Foreground="#C7D2FE"/>
            <DockPanel>
              <TextBox Name="TxtPattern" Width="680" Text="#z#" Background="#1F2937" Foreground="#E5E7EB" Padding="6"/>
              <Button Name="BtnGuessPattern" Content="Guess Pattern" Width="140" Margin="8,0,0,0"/>
            </DockPanel>
            <TextBlock Text="Examples: '#z#', 'Tile_#z#-#x#___#y#-#r#', '#x#_#z#'" Foreground="#9CA3AF" Margin="0,4,0,0"/>
          </StackPanel>

          <WrapPanel Name="PanelGuesses" Margin="0,8,0,0" Visibility="Collapsed">
            <TextBlock Text="Guesses:" Foreground="#9CA3AF" Margin="0,0,8,0"/>
            <ComboBox Name="CmbGuesses" Width="720"/>
            <Button Name="BtnUseGuess" Content="Use Selected" Width="120" Margin="8,0,0,0"/>
          </WrapPanel>
          <WrapPanel Margin="0,8,0,0">
            <TextBlock Text="Output file (TSV):" Foreground="#E5E7EB" Margin="0,0,8,0"/>
            <TextBox Name="TxtOutName" Width="240" Text="RawImgList.txt"/>
          </WrapPanel>
          <WrapPanel Margin="0,8,0,0">
            <Button Name="BtnRunImport" Content="Create TSV" Width="220" Margin="0,0,8,0"/>
            <Button Name="BtnCopyOnly" Content="Copy Files" Width="220"/>
          </WrapPanel>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- LOG -->
    <Border Grid.Row="2" Padding="12" CornerRadius="12" Background="#0B1220" BorderBrush="#374151" BorderThickness="1">
      <DockPanel>
        <TextBlock Text="Log / Preview" Foreground="#93C5FD" FontWeight="Bold" DockPanel.Dock="Top" Margin="0,0,0,8"/>
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <TextBox Name="TxtLog" AcceptsReturn="True" TextWrapping="Wrap" Background="#0B1220" Foreground="#E5E7EB" BorderBrush="#1F2937" />
        </ScrollViewer>
      </DockPanel>
    </Border>

    <!-- FOOTER -->
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button Name="BtnOpenOut" Content="Open Output Folder" Padding="10,6" Margin="0,0,8,0"/>
      <Button Name="BtnClose" Content="Close" Padding="10,6"/>
    </StackPanel>
  </Grid>
</Window>
"@

# ---------------------------
# Initialize Window
# ---------------------------
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)
$TxtFolder=$window.FindName('TxtFolder'); $BtnBrowse=$window.FindName('BtnBrowse')
$BtnPreview=$window.FindName('BtnPreview'); $TxtDefaultExt=$window.FindName('TxtDefaultExt')
$RbRename=$window.FindName('RbRename'); $RbImport=$window.FindName('RbImport')
$PanelRename=$window.FindName('PanelRename'); $PanelImport=$window.FindName('PanelImport')
$CmbLevel=$window.FindName('CmbLevel'); $TxtRenameExt=$window.FindName('TxtRenameExt')
$ChkDryRun=$window.FindName('ChkDryRun'); $BtnRunRename=$window.FindName('BtnRunRename')
$TxtTileX=$window.FindName('TxtTileX'); $TxtTileY=$window.FindName('TxtTileY'); $BtnDetectSize=$window.FindName('BtnDetectSize')
$TxtOvX=$window.FindName('TxtOvX'); $TxtOvY=$window.FindName('TxtOvY'); $TxtZStart=$window.FindName('TxtZStart')
$CmbBit=$window.FindName('CmbBit'); $TxtImpExt=$window.FindName('TxtImpExt')
$CmbTemplate=$window.FindName('CmbTemplate')
$TxtNumberOrder=$window.FindName('TxtNumberOrder'); $BtnGuessOrder=$window.FindName('BtnGuessOrder')
$TxtPattern=$window.FindName('TxtPattern'); $BtnGuessPattern=$window.FindName('BtnGuessPattern')
$CmbGuesses=$window.FindName('CmbGuesses'); $BtnUseGuess=$window.FindName('BtnUseGuess')
$PanelSequential=$window.FindName('PanelSequential'); $PanelCustomPattern=$window.FindName('PanelCustomPattern'); $PanelGuesses=$window.FindName('PanelGuesses')
$CmbPathFormat=$window.FindName('CmbPathFormat'); $ChkIncludeBase=$window.FindName('ChkIncludeBase')
$TxtOutName=$window.FindName('TxtOutName'); $BtnRunImport=$window.FindName('BtnRunImport'); $BtnCopyOnly=$window.FindName('BtnCopyOnly')
$TxtLog=$window.FindName('TxtLog'); $BtnOpenOut=$window.FindName('BtnOpenOut'); $BtnClose=$window.FindName('BtnClose')

# ---------------------------
# Utility Functions
# ---------------------------
function Append-Log($msg){$TxtLog.AppendText("$msg`r`n");$TxtLog.ScrollToEnd()}
function Select-FolderDialog(){ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog()-eq"OK"){$dlg.SelectedPath}else{$null} }
function Get-PathMode(){ $sel=($CmbPathFormat.SelectedItem.Content).ToString(); if($sel -like "Relative*"){"relative"}elseif($sel -like "File name*"){"name"}else{"full"} }

function Get-NumberedPath {
    param([string]$basePath)
    if (-not (Test-Path $basePath)) { return $basePath }
    $dir  = Split-Path -Parent $basePath
    $name = Split-Path -Leaf $basePath
    $ext  = [IO.Path]::GetExtension($name)
    $stem = if ($ext) { $name.Substring(0, $name.Length - $ext.Length) } else { $name }
    $i=1
    do {
        $candidate = if ($ext) { Join-Path $dir ("{0}_{1}{2}" -f $stem,$i,$ext) } else { Join-Path $dir ("{0}_{1}" -f $stem,$i) }
        $i++
    } while (Test-Path $candidate)
    return $candidate
}

function Get-ImageDimensions {
    param([string]$imagePath)
    try {
        $image = [System.Drawing.Image]::FromFile($imagePath)
        $dims = @{Width = $image.Width; Height = $image.Height}
        $image.Dispose()
        return $dims
    } catch {
        Append-Log "[WARN] Could not read image dimensions: $($_.Exception.Message)"
        return $null
    }
}

function Update-TemplateVisibility {
    $preset = ($CmbTemplate.SelectedItem.Content).ToString()
    
    if ($preset -eq "Sequential Numbers") {
        $PanelSequential.Visibility = "Visible"
        $PanelCustomPattern.Visibility = "Collapsed"
        $PanelGuesses.Visibility = "Collapsed"
    } elseif ($preset -eq "Custom Pattern") {
        $PanelSequential.Visibility = "Collapsed"
        $PanelCustomPattern.Visibility = "Visible"
        $PanelGuesses.Visibility = "Visible"
    } else {
        $PanelSequential.Visibility = "Collapsed"
        $PanelCustomPattern.Visibility = "Collapsed"
        $PanelGuesses.Visibility = "Collapsed"
    }
}

function Format-PathForOutput([string]$full,[string]$root,[string]$mode,[bool]$IncludeBase){
 switch($mode){
  "name"{return [System.IO.Path]::GetFileName($full)}
  "relative"{
    try{
      $resolvedRoot=(Resolve-Path $root).ProviderPath
      $resolvedFile=(Resolve-Path $full).ProviderPath
      $relative=$resolvedFile.Substring($resolvedRoot.Length)
      if($relative.StartsWith('\') -or $relative.StartsWith('/')){$relative=$relative.Substring(1)}
      if(-not $IncludeBase){
        $baseName=Split-Path $resolvedRoot -Leaf
        if($relative -like "$baseName*"){$relative=$relative.Substring($baseName.Length+1)}
      }
      $relative=$relative -replace '/', '\'
      if(-not [string]::IsNullOrWhiteSpace($relative)){$relative=".\$relative"}else{$relative=".\$([System.IO.Path]::GetFileName($full))"}
      return $relative
    }catch{return [System.IO.Path]::GetFileName($full)}
  }
  default{return (Resolve-Path $full).ProviderPath}
 }
}

# ---------------------------
# FLEXIBLE: Sequential Number Extraction (handles any combination)
# ---------------------------
function Extract-SequentialNumbers {
    param([string]$filename, [string]$order)
    
    # Extract all numbers from filename
    $numbers = [regex]::Matches($filename, '\d+') | ForEach-Object { [int]$_.Value }
    
    if ($numbers.Count -eq 0) { return $null }
    
    # Initialize all coordinates to 0
    $coords = @{x=0; y=0; z=0}
    $orderLower = $order.ToLower()
    $foundAny = $false
    
    # Map order characters to numbers
    for ($i=0; $i -lt $orderLower.Length -and $i -lt $numbers.Count; $i++) {
        $char = $orderLower[$i]
        switch ($char) {
            'x' { $coords.x = $numbers[$i]; $foundAny = $true }
            'y' { $coords.y = $numbers[$i]; $foundAny = $true }
            'z' { $coords.z = $numbers[$i]; $foundAny = $true }
            # 'r' means skip this number
        }
    }
    
    # Return coordinates if we found at least one valid coordinate
    if ($foundAny) {
        return $coords
    }
    
    return $null
}

# ---------------------------
# IMPROVED: Guess number order (tests all combinations)
# ---------------------------
function Guess-NumberOrder {
    param([string]$root, [string]$ext)
    
    if (-not (Test-Path $root)) { return @() }
    $filter = if ($ext) { "*.$ext" } else { "*" }
    $samples = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $filter -ErrorAction SilentlyContinue | Select-Object -First 50
    
    if (-not $samples -or $samples.Count -lt 2) {
        Append-Log "[WARN] Need at least 2 files to guess order"
        return @()
    }
    
    Append-Log "[INFO] Analyzing $($samples.Count) files to guess number order..."
    
    # Check how many numbers per filename
    $numCounts = @{}
    foreach ($f in $samples | Select-Object -First 10) {
        $baseName = [IO.Path]::GetFileNameWithoutExtension($f.Name)
        $numbers = [regex]::Matches($baseName, '\d+')
        $count = $numbers.Count
        if ($numCounts.ContainsKey($count)) {
            $numCounts[$count]++
        } else {
            $numCounts[$count] = 1
        }
    }
    
    $mostCommonCount = ($numCounts.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Key
    Append-Log "[DEBUG] Most files have $mostCommonCount number(s) in filename"
    
    # Test ALL possible combinations
    $patterns = @()
    if ($mostCommonCount -eq 1) {
        $patterns = @('z', 'x', 'y')
    } elseif ($mostCommonCount -eq 2) {
        $patterns = @('xz', 'yz', 'xy', 'zx', 'zy', 'yx')
    } elseif ($mostCommonCount -eq 3) {
        $patterns = @('xyz', 'xzy', 'yxz', 'yzx', 'zxy', 'zyx')
    } else {
        # More than 3 numbers
        $patterns = @('xyz', 'xyzr', 'xyzrr', 'rxyz', 'rxyzt', 'xzr', 'yzr')
    }
    
    $results = @()
    
    foreach ($pattern in $patterns) {
        $validCount = 0
        $xVals = @(); $yVals = @(); $zVals = @()
        
        foreach ($f in $samples) {
            $baseName = [IO.Path]::GetFileNameWithoutExtension($f.Name)
            $coords = Extract-SequentialNumbers -filename $baseName -order $pattern
            if ($coords) {
                $validCount++
                $xVals += $coords.x
                $yVals += $coords.y
                $zVals += $coords.z
            }
        }
        
        if ($validCount -gt 0) {
            $xUnique = ($xVals | Select-Object -Unique).Count
            $yUnique = ($yVals | Select-Object -Unique).Count
            $zUnique = ($zVals | Select-Object -Unique).Count
            
            $coverage = ($validCount / $samples.Count) * 100
            $diversity = $xUnique + $yUnique + $zUnique
            $score = $coverage + ($diversity * 2)
            
            $results += [PSCustomObject]@{
                Pattern = $pattern
                Score = $score
                Matches = $validCount
                Coverage = [math]::Round($coverage, 1)
                XUnique = $xUnique
                YUnique = $yUnique
                ZUnique = $zUnique
            }
        }
    }
    
    $sorted = $results | Sort-Object -Property Score -Descending
    
    if ($sorted.Count -gt 0) {
        Append-Log "[SCORE] Top number order patterns:"
        foreach ($r in ($sorted | Select-Object -First 3)) {
            Append-Log "  '$($r.Pattern)' | Score: $([math]::Round($r.Score,1)) | Matches: $($r.Matches) | Unique X:$($r.XUnique) Y:$($r.YUnique) Z:$($r.ZUnique)"
        }
        return $sorted
    }
    
    return @()
}

# ---------------------------
# FLEXIBLE: Template builder (handles optional coordinates)
# ---------------------------
function Build-RegexFromTemplate {
    param(
        [string]$template,
        [ValidateSet("Zeiss","Thermo","Sequential Numbers","Custom Pattern")] [string]$preset,
        [string]$numberOrder = ""
    )
    
    if ($preset -eq "Zeiss") {
        return @{
            RegTemplate = '^Tile_r(?<y>\d+)-c(?<x>\d+)_S_(?<z>\d+)(?:_(?<r>\d+))?$'
            Template    = 'Tile_r#y#-c#x#_S_#z#_#r#'
            Mode        = 'regex'
        }
    } elseif ($preset -eq "Thermo") {
        return @{
            RegTemplate = '^Tile_(?<y>\d+)-(?<x>\d+)-(?<r>\d+)_\d+-\d+\.s(?<z>\d+)_e\d+$'
            Template    = 'Tile_#y#-#x#-#r#_0-000.s#z#_e00'
            Mode        = 'regex'
        }
    } elseif ($preset -eq "Sequential Numbers") {
        return @{
            NumberOrder = $numberOrder
            Template    = "Sequential: $numberOrder"
            Mode        = 'sequential'
        }
    } else {
        # Custom Pattern - FLEXIBLE: supports any combination of #x#, #y#, #z#
        $temp = $template
        $temp = $temp -replace '#x#', '<!X!>'
        $temp = $temp -replace '#y#', '<!Y!>'
        $temp = $temp -replace '#z#', '<!Z!>'
        $temp = $temp -replace '#r#', '<!R!>'
        
        # Escape literal parts
        $escaped = [Regex]::Escape($temp)
        
        # Replace temp markers with regex groups
        $escaped = $escaped -replace [Regex]::Escape('<!X!>'), '(?<x>\d+)'
        $escaped = $escaped -replace [Regex]::Escape('<!Y!>'), '(?<y>\d+)'
        $escaped = $escaped -replace [Regex]::Escape('<!Z!>'), '(?<z>\d+)'
        $escaped = $escaped -replace [Regex]::Escape('<!R!>'), '\d+'
        
        $RegTemplate = '^' + $escaped + '$'
        return @{
            RegTemplate = $RegTemplate
            Template    = $template
            Mode        = 'regex'
        }
    }
}

# ---------------------------
# Custom Pattern Guesser
# ---------------------------
function Guess-CustomPatterns {
    param([string]$root, [string]$ext)
    
    $candidatePatterns = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    
    if (-not (Test-Path $root)) { return @() }
    $filter = if ($ext) { "*.$ext" } else { "*" }
    $samples = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $filter -ErrorAction SilentlyContinue | Select-Object -First 100
    
    if (-not $samples -or $samples.Count -eq 0) { return @() }
    
    Append-Log "[INFO] Analyzing $($samples.Count) files for custom pattern detection..."
    
    # Check for simple sequential
    $firstFile = [IO.Path]::GetFileNameWithoutExtension($samples[0].Name)
    $numbers = [regex]::Matches($firstFile, '\d+')
    
    if ($numbers.Count -eq 1) {
        [void]$candidatePatterns.Add('#z#')
    }
    
    # Generate patterns
    foreach ($f in $samples) {
        $name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
        if ($name -match '^\d+$') { continue }
        
        $c = $name
        
        # Detect vendor patterns
        if ($name -match 'Tile_r\d+-c\d+') {
            [void]$candidatePatterns.Add('Tile_r#y#-c#x#_S_#z#_#r#')
        }
        if ($name -match 'Tile_\d+-\d+-\d+') {
            [void]$candidatePatterns.Add('Tile_#y#-#x#-#r#_0-000.s#z#_e00')
        }
        
        # Generic detection
        $c = $c -replace '(?i)\bTile_r(\d+)-c(\d+)', 'Tile_r#y#-c#x#'
        $c = $c -replace '(?i)\b(row|r)[-_]?(\d+)\b', 'row#y#'
        $c = $c -replace '(?i)\b(col|column|c)[-_]?(\d+)\b', 'col#x#'
        $c = $c -replace '(?i)\b(slice|sec|section|layer)[-_]?(\d+)\b', 'slice#z#'
        $c = $c -replace '(?i)\bx[-_]?(\d+)\b', 'x#x#'
        $c = $c -replace '(?i)\by[-_]?(\d+)\b', 'y#y#'
        $c = $c -replace '(?i)\bz[-_]?(\d+)\b', 'z#z#'
        $c = $c -replace '(?i)_[sS](\d+)_', '_S#z#_'
        $c = [Regex]::Replace($c, '(?<!#)(\d+)(?!#)', '#r#')
        $c = $c -replace '(#r#[-_.]?)+', '#r#'
        
        if ($c -match '#x#|#y#|#z#') {
            [void]$candidatePatterns.Add($c)
        }
    }
    
    if ($candidatePatterns.Count -eq 0) {
        return @()
    }
    
    # Score patterns
    $scoredPatterns = @()
    foreach ($pattern in $candidatePatterns) {
        $tpl = Build-RegexFromTemplate -template $pattern -preset "Custom Pattern"
        if (-not $tpl.RegTemplate) { continue }
        
        try {
            $reg = [Regex]::new($tpl.RegTemplate, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $matchCount = 0
            
            foreach ($f in $samples) {
                $baseName = [IO.Path]::GetFileNameWithoutExtension($f.Name)
                $m = $reg.Match($baseName)
                # Accept if ANY coordinate is found
                if ($m.Success -and ($m.Groups['x'].Success -or $m.Groups['y'].Success -or $m.Groups['z'].Success)) {
                    $matchCount++
                }
            }
            
            if ($matchCount -gt 0) {
                $coverage = ($matchCount / $samples.Count) * 100
                $scoredPatterns += [PSCustomObject]@{
                    Pattern = $pattern
                    Score = $coverage
                    Matches = $matchCount
                }
            }
        } catch {
            continue
        }
    }
    
    $sorted = $scoredPatterns | Sort-Object -Property Score -Descending
    
    if ($sorted.Count -gt 0) {
        Append-Log "[SCORE] Top custom patterns:"
        foreach ($sp in ($sorted | Select-Object -First 3)) {
            Append-Log "  Score: $([math]::Round($sp.Score,1))% | Matches: $($sp.Matches) | Pattern: $($sp.Pattern)"
        }
    }
    
    return @($sorted | Select-Object -ExpandProperty Pattern)
}

# ---------------------------
# RENAME MODE
# ---------------------------
function Run-Rename {
    param([string]$root,[ValidateSet("parent","grandparent","greatgrand")] [string]$level,[string]$ext,[switch]$DryRun)
    if (-not (Test-Path $root)) { Show-Error "Folder not found: $root"; return }
    if (-not $ext) { Show-Error "Please enter a file extension (e.g., tif)"; return }
    $pattern = "*.$ext"
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pattern -ErrorAction SilentlyContinue
    if (-not $files) { Append-Log "[WARN] No *.$ext files found under $root"; return }
    Append-Log "[INFO] Found $($files.Count) file(s) with extension .$ext"
    $n=0; $skipped=0
    foreach ($f in $files) {
        try {
            $prefix = switch ($level) {
                "parent"      { $f.Directory.Name }
                "grandparent" { $f.Directory.Parent.Name }
                "greatgrand"  { $f.Directory.Parent.Parent.Name }
            }
            if (-not $prefix) { $skipped++; continue }
            $newName = "$prefix-$($f.Name)"
            if ($DryRun) { Append-Log "[DRY] $($f.FullName) -> $newName" }
            else { Rename-Item -LiteralPath $f.FullName -NewName $newName -ErrorAction Stop; Append-Log "$($f.Name) -> $newName" }
            $n++
        } catch { Append-Log "[ERR] $($f.FullName) : $($_.Exception.Message)"; $skipped++ }
    }
    Append-Log "[DONE] Rename complete. Renamed: $n, Skipped: $skipped"
}

# ---------------------------
# FLEXIBLE MATCH COLLECTOR (handles partial coordinates)
# ---------------------------
$script:LastMatchContext = $null
function Get-Matches {
    param(
        [string]$root,
        [string]$ext,
        [ValidateSet("Zeiss","Thermo","Sequential Numbers","Custom Pattern")] [string]$preset,
        [string]$template,
        [string]$numberOrder = ""
    )
    
    $tpl = Build-RegexFromTemplate -template $template -preset $preset -numberOrder $numberOrder
    
    Append-Log "[INFO] Template preset: $preset"
    if ($tpl.Mode -eq 'sequential') {
        Append-Log "[DEBUG] Using sequential extraction: $($tpl.NumberOrder)"
    } else {
        Append-Log "[DEBUG] Using regex: $($tpl.RegTemplate)"
    }
    
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.$ext" -ErrorAction SilentlyContinue
    if (-not $files) { Append-Log "[WARN] No *.$ext files found."; return $null }
    Append-Log "[INFO] Found $($files.Count) candidate file(s). Matching template..."
    
    $matched = New-Object System.Collections.Generic.List[object]
    $X = New-Object System.Collections.Generic.List[int]
    $Y = New-Object System.Collections.Generic.List[int]
    $Z = New-Object System.Collections.Generic.List[int]
    $FullNames = New-Object System.Collections.Generic.List[string]
    
    if ($tpl.Mode -eq 'sequential') {
        # Sequential: coordinates default to 0 if not in pattern
        foreach ($f in $files) {
            $baseName = [IO.Path]::GetFileNameWithoutExtension($f.Name)
            $coords = Extract-SequentialNumbers -filename $baseName -order $tpl.NumberOrder
            if ($coords) {
                $matched.Add($f)
                $X.Add($coords.x)
                $Y.Add($coords.y)
                $Z.Add($coords.z)
                $FullNames.Add($f.FullName)
            }
        }
    } else {
        # Regex: accept if ANY coordinate matches, fill missing with 0
        $reg = [Regex]::new($tpl.RegTemplate, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($f in $files) {
            $baseName = [IO.Path]::GetFileNameWithoutExtension($f.Name)
            $m = $reg.Match($baseName)
            
            # Accept if we have at least ONE coordinate
            if ($m.Success -and ($m.Groups['x'].Success -or $m.Groups['y'].Success -or $m.Groups['z'].Success)) {
                $xVal = if ($m.Groups['x'].Success) { [int]$m.Groups['x'].Value } else { 0 }
                $yVal = if ($m.Groups['y'].Success) { [int]$m.Groups['y'].Value } else { 0 }
                $zVal = if ($m.Groups['z'].Success) { [int]$m.Groups['z'].Value } else { 0 }
                
                $matched.Add($f)
                $X.Add($xVal)
                $Y.Add($yVal)
                $Z.Add($zVal)
                $FullNames.Add($f.FullName)
            }
        }
    }
    
    Append-Log "[INFO] Matched $($matched.Count) file(s)."
    if ($matched.Count -eq 0) { return $null }
    
    # Sort by Z, then Y, then X
    $idx = 0..($matched.Count-1)
    $sortedIdx = $idx | Sort-Object { [int]$Z[$_] }, { [int]$Y[$_] }, { [int]$X[$_] }
    
    $ctx = [PSCustomObject]@{
        Files      = $matched[$sortedIdx]
        X          = $X[$sortedIdx]
        Y          = $Y[$sortedIdx]
        Z          = $Z[$sortedIdx]
        FullNames  = $FullNames[$sortedIdx]
        Preset     = $preset
        Template   = $tpl.Template
        Ext        = $ext
    }
    $script:LastMatchContext = $ctx
    return $ctx
}

# ---------------------------
# IMPORT MODE (TSV)
# ---------------------------
function Run-Import {
    param(
        [string]$root,[int]$TileX,[int]$TileY,[int]$OvX,[int]$OvY,[int]$ZStart,
        [ValidateSet("8bit","16bit")] [string]$BitDepth,
        [string]$ext,[ValidateSet("Zeiss","Thermo","Sequential Numbers","Custom Pattern")] [string]$preset,
        [string]$template,[string]$numberOrder,
        [string]$pathMode,[bool]$IncludeBase,[string]$OutName
    )
    if (-not (Test-Path $root)) { Show-Error "Folder not found: $root"; return $null }
    if (-not $ext) { Show-Error "Please enter a file extension (e.g., tif)"; return $null }
    if (-not $OutName) { $OutName = "RawImgList.txt" }
    
    $ctx = Get-Matches -root $root -ext $ext -preset $preset -template $template -numberOrder $numberOrder
    if (-not $ctx) { Append-Log "[WARN] Nothing matched. Aborting."; return $null }
    
    $imgType = if ($BitDepth -eq "16bit") { 1 } else { 0 }
    $minInt  = 0
    $maxInt  = if ($imgType -eq 1) { 65535 } else { 255 }
    
    $xar = for ($i=0; $i -lt $ctx.Files.Count; $i++) { $ctx.X[$i] * ($TileX - $OvX) }
    $yar = for ($i=0; $i -lt $ctx.Files.Count; $i++) { $ctx.Y[$i] * ($TileY - $OvY) }
    $zar = for ($i=0; $i -lt $ctx.Files.Count; $i++) { $ctx.Z[$i] + $ZStart }
    
    $parent = Split-Path -Parent (Resolve-Path $root)
    if (-not $parent) { $parent = $root }
    $outPath = Join-Path -Path $parent -ChildPath $OutName
    if (Test-Path $outPath) { $outPath = Get-NumberedPath -basePath $outPath }
    Append-Log "[IO] Writing TSV: $outPath"
    
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i=0; $i -lt $ctx.Files.Count; $i++) {
        $fileOut = Format-PathForOutput -full $ctx.FullNames[$i] -root $root -mode $pathMode -IncludeBase:$IncludeBase
        $rows.Add([PSCustomObject]@{
            FileName     = $fileOut
            XPos         = $xar[$i]
            YPos         = $yar[$i]
            Section      = $zar[$i]
            TileSizeX    = $TileX
            TileSizeY    = $TileY
            MinIntensity = $minInt
            MaxIntensity = $maxInt
            ImageType    = $imgType
        })
    }
    
    $rows | Export-Csv -Delimiter "`t" -Path $outPath -NoTypeInformation -Encoding UTF8
    (Get-Content -Path $outPath -Raw) -replace '"','' | Set-Content -Path $outPath -Encoding UTF8
    Append-Log "[DONE] TSV written: $outPath"
    return $outPath
}

function Run-CopyOnly {
    param([string]$root)
    if (-not $script:LastMatchContext) { Append-Log "[WARN] No match context found. Click 'Create TSV' first."; return $null }
    $ctx = $script:LastMatchContext
    $parent = Split-Path -Parent (Resolve-Path $root)
    if (-not $parent) { $parent = $root }
    
    $destBase = Join-Path $parent "copied_rawdata"
    $dest = if (Test-Path $destBase) { Get-NumberedPath -basePath $destBase } else { $destBase }
    New-Item -ItemType Directory -Path $dest | Out-Null
    
    Append-Log "[IO] Copying $($ctx.FullNames.Count) matched file(s) to: $dest"
    $copied=0; $errors=0
    foreach ($p in $ctx.FullNames) {
        try { Copy-Item -LiteralPath $p -Destination $dest -ErrorAction Stop; $copied++ }
        catch { $errors++; Append-Log "[ERR] Copy: $p : $($_.Exception.Message)" }
    }
    Append-Log "[DONE] Copy complete. Copied: $copied, Errors: $errors"
    return $dest
}

function Run-Preview {
    param([string]$root,[string]$ext = "",[int]$limit = 25,[bool]$IncludeBase = $false)
    if (-not (Test-Path $root)) { Append-Log "[ERR] Folder not found: $root"; return }
    $filter = if ($ext) { "*.$ext" } else { "*" }
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $filter -ErrorAction SilentlyContinue
    if (-not $files) { Append-Log "[WARN] No files found for preview (filter: $filter)"; return }
    $TxtLog.Clear()
    Append-Log "📁 Data folder: $root"
    Append-Log "Found $($files.Count) file(s) (filter: $filter). Showing up to $limit samples:`r`n"
    $mode = Get-PathMode
    $count = 0
    foreach ($f in $files) {
        if ($count -ge $limit) { break }
        $rel = Format-PathForOutput -full $f.FullName -root $root -mode $mode -IncludeBase:$IncludeBase
        Append-Log "  $rel"
        $count++
    }
    if ($files.Count -gt $limit) { Append-Log "`r`n... ($($files.Count - $limit) more)" }
}

# ---------------------------
# Event wiring
# ---------------------------
$BtnBrowse.Add_Click({ $sel=Select-FolderDialog; if($sel){$TxtFolder.Text=$sel} })
$RbRename.Add_Checked({$PanelRename.Visibility="Visible";$PanelImport.Visibility="Collapsed"})
$RbImport.Add_Checked({$PanelRename.Visibility="Collapsed";$PanelImport.Visibility="Visible"})

$CmbTemplate.Add_SelectionChanged({ Update-TemplateVisibility })

$BtnDetectSize.Add_Click({
    $root = $TxtFolder.Text
    if (-not (Test-Path $root)) { Show-Error "Please select a valid folder first."; return }
    $ext = $TxtImpExt.Text.Trim('.').Trim()
    if (-not $ext) { $ext = $TxtDefaultExt.Text.Trim('.').Trim() }
    $filter = if ($ext) { "*.$ext" } else { "*" }
    $firstImage = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $filter -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($firstImage) {
        $dims = Get-ImageDimensions -imagePath $firstImage.FullName
        if ($dims) {
            $TxtTileX.Text = $dims.Width
            $TxtTileY.Text = $dims.Height
            Append-Log "[AUTO] Detected dimensions from $($firstImage.Name): $($dims.Width) x $($dims.Height)"
            Show-Info "Detected: $($dims.Width) x $($dims.Height) from $($firstImage.Name)"
        }
    } else {
        Show-Error "No image files found in folder."
    }
})

$BtnGuessOrder.Add_Click({
    $root = $TxtFolder.Text
    if (-not (Test-Path $root)) { Show-Error "Please select a valid data folder."; return }
    $ext = $TxtImpExt.Text.Trim('.').Trim()
    $guesses = @(Guess-NumberOrder -root $root -ext $ext)
    
    if ($guesses.Count -eq 0) {
        Append-Log "[WARN] Could not guess number order. Try manual entry (e.g., 'z', 'xz', 'xyz')."
        return
    }
    
    $TxtNumberOrder.Text = $guesses[0].Pattern
    Append-Log "[INFO] Applied best guess: '$($guesses[0].Pattern)'"
})

$BtnGuessPattern.Add_Click({
    $root = $TxtFolder.Text
    if (-not (Test-Path $root)) { Show-Error "Please select a valid data folder."; return }
    $ext = $TxtImpExt.Text.Trim('.').Trim()
    $guesses = @(Guess-CustomPatterns -root $root -ext $ext)
    $CmbGuesses.Items.Clear()
    
    if ($guesses.Count -eq 0) {
        Append-Log "[INFO] No obvious patterns. Use manual entry like '#z#' or 'Tile_#x#_#z#'."
        return
    }
    
    foreach ($g in $guesses) {
        if (-not [string]::IsNullOrWhiteSpace($g)) { [void]$CmbGuesses.Items.Add($g) }
    }
    if ($CmbGuesses.Items.Count -gt 0) {
        $CmbGuesses.SelectedIndex = 0
        $TxtPattern.Text = $CmbGuesses.SelectedItem
        Append-Log "[INFO] Found $($guesses.Count) pattern(s). Top pattern applied."
    }
})

$BtnUseGuess.Add_Click({
    if($CmbGuesses.SelectedItem){
        $TxtPattern.Text=$CmbGuesses.SelectedItem
        Append-Log "Using custom pattern: $($TxtPattern.Text)"
    }
})

$ChkIncludeBase.Add_Click({
 if($TxtFolder.Text -and (Test-Path $TxtFolder.Text)){
   Append-Log "[INFO] Refreshing preview..."
   Run-Preview -root $TxtFolder.Text -ext ($TxtDefaultExt.Text.Trim('.').Trim()) -limit 10 -IncludeBase:([bool]$ChkIncludeBase.IsChecked)
 }
})

$BtnPreview.Add_Click({
 if(-not(Test-Path $TxtFolder.Text)){Show-Error "Please select a valid folder.";return}
 Run-Preview -root $TxtFolder.Text -ext ($TxtDefaultExt.Text.Trim('.').Trim()) -limit 25 -IncludeBase:([bool]$ChkIncludeBase.IsChecked)
})

$BtnRunRename.Add_Click({
    $root = $TxtFolder.Text
    if (-not (Test-Path $root)) { Show-Error "Please select a valid data folder."; return }
    $level = ($CmbLevel.SelectedItem.Content).ToString()
    $ext   = $TxtRenameExt.Text.Trim('.').Trim()
    $dry   = [bool]$ChkDryRun.IsChecked
    Append-Log "[RUN] Rename in: $root  | level=$level | ext=$ext | dry=$dry"
    Run-Rename -root $root -level $level -ext $ext -DryRun:([bool]$dry)
    Show-Info "Rename finished. See log for details."
})

$BtnRunImport.Add_Click({
    $root = $TxtFolder.Text
    if (-not (Test-Path $root)) { Show-Error "Please select a valid data folder."; return }
    try {
        $tileX = [int]$TxtTileX.Text; $tileY = [int]$TxtTileY.Text
        $ovx   = [int]$TxtOvX.Text;  $ovy   = [int]$TxtOvY.Text; $zst = [int]$TxtZStart.Text
    } catch { Show-Error "Tile sizes, overlaps and Z offset must be integers."; return }
    
    $bit   = ($CmbBit.SelectedItem.Content).ToString()
    $ext   = $TxtImpExt.Text.Trim('.').Trim()
    $preset= ($CmbTemplate.SelectedItem.Content).ToString()
    
    $template = ""
    $numberOrder = ""
    
    if ($preset -eq "Sequential Numbers") {
        $numberOrder = $TxtNumberOrder.Text.Trim()
        if (-not $numberOrder) { Show-Error "Please enter a number order (e.g., 'z', 'xz', 'xyz')"; return }
    } elseif ($preset -eq "Custom Pattern") {
        $template = $TxtPattern.Text
        if (-not $template) { Show-Error "Please enter a custom pattern (e.g., '#z#' or 'Tile_#x#_#z#')"; return }
    }
    
    $pathMode = Get-PathMode
    $includeBase = [bool]$ChkIncludeBase.IsChecked
    $outn  = $TxtOutName.Text
    
    Append-Log "[RUN] Create TSV"
    $out = Run-Import -root $root -TileX $tileX -TileY $tileY -OvX $ovx -OvY $ovy -ZStart $zst -BitDepth $bit `
                      -ext $ext -preset $preset -template $template -numberOrder $numberOrder `
                      -pathMode $pathMode -IncludeBase:$includeBase -OutName $outn
    if ($out) { Append-Log "[OK] Output file: $out" ; Show-Info "Finished. Output: $out" }
})

$BtnCopyOnly.Add_Click({
    $root = $TxtFolder.Text
    if (-not (Test-Path $root)) { Show-Error "Please select a valid data folder."; return }
    if (-not $script:LastMatchContext) {
        $ext   = $TxtImpExt.Text.Trim('.').Trim()
        $preset= ($CmbTemplate.SelectedItem.Content).ToString()
        
        $template = ""
        $numberOrder = ""
        
        if ($preset -eq "Sequential Numbers") {
            $numberOrder = $TxtNumberOrder.Text.Trim()
        } elseif ($preset -eq "Custom Pattern") {
            $template = $TxtPattern.Text
        }
        
        $null = Get-Matches -root $root -ext $ext -preset $preset -template $template -numberOrder $numberOrder
        if (-not $script:LastMatchContext) { Append-Log "[WARN] Nothing matched to copy."; return }
    }
    $dest = Run-CopyOnly -root $root
    if ($dest) { Append-Log "[OK] Copied to: $dest"; Show-Info "Copy finished: $dest" }
})

$BtnOpenOut.Add_Click({
    $path = $TxtFolder.Text
    if (Test-Path $path) {
        $parent = Split-Path -Parent (Resolve-Path $path)
        if (-not $parent) { $parent = $path }
        Start-Process explorer.exe $parent
    } else { Show-Error "Select a valid folder first." }
})

$BtnClose.Add_Click({$window.Close()})

Update-TemplateVisibility

[void]$window.ShowDialog()
