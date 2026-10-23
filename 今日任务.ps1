$ErrorActionPreference = 'Stop'
$script:instanceMutex = New-Object Threading.Mutex($false,'Local\CodexTodayTasksV2')
if (-not $script:instanceMutex.WaitOne(0)) { exit }
$log = Join-Path $PSScriptRoot 'startup.log'
Set-Content -LiteralPath $log -Value ('START ' + (Get-Date))
try {
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
Add-Content -LiteralPath $log -Value 'FORMS_LOADED'
$dataPath = Join-Path $PSScriptRoot 'tasks.json'
$script:items = @()
if (Test-Path -LiteralPath $dataPath) {
 try { $script:items = @(Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.text) }) } catch { [Windows.Forms.MessageBox]::Show('任务数据无法读取，请保留 tasks.json 文件后检查。'); exit }
}
$script:items = @(foreach ($batch in $script:items) { foreach ($task in $batch) {
 [pscustomobject]@{
  text = [string]$task.text
  done = [bool]$task.done
  deadline = $task.deadline
  priority = $(if ($null -eq $task.priority) { 1 } else { [int]$task.priority })
 }
} })
function Get-DeadlineTime([string]$value) {
 if ($value.Length -eq 10) {
  return [datetime]::ParseExact($value,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddDays(1).AddTicks(-1)
 }
 return [datetime]::ParseExact($value,'yyyy-MM-dd HH:mm',[Globalization.CultureInfo]::InvariantCulture)
}
function Get-TaskColor($task) {
 if ($task.done) { return [Drawing.Color]::FromArgb(228,234,231) }
 $urgency = 0.0
 if ($task.deadline) {
  $days = ((Get-DeadlineTime $task.deadline) - (Get-Date)).TotalDays
  $urgency = [Math]::Max(0.0,[Math]::Min(1.0,(14.0-$days)/14.0))
 }
 $score = [Math]::Min(1.0,0.45 * ([int]$task.priority/3.0) + 0.55*$urgency)
 return [Drawing.Color]::FromArgb([int](250-79*$score),[int](239-158*$score),[int](237-145*$score))
}
function Set-TaskRowStyle($row,$task) {
 $row.BackColor = Get-TaskColor $task
 $row.ForeColor = if ($task.done) { [Drawing.Color]::FromArgb(119,134,127) } elseif ($row.BackColor.GetBrightness() -lt 0.57) { [Drawing.Color]::FromArgb(255,249,246) } else { [Drawing.Color]::FromArgb(65,37,40) }
}
function Edit-TaskDetails($task) {
 $dialog = New-Object Windows.Forms.Form
 $dialog.Text = '任务设置'
 $dialog.Size = New-Object Drawing.Size(390,435)
 $dialog.StartPosition = 'CenterParent'
 $dialog.FormBorderStyle = 'FixedDialog'
 $dialog.MaximizeBox = $false
 $dialog.MinimizeBox = $false
 $dialog.Font = $form.Font
  $dialog.BackColor = $form.BackColor
 $contentLabel = New-Object Windows.Forms.Label
 $contentLabel.Text = '任务内容'
 $contentLabel.SetBounds(20,18,330,24)
 $dialog.Controls.Add($contentLabel)
 $content = New-Object Windows.Forms.TextBox
 $content.Text = [string]$task.text
 $content.SetBounds(20,47,330,30)
 $dialog.Controls.Add($content)
 $label = New-Object Windows.Forms.Label
 $label.Text = '截止日期（勾选启用）'
 $label.SetBounds(20,100,330,24)
 $dialog.Controls.Add($label)
 $date = New-Object Windows.Forms.DateTimePicker
 $date.SetBounds(20,130,330,30)
 $date.Format = 'Custom'
 $date.CustomFormat = 'yyyy年MM月dd日'
 $date.ShowCheckBox = $true
 $date.Checked = $false
 if ($task.deadline) { $date.Value = Get-DeadlineTime $task.deadline; $date.Checked = $true }
 $dialog.Controls.Add($date)
 $useTime = New-Object Windows.Forms.CheckBox
 $useTime.Text = '指定时间（可选）'
 $useTime.SetBounds(20,170,180,28)
 $dialog.Controls.Add($useTime)
 $time = New-Object Windows.Forms.DateTimePicker
 $time.Format = 'Custom'
 $time.CustomFormat = 'HH:mm'
 $time.ShowUpDown = $true
 $time.SetBounds(218,170,132,28)
 $time.Value = (Get-Date).AddHours(1)
 $useTime.Checked = [bool]($task.deadline -and $task.deadline.Length -gt 10)
 if ($useTime.Checked) { $time.Value = Get-DeadlineTime $task.deadline }
 $time.Enabled = $date.Checked -and $useTime.Checked
 $useTime.Enabled = $date.Checked
 $dialog.Controls.Add($time)
 $updateTimeState = { $useTime.Enabled = $date.Checked; $time.Enabled = $date.Checked -and $useTime.Checked }
 $date.Add_ValueChanged($updateTimeState)
 $useTime.Add_CheckedChanged($updateTimeState)
 $label2 = New-Object Windows.Forms.Label
 $label2.Text = '重要程度'
 $label2.SetBounds(20,216,330,24)
 $dialog.Controls.Add($label2)
 $priority = New-Object Windows.Forms.ComboBox
 $priority.DropDownStyle = 'DropDownList'
 $priority.SetBounds(20,245,330,30)
 [void]$priority.Items.AddRange([object[]]@('低','普通','重要','非常重要'))
 $priority.SelectedIndex = [Math]::Max(0,[Math]::Min(3,[int]$task.priority))
 $dialog.Controls.Add($priority)
 $ok = New-Object Windows.Forms.Button
 $ok.Text = '保存'
 $ok.SetBounds(175,329,80,34)
  $ok.Add_Click({
  if ([string]::IsNullOrWhiteSpace($content.Text)) {
   [void][Windows.Forms.MessageBox]::Show($dialog,'任务内容不能为空。','任务设置')
   $content.Focus()
   return
  }
  $dialog.DialogResult = 'OK'
 })
 $dialog.Controls.Add($ok)
 $cancel = New-Object Windows.Forms.Button
 $cancel.Text = '取消'
 $cancel.SetBounds(267,329,80,34)
 $cancel.DialogResult = 'Cancel'
 $dialog.Controls.Add($cancel)
 $dialog.AcceptButton = $ok
 $dialog.CancelButton = $cancel
 $accepted = $dialog.ShowDialog($form) -eq 'OK'
 if ($accepted) {
  $task.text = $content.Text.Trim()
    $task.deadline = if (-not $date.Checked) { $null } elseif ($useTime.Checked) {
   $date.Value.Date.AddHours($time.Value.Hour).AddMinutes($time.Value.Minute).ToString('yyyy-MM-dd HH:mm')
  } else { $date.Value.ToString('yyyy-MM-dd') }
  $task.priority = $priority.SelectedIndex
 }
 $dialog.Dispose()
 return $accepted
}
function Save-Tasks {
 $json = ConvertTo-Json -InputObject @($script:items) -Depth 4
 $tempPath = "$dataPath.tmp"
 [IO.File]::WriteAllText($tempPath,$json,(New-Object Text.UTF8Encoding($false)))
 if (Test-Path -LiteralPath $dataPath) { Copy-Item -LiteralPath $dataPath -Destination ($dataPath + '.bak') -Force }
 Move-Item -LiteralPath $tempPath -Destination $dataPath -Force
}
$form = New-Object Windows.Forms.Form
$form.Text = '今日任务'
$form.Icon = New-Object Drawing.Icon((Join-Path $PSScriptRoot 'note.ico'))
$form.Size = New-Object Drawing.Size(390,600)
$form.StartPosition = 'CenterScreen'

$form.BackColor = [Drawing.Color]::FromArgb(237,240,242)
$form.Font = New-Object Drawing.Font('KaiTi',10)
$form.FormBorderStyle = 'Sizable'
$form.Opacity = 0.88
$form.MaximizeBox = $true
$form.ShowInTaskbar = $true
$title = New-Object Windows.Forms.Label
$title.Text = '今日任务'
$title.Font = New-Object Drawing.Font('KaiTi',19,[Drawing.FontStyle]::Bold)
$title.SetBounds(20,18,280,40)
$form.Controls.Add($title)
$summary = New-Object Windows.Forms.Label
$summary.SetBounds(22,65,300,25)
$summary.ForeColor = [Drawing.Color]::DimGray
$form.Controls.Add($summary)
$list = New-Object Windows.Forms.ListView
$list.SetBounds(20,102,350,285)
$list.BackColor = $form.BackColor
$list.BorderStyle = 'None'
$list.CheckBoxes = $true
$list.View = 'Details'
$list.HeaderStyle = 'None'
$list.FullRowSelect = $true
$list.MultiSelect = $true
$list.HideSelection = $false
[void]$list.Columns.Add('任务',325)


$form.Controls.Add($list)
$inputBox = New-Object Windows.Forms.TextBox
$inputBox.SetBounds(20,400,270,28)
$form.Controls.Add($inputBox)
$add = New-Object Windows.Forms.Button
$add.Text = '添加'
$add.SetBounds(300,398,70,32)
$form.Controls.Add($add)
$remove = New-Object Windows.Forms.Button
$remove.Text = '删除选中'
$remove.SetBounds(20,446,105,32)
$form.Controls.Add($remove)
$hint = New-Object Windows.Forms.Label
$hint.Text = 'Ctrl / Shift 多选删除 · 勾选完成 · 自动保存'
$hint.Font = New-Object Drawing.Font('KaiTi',8)
$hint.ForeColor = [Drawing.Color]::DimGray
$hint.SetBounds(20,494,350,25)
$form.Controls.Add($hint)
function Update-Summary {
 $done = @($script:items | Where-Object { $_.done }).Count
 $summary.Text = (Get-Date -Format 'M月d日') + "   ·   已完成 $done / $($script:items.Count)"
}
function Get-SortedTasks($tasks,$mode) {
 $dueKey = { if ($_.deadline) { Get-DeadlineTime $_.deadline } else { [datetime]::MaxValue } }
 switch ($mode) {
  1 { $tasks | Sort-Object @{Expression=$dueKey},@{Expression={-[int]$_.priority}} }
  2 { $tasks | Sort-Object @{Expression={-[int]$_.priority}},@{Expression=$dueKey} }
  3 { $tasks | Sort-Object @{Expression={(Get-TaskColor $_).G}},@{Expression=$dueKey} }
  default { $tasks }
 }
}
function Refresh-Tasks {
 if ($null -ne $sortBox) { $script:items = @(Get-SortedTasks $script:items $sortBox.SelectedIndex) }
 $script:refreshing = $true
 $list.Items.Clear()
  foreach ($task in $script:items) {
  $priorityText = @('低','普通','重要','非常重要')[[int]$task.priority]
  $dueText = if ($task.deadline) { $task.deadline } else { '无截止日期' }
  $row = New-Object Windows.Forms.ListViewItem(([string]$task.text + '   · ' + $priorityText + ' · ' + $dueText))
  $row.Checked = [bool]$task.done
  Set-TaskRowStyle $row $task
  [void]$list.Items.Add($row)
 }
 $script:refreshing = $false
 Update-Summary
}
$addTask = {
 $value = $inputBox.Text.Trim()
 if ($value.Length -eq 0) { return }
 $newTask = [pscustomobject]@{text=$value;done=$false;deadline=$null;priority=1}
 if (-not (Edit-TaskDetails $newTask)) { return }
 $script:items += $newTask
 Save-Tasks
 Refresh-Tasks
 $inputBox.Clear()
 $inputBox.Focus()
}
$add.Add_Click($addTask)
$inputBox.Add_KeyDown({if($_.KeyCode -eq 'Enter') { & $addTask; $_.SuppressKeyPress=$true }})
$list.Add_ItemCheck({
 if (-not $script:refreshing) {
  $script:items[$_.Index].done = ($_.NewValue -eq [Windows.Forms.CheckState]::Checked)
  Save-Tasks
  Update-Summary
 }
})
$remove.Add_Click({
 $selected = @($list.SelectedIndices)
 if ($selected.Count -gt 0) {
  $script:items = @(for($i=0;$i -lt $script:items.Count;$i++){if($i -notin $selected){$script:items[$i]}})
  Save-Tasks
  Refresh-Tasks
 }
})
$pin = New-Object Windows.Forms.CheckBox
$pin.Text = '置顶'
$pin.SetBounds(250,27,65,28)
$pin.Add_CheckedChanged({ $form.TopMost = $pin.Checked })
$form.Controls.Add($pin)
$close = New-Object Windows.Forms.Button
$close.Text = '×'
$close.SetBounds(331,20,38,32)
$close.Add_Click({ $form.Close() })
$form.Controls.Add($close)
$clear = New-Object Windows.Forms.Button
$clear.Text = '清除已完成'
$clear.SetBounds(137,446,125,32)
$clear.Add_Click({
 $script:items = @($script:items | Where-Object { -not $_.done })
 Save-Tasks
 Refresh-Tasks
})
$form.Controls.Add($clear)
foreach ($button in @($add,$remove,$clear,$close)) {
 $button.FlatStyle = 'Flat'
 $button.FlatAppearance.BorderSize = 0
 $button.BackColor = [Drawing.Color]::FromArgb(210,216,226)
 $button.Cursor = 'Hand'
}
$add.BackColor = [Drawing.Color]::FromArgb(73,98,131)
$add.ForeColor = [Drawing.Color]::White
$title.Width = 215
$script:dragPoint = $null
$dragDown = { if ($_.Button -eq 'Left') { $script:dragPoint = [Windows.Forms.Cursor]::Position; $script:dragOrigin = $form.Location } }
$dragMove = { if ($null -ne $script:dragPoint) { $pos = [Windows.Forms.Cursor]::Position; $form.Location = New-Object Drawing.Point(($script:dragOrigin.X + $pos.X - $script:dragPoint.X),($script:dragOrigin.Y + $pos.Y - $script:dragPoint.Y)) } }
foreach ($surface in @($form,$title,$summary)) {
 $surface.Add_MouseDown($dragDown)
 $surface.Add_MouseMove($dragMove)
 $surface.Add_MouseUp({ $script:dragPoint = $null })
}
function Set-RoundedControl($control, $diameter) {
 $shape = New-Object Drawing.Drawing2D.GraphicsPath
 $w=$control.Width; $h=$control.Height
 $shape.AddArc(0,0,$diameter,$diameter,180,90)
 $shape.AddArc(($w-$diameter),0,$diameter,$diameter,270,90)
 $shape.AddArc(($w-$diameter),($h-$diameter),$diameter,$diameter,0,90)
 $shape.AddArc(0,($h-$diameter),$diameter,$diameter,90,90)
 $shape.CloseFigure()
 $control.Region = New-Object Drawing.Region($shape)
 $shape.Dispose()
}
$form.ForeColor = [Drawing.Color]::FromArgb(63,76,84)
$title.ForeColor = $form.ForeColor
$inputBox.BorderStyle = 'None'
$inputBox.BackColor = [Drawing.Color]::FromArgb(249,250,251)
$inputBox.SetBounds(32,409,242,23)
$inputSurface = New-Object Windows.Forms.Panel
$inputSurface.SetBounds(20,398,270,42)
$inputSurface.BackColor = $inputBox.BackColor
$form.Controls.Add($inputSurface)
$inputBox.Parent = $inputSurface
$inputBox.Location = New-Object Drawing.Point(12,10)
Set-RoundedControl $inputSurface 18
$add.SetBounds(300,398,70,42)
foreach ($button in @($add,$remove,$clear,$close)) {
 $button.UseVisualStyleBackColor = $false
 $button.BackColor = [Drawing.Color]::FromArgb(221,229,231)
 $button.ForeColor = [Drawing.Color]::FromArgb(67,86,91)
 $button.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(202,218,219)
 $button.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(185,205,207)
 Set-RoundedControl $button 16
}
$add.BackColor = [Drawing.Color]::FromArgb(112,148,151)
$add.ForeColor = [Drawing.Color]::White
$close.BackColor = $form.BackColor
$summary.ForeColor = [Drawing.Color]::FromArgb(112,125,132)
$hint.ForeColor = $summary.ForeColor
$opacityLabel = New-Object Windows.Forms.Label
$opacityLabel.SetBounds(22,538,135,25)
$form.Controls.Add($opacityLabel)
$opacitySlider = New-Object Windows.Forms.TrackBar
$opacitySlider.SetBounds(157,530,211,42)
$opacitySlider.Minimum = 0
$opacitySlider.Maximum = 55
$opacitySlider.Value = 12
$opacitySlider.TickStyle = 'None'
$opacitySlider.SmallChange = 1
$opacitySlider.LargeChange = 5
$opacitySlider.BackColor = $form.BackColor
$form.Controls.Add($opacitySlider)
$settingsPath = Join-Path $PSScriptRoot 'appearance.json'
if (Test-Path -LiteralPath $settingsPath) {
 try {
  $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
  $opacitySlider.Value = [Math]::Max(0,[Math]::Min(55,[int]$settings.transparency))
 } catch { }
}
$applyOpacity = {
 $form.Opacity = (100 - $opacitySlider.Value) / 100.0
 $opacityLabel.Text = "透明度  $($opacitySlider.Value)%"
}
$opacitySlider.Add_ValueChanged($applyOpacity)
$opacitySlider.Add_MouseUp({
 [IO.File]::WriteAllText($settingsPath,('{"transparency":' + $opacitySlider.Value + '}'))
})
$form.Add_FormClosing({
 [IO.File]::WriteAllText($settingsPath,('{"transparency":' + $opacitySlider.Value + '}'))
})
& $applyOpacity
$form.MinimumSize = New-Object Drawing.Size(390,440)
$list.Anchor = 'Top,Bottom,Left,Right'
$inputSurface.Anchor = 'Bottom,Left,Right'
$inputBox.Anchor = 'Top,Left,Right'
$add.Anchor = 'Bottom,Right'
$remove.Anchor = 'Bottom,Left'
$clear.Anchor = 'Bottom,Left'
$hint.Anchor = 'Bottom,Left'
$opacityLabel.Anchor = 'Bottom,Left'
$opacitySlider.Anchor = 'Bottom,Left,Right'
$pin.Anchor = 'Top,Right'
$close.Anchor = 'Top,Right'
$windowStatePath = Join-Path $PSScriptRoot 'window.json'
$saveWindow = {
 $state = @{width=$form.Width;height=$form.Height} | ConvertTo-Json
 [IO.File]::WriteAllText($windowStatePath,$state)
}
$form.Add_ResizeEnd($saveWindow)
$form.Add_FormClosing($saveWindow)
$form.Add_SizeChanged({

 Set-RoundedControl $inputSurface 18
 $list.Columns[0].Width = [Math]::Max(100,$list.ClientSize.Width - 24)
})
if (Test-Path -LiteralPath $windowStatePath) {
 try {
  $state = Get-Content -LiteralPath $windowStatePath -Raw | ConvertFrom-Json
  $area = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
  $form.Size = New-Object Drawing.Size(([Math]::Max(390,[Math]::Min($area.Width,[int]$state.width))),([Math]::Max(440,[Math]::Min($area.Height,[int]$state.height))))
 } catch { }
}
$edit = New-Object Windows.Forms.Button
$edit.Text = '编辑任务'
$edit.SetBounds(274,446,96,32)
$edit.Anchor = 'Bottom,Right'
$edit.FlatStyle = 'Flat'
$edit.FlatAppearance.BorderSize = 0
$edit.BackColor = [Drawing.Color]::FromArgb(221,229,231)
$edit.ForeColor = [Drawing.Color]::FromArgb(67,86,91)
Set-RoundedControl $edit 16
$form.Controls.Add($edit)
$editAction = {
 if ($list.SelectedIndices.Count -ne 1) { [void][Windows.Forms.MessageBox]::Show($form,'请先选择一条任务。','任务设置'); return }
 $index = $list.SelectedIndices[0]
 if (Edit-TaskDetails $script:items[$index]) { Save-Tasks; Refresh-Tasks }
}
$edit.Add_Click($editAction)
$list.Add_DoubleClick($editAction)
$hint.Text = '双击编辑内容、DDL、重要度 · Ctrl 多选'
$colorTimer = New-Object Windows.Forms.Timer
$colorTimer.Interval = 60000
$colorTimer.Add_Tick({
 for($i=0;$i -lt $list.Items.Count;$i++) { Set-TaskRowStyle $list.Items[$i] $script:items[$i] }
 Update-Summary
})
$colorTimer.Start()
$sortBox = New-Object Windows.Forms.ComboBox
$sortBox.DropDownStyle = 'DropDownList'
$sortBox.FlatStyle = 'Flat'
$sortBox.BackColor = [Drawing.Color]::FromArgb(249,250,251)
[void]$sortBox.Items.AddRange([object[]]@('当前顺序','DDL：最近优先','重要度：高到低','底色：深红优先'))
$sortBox.SelectedIndex = 0
$sortPath = Join-Path $PSScriptRoot 'sorting.json'
if (Test-Path -LiteralPath $sortPath) {
 try { $sortBox.SelectedIndex = [Math]::Max(0,[Math]::Min(3,[int]([IO.File]::ReadAllText($sortPath)))) } catch { }
}
$form.Controls.Add($sortBox)
$sortBox.Add_SelectedIndexChanged({
 Refresh-Tasks
 [IO.File]::WriteAllText($sortPath,[string]$sortBox.SelectedIndex)
})
function Update-Layout {
 $w = $form.ClientSize.Width
 $h = $form.ClientSize.Height
 $inputY = $h - 192
 $summary.SetBounds(22,65,($w-44),25)
 $sortBox.SetBounds(20,96,($w-40),28)
 $list.SetBounds(20,136,($w-40),([Math]::Max(80,$inputY-150)))
 $inputSurface.SetBounds(20,$inputY,($w-130),42)
 $inputBox.SetBounds(12,10,($inputSurface.Width-24),23)
 $add.SetBounds(($w-96),$inputY,76,42)
 $buttonY = $inputY + 58
 $gap = 10
 $buttonWidth = [int](($w-60)/3)
 $remove.SetBounds(20,$buttonY,$buttonWidth,34)
 $clear.SetBounds((30+$buttonWidth),$buttonY,$buttonWidth,34)
 $edit.SetBounds((40+2*$buttonWidth),$buttonY,$buttonWidth,34)
 $hint.SetBounds(20,($buttonY+44),($w-40),22)
 $opacityLabel.SetBounds(22,($h-48),125,25)
 $opacitySlider.SetBounds(150,($h-56),($w-170),42)
 $close.SetBounds(($w-55),20,35,32)
 $pin.SetBounds(($w-135),25,70,28)
 $list.Columns[0].Width = [Math]::Max(100,$list.ClientSize.Width-24)
 Set-RoundedControl $inputSurface 18
 foreach($button in @($add,$remove,$clear,$edit,$close)){Set-RoundedControl $button 16}
}
$form.MinimumSize = New-Object Drawing.Size(430,550)
$form.Add_SizeChanged({ Update-Layout })
Update-Layout
$form.Add_Shown({
 $form.Activate()
 $form.BringToFront()
 Add-Content -LiteralPath $log -Value 'WINDOW_SHOWN'
})
Refresh-Tasks
Add-Content -LiteralPath (Join-Path $PSScriptRoot 'startup.log') -Value 'ENTER_MESSAGE_LOOP'
[Windows.Forms.Application]::Run($form)
} catch {
 ($_ | Out-String) + $_.ScriptStackTrace | Add-Content -LiteralPath $log
 throw
}